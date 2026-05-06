"""
CorretorAMF — Backend FastAPI
Verifica conformidade de trabalhos acadêmicos com as normas da AMF via Google Gemini.
"""

import io
import json
import logging
import os
import re
import time
from contextlib import asynccontextmanager
from pathlib import Path

import google.generativeai as genai
from docx import Document
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from prompts import SYSTEM_PROMPT, build_file_prompt, build_user_prompt

load_dotenv()

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
MODEL_NAME = "gemini-3-flash-preview"
MAX_FILE_SIZE_MB = 10
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024
ALLOWED_MIME_TYPES = {
    "application/pdf": ".pdf",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
}


@asynccontextmanager
async def lifespan(app: FastAPI):
    if not GEMINI_API_KEY:
        logger.warning("GEMINI_API_KEY not set — /analyze will fail at runtime")
    else:
        genai.configure(api_key=GEMINI_API_KEY)
        logger.info("Gemini SDK configured with model: %s", MODEL_NAME)
    yield


app = FastAPI(
    title="CorretorAMF",
    description="Verificador de conformidade de trabalhos acadêmicos com normas AMF via IA",
    version="1.0.0",
    lifespan=lifespan,
)

static_dir = Path(__file__).parent / "static"
app.mount("/static", StaticFiles(directory=str(static_dir)), name="static")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _extract_docx_text(file_bytes: bytes) -> str:
    """Extrai texto de um arquivo DOCX usando python-docx."""
    doc = Document(io.BytesIO(file_bytes))
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    # Inclui texto de tabelas também
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip():
                    paragraphs.append(cell.text.strip())
    return "\n".join(paragraphs)


def _parse_gemini_response(raw: str) -> dict:
    """
    Extrai JSON da resposta do Gemini.
    O modelo às vezes envolve o JSON em markdown mesmo sendo instruído a não fazer isso —
    este parser tenta recuperar o JSON puro em qualquer formato de resposta.
    """
    text = raw.strip()

    # Remover bloco de código markdown se presente
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
        text = text.strip()

    # Tentar parse direto
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Tentar encontrar o primeiro objeto JSON no texto
    match = re.search(r"\{.*\}", text, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    raise ValueError(f"Não foi possível extrair JSON válido da resposta do modelo: {raw[:200]}")


def _call_gemini_with_text(document_text: str, extra_context: str) -> dict:
    """
    Chama o Gemini usando o texto extraído do documento (fluxo DOCX).
    O conteúdo é passado diretamente no prompt de texto.
    """
    model = genai.GenerativeModel(
        model_name=MODEL_NAME,
        system_instruction=SYSTEM_PROMPT,
    )
    user_prompt = build_user_prompt(document_text, extra_context)
    response = model.generate_content(
        user_prompt,
        generation_config=genai.GenerationConfig(
            temperature=0.1,   # baixa temperatura para avaliação determinística
            max_output_tokens=4096,
        ),
    )
    return _parse_gemini_response(response.text)


def _call_gemini_with_file(file_bytes: bytes, mime_type: str, extra_context: str) -> dict:
    """
    Chama o Gemini usando a File API para envio multimodal do PDF.
    O arquivo é enviado, processado e depois deletado para não acumular na conta.
    """
    model = genai.GenerativeModel(
        model_name=MODEL_NAME,
        system_instruction=SYSTEM_PROMPT,
    )

    # Upload via File API — o Gemini processa o PDF diretamente (OCR + estrutura)
    uploaded_file = genai.upload_file(
        path=io.BytesIO(file_bytes),
        mime_type=mime_type,
        display_name="trabalho_academico",
    )

    # Aguardar o arquivo ficar disponível (estado ACTIVE)
    max_wait = 30
    waited = 0
    while uploaded_file.state.name == "PROCESSING" and waited < max_wait:
        time.sleep(2)
        waited += 2
        uploaded_file = genai.get_file(uploaded_file.name)

    if uploaded_file.state.name != "ACTIVE":
        raise RuntimeError(f"Arquivo não ficou ativo após {max_wait}s: {uploaded_file.state.name}")

    try:
        user_prompt = build_file_prompt(extra_context)
        response = model.generate_content(
            [uploaded_file, user_prompt],
            generation_config=genai.GenerationConfig(
                temperature=0.1,
                max_output_tokens=4096,
            ),
        )
        return _parse_gemini_response(response.text)
    finally:
        # Limpar o arquivo do Gemini imediatamente após o uso
        try:
            genai.delete_file(uploaded_file.name)
        except Exception as e:
            logger.warning("Falha ao deletar arquivo do Gemini: %s", e)


# ---------------------------------------------------------------------------
# Rotas
# ---------------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def serve_frontend():
    """Serve o frontend principal."""
    index_path = static_dir / "index.html"
    return HTMLResponse(content=index_path.read_text(encoding="utf-8"))


@app.get("/health")
async def health_check():
    """Verifica se o serviço está operacional."""
    return {"status": "ok", "model": MODEL_NAME}


@app.post("/analyze")
async def analyze_document(
    file: UploadFile = File(..., description="PDF ou DOCX do trabalho acadêmico"),
    context: str = Form(default="", description="Contexto adicional opcional informado pelo aluno"),
):
    """
    Analisa um trabalho acadêmico e retorna um relatório de conformidade com as normas AMF.

    - **file**: arquivo PDF ou DOCX, máximo 10 MB
    - **context**: informações extras sobre o trabalho (tipo, disciplina, restrições aplicáveis)
    """
    if not GEMINI_API_KEY:
        raise HTTPException(
            status_code=500,
            detail="Serviço não configurado: GEMINI_API_KEY ausente. Contate o administrador.",
        )

    # Validar tipo de arquivo
    content_type = file.content_type or ""
    # Normalizar content_type que às vezes vem com charset
    content_type_base = content_type.split(";")[0].strip()

    # Aceitar também por extensão de nome de arquivo como fallback
    filename = file.filename or ""
    if content_type_base not in ALLOWED_MIME_TYPES:
        if filename.lower().endswith(".pdf"):
            content_type_base = "application/pdf"
        elif filename.lower().endswith(".docx"):
            content_type_base = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            raise HTTPException(
                status_code=400,
                detail=f"Tipo de arquivo não suportado: '{content_type}'. Envie um PDF ou DOCX.",
            )

    # Ler arquivo em memória
    file_bytes = await file.read()

    if len(file_bytes) == 0:
        raise HTTPException(status_code=400, detail="Arquivo enviado está vazio.")

    if len(file_bytes) > MAX_FILE_SIZE_BYTES:
        size_mb = len(file_bytes) / (1024 * 1024)
        raise HTTPException(
            status_code=413,
            detail=f"Arquivo muito grande: {size_mb:.1f} MB. Limite: {MAX_FILE_SIZE_MB} MB.",
        )

    logger.info(
        "Analisando arquivo: %s | tipo: %s | tamanho: %.1f KB",
        filename,
        content_type_base,
        len(file_bytes) / 1024,
    )

    try:
        if content_type_base == "application/pdf":
            # PDF: enviar diretamente ao Gemini via File API (processamento multimodal)
            result = _call_gemini_with_file(file_bytes, content_type_base, context)
        else:
            # DOCX: extrair texto com python-docx e enviar como texto plano
            document_text = _extract_docx_text(file_bytes)
            if not document_text.strip():
                raise HTTPException(
                    status_code=422,
                    detail="Não foi possível extrair texto do DOCX. O arquivo pode estar corrompido.",
                )
            result = _call_gemini_with_text(document_text, context)

    except HTTPException:
        raise
    except Exception as e:
        logger.exception("Erro ao processar documento")
        raise HTTPException(
            status_code=500,
            detail=f"Erro ao analisar o documento: {str(e)}. Tente novamente.",
        )

    # Garantir campos mínimos na resposta mesmo se o modelo retornar algo inesperado
    result.setdefault("score", "0%")
    result.setdefault("criteria", [])
    result.setdefault("suggestions", [])

    return JSONResponse(content=result)

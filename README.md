# CorretorAMF

Verificador automático de conformidade de trabalhos acadêmicos com as normas oficiais da **Antonio Meneghetti Faculdade (AMF)**.

> Projeto prático — Disciplina de Inteligência Artificial · AMF 2026/01

**Stack:** Python · FastAPI · Google Gemini · HTML/CSS/JS puro

---

## O que faz

O aluno envia seu trabalho (PDF ou DOCX) e recebe em segundos um relatório com:

- **Score de conformidade** — porcentagem geral
- **Critérios avaliados** — cada norma com status ✅ Conforme / ⚠️ Atenção / ❌ Não conforme
- **Sugestões priorizadas** — o que corrigir antes de entregar ao professor

O sistema identifica automaticamente o **tipo de documento** (TCC, pré-projeto, artigo, seminário...) e aplica apenas as regras pertinentes a cada tipo — evitando punir por elementos que não são obrigatórios naquele contexto.

---

## Pré-requisitos

- Python **3.11+**
- Chave de API do **Google Gemini** (gratuita no Google AI Studio)
- Conexão com a internet

---

## Instalação

```bash
# 1. Clone o repositório
git clone <url-do-repo>
cd corretor-amf

# 2. Crie e ative o ambiente virtual
python -m venv .venv

# Windows (Git Bash / PowerShell)
source .venv/Scripts/activate    # Git Bash
.venv\Scripts\activate           # PowerShell

# Linux / macOS
source .venv/bin/activate

# 3. Instale as dependências
pip install -r requirements.txt

# 4. Configure a chave de API
cp .env.example .env
# Edite .env e preencha GEMINI_API_KEY
```

---

## Como obter a GEMINI_API_KEY

1. Acesse [Google AI Studio](https://aistudio.google.com/app/apikey)
2. Clique em **"Create API key"**
3. Copie a chave e cole no `.env`:

```env
GEMINI_API_KEY=AIzaSy...sua_chave_aqui
```

---

## Execução

```bash
# Desenvolvimento (reinicia automaticamente ao salvar)
.venv/Scripts/uvicorn main:app --reload --port 8000   # Windows/Git Bash
uvicorn main:app --reload --port 8000                  # Linux/macOS (com venv ativo)
```

Acesse: **http://localhost:8000**

---

## Exemplos via curl

```bash
# Analisar PDF com contexto
curl -X POST http://localhost:8000/analyze \
  -F "file=@pre_projeto.pdf" \
  -F "context=Pré-projeto de TCC, sem folha de rosto ou sumário"

# Analisar DOCX
curl -X POST http://localhost:8000/analyze \
  -F "file=@tcc_final.docx"

# Health check
curl http://localhost:8000/health
```

### Resposta exemplo

```json
{
  "score": "78%",
  "criteria": [
    {
      "name": "Tipo de documento identificado",
      "status": "pass",
      "detail": "Pré-projeto de TCC — avaliação restrita aos elementos exigidos para este tipo."
    },
    {
      "name": "Fonte e tipografia",
      "status": "pass",
      "detail": "Times New Roman 12 pt identificado, conforme norma AMF."
    },
    {
      "name": "Referências (ABNT NBR 6023)",
      "status": "fail",
      "detail": "Faltam dados de edição e local de publicação em 2 referências."
    },
    {
      "name": "Margens",
      "status": "warn",
      "detail": "Não foi possível verificar as margens no conteúdo extraído do documento."
    }
  ],
  "suggestions": [
    "Corrija as referências para seguir ABNT NBR 6023:2025 (local, editora, ano obrigatórios).",
    "Confirme as margens: superior e esquerda 3 cm; inferior e direita 2 cm."
  ]
}
```

---

## Arquitetura de prompt

### Por que as normas ficam no system prompt?

As normas da AMF são embutidas no **system prompt** — não no turno do usuário — por duas razões principais:

1. **Autoridade imutável.** O modelo trata o system prompt como instrução do sistema, dificultando que conteúdo do documento enviado "sobrescreva" as regras de avaliação.
2. **Guardrail contra alucinação.** A instrução `"Avalie APENAS com base nas normas acima"` no contexto de sistema tem maior peso semântico do que no turno do usuário.

### Identificação do tipo de documento

Antes de avaliar qualquer critério, o modelo executa um **Passo 0** obrigatório: classificar o tipo do documento (TCC completo, pré-projeto, artigo, seminário, relatório...). Só então seleciona os critérios aplicáveis. Isso evita o erro clássico de punir um pré-projeto por não ter folha de aprovação.

O campo **"Contexto adicional"** do frontend tem prioridade máxima: se o aluno escrever `"Pré-projeto de TCC"`, isso sobrepõe qualquer inferência do modelo.

### Estratégia conservadora: warn antes de fail

Critérios que não podem ser verificados no texto extraído (ex.: margens, espaçamento exato) são marcados como `warn` — nunca `fail`. Isso evita falsos negativos que prejudicariam injustamente o aluno por limitações técnicas da extração.

### Fluxo por tipo de arquivo

| Tipo | Estratégia |
|------|-----------|
| **PDF** | Gemini File API → processamento multimodal (lê layout + texto diretamente) |
| **DOCX** | `python-docx` extrai o texto → enviado como texto plano no prompt |

---

## Estrutura de pastas

```
corretor-amf/
├── main.py          # FastAPI: rotas /analyze, /health e /; validação e orquestração
├── prompts.py       # System prompt com normas injetadas; build_user_prompt(); build_file_prompt()
├── normas.py        # NORMAS_AMF: conteúdo extraído do Manual AMF 3ª Ed. 2025 (Biblioteca Humanitas)
├── static/
│   └── index.html   # Frontend completo: drag & drop, loading animado, relatório com cards
├── .env             # GEMINI_API_KEY — não versionado (.gitignore)
├── .env.example     # Template seguro para compartilhar
├── requirements.txt # Dependências com versões fixadas
└── README.md
```

---

## Dependências

| Pacote | Versão | Função |
|--------|--------|--------|
| `fastapi` | 0.115.5 | Framework web |
| `uvicorn[standard]` | 0.32.1 | Servidor ASGI |
| `python-multipart` | 0.0.19 | Parse de multipart/form-data |
| `google-generativeai` | 0.8.3 | SDK oficial do Gemini (File API + generate) |
| `python-docx` | 1.1.2 | Extração de texto de arquivos DOCX |
| `python-dotenv` | 1.0.1 | Carregamento do `.env` |

---

## Limitações conhecidas

- **Margens, espaçamento e fonte** não são verificáveis com precisão em texto extraído — marcados como `warn`.
- Documentos muito grandes (> ~100 páginas) podem exceder o contexto do modelo.
- O PDF do Manual AMF não é versionado no repositório por razões de direitos autorais; o conteúdo normativo foi transcrito em `normas.py`.

---

## Equipe

Desenvolvido como trabalho prático da disciplina de **Inteligência Artificial** — AMF 2026/01.

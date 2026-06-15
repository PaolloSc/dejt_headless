# DEJT Headless — Diário da Justiça do Trabalho (TST)

[![Python](https://img.shields.io/badge/Python-3.10+-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-0078D4?style=flat-square&logo=microsoft&logoColor=white)](https://learn.microsoft.com/graph)

Robô **headless** (sem GUI) que coleta automaticamente o Diário Eletrônico da Justiça do Trabalho do TST, gera os arquivos em PDF e DOCX e envia por e-mail via Microsoft Graph API. Roda em qualquer SO (Linux/VPS, Windows, macOS).

## O que faz

- Coleta o caderno do dia direto da fonte (requests + BeautifulSoup/lxml)
- Gera `Diario_J_TST_AAAA-MM-DD.pdf` e o DOCX correspondente com sumário (TOC) montado em python-docx puro — sem dependência de Word/COM
- Envia o documento por e-mail via Microsoft Graph
- Agendamento semanal (sexta) via scripts `.bat`

## Saídas

```
Diario_J_TST_AAAA-MM-DD.pdf
Diario_J_TST_com_variaveis_AAAA-MM-DD.docx
logs/dejt_headless.log
```

## Instalação

```bash
pip install requests beautifulsoup4 lxml pdf2docx python-docx python-dotenv
```

## Configuração

Variáveis de ambiente (ou arquivo `.env` na mesma pasta):

| Variável | Descrição |
|----------|-----------|
| `AZURE_TENANT_ID` | Tenant do app registrado no Azure AD |
| `AZURE_CLIENT_ID` | Client ID da aplicação |
| `AZURE_CLIENT_SECRET` | Client secret |
| `GRAPH_SENDER_EMAIL` | E-mail remetente |
| `GRAPH_RECIPIENT_EMAIL` | E-mail destinatário |
| `USE_ADOBE_PDF_SERVICES` | `0` força conversão via pdf2docx (padrão: tenta Adobe primeiro) |

## Uso

```bash
python dejt_tst_daily_headless.py      # execução diária
python executar_sexta.py               # rotina semanal (sexta)
```

No Windows, `agendar_sexta.bat` registra a tarefa no Agendador; `executar_sexta.bat` dispara manualmente.

# Planilha Programar — migração para sistema web

Repositório da migração da planilha de **planejamento e programação de trançadeiras**
(Excel com macros, ~36 MB) para um sistema web da empresa.

## Situação atual

A planilha não cabe no Git. Em vez de versionar o arquivo, versionamos **o que ele
contém**: as fórmulas, a estrutura e as regras, extraídas em arquivos de texto.

## Onde começar

| Se você quer… | Leia |
|---|---|
| Exportar a planilha para o Git | [`docs/COMO-EXPORTAR.md`](docs/COMO-EXPORTAR.md) |
| Entender o que as macros fazem | [`docs/ENTENDIMENTO-VBA.md`](docs/ENTENDIMENTO-VBA.md) |
| Saber para onde o sistema vai | [`docs/ARQUITETURA-ALVO.md`](docs/ARQUITETURA-ALVO.md) |

## Estrutura

```
├── Módulo1.bas … Módulo18.bas   Macros atuais da planilha (referência)
├── ferramentas/
│   └── ExportarPlanilha.bas     Macro que extrai a planilha para texto
└── docs/
    ├── COMO-EXPORTAR.md         Guia passo a passo
    ├── ENTENDIMENTO-VBA.md      Análise dos 18 módulos
    ├── ARQUITETURA-ALVO.md      Decisões e perguntas em aberto
    └── planilha/                ← destino dos arquivos exportados
```

## Próximo passo

Rodar `ExportarTudo` e commitar o resultado em `docs/planilha/`. Sem isso, as fórmulas —
que são a lógica real do sistema — continuam invisíveis para quem não abre o arquivo.

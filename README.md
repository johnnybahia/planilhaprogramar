# Planilha Programar — migração para sistema web

Repositório da migração da planilha de **planejamento e programação de trançadeiras**
(Excel com macros, ~36 MB) para um sistema web da empresa.

## Situação atual

A planilha não cabe no Git. Em vez de versionar o arquivo, versionamos **o que ele
contém**: as fórmulas, a estrutura e as regras, extraídas em arquivos de texto.

## Onde começar

| Se você quer… | Leia |
|---|---|
| **Entender como a planilha funciona** | [`docs/ANALISE-PLANILHA.md`](docs/ANALISE-PLANILHA.md) ← comece aqui |
| Como o sistema funciona hoje, em detalhe | [`docs/COMO-FUNCIONA.md`](docs/COMO-FUNCIONA.md) |
| O que vamos construir | [`docs/REQUISITOS.md`](docs/REQUISITOS.md) |
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

## Situação

A exportação foi feita: **4.733.724 fórmulas** da planilha estão descritas em 1,5 MB de
texto neste repositório (arquivos `00_*` a `99_*`). A leitura completa está em
[`docs/ANALISE-PLANILHA.md`](docs/ANALISE-PLANILHA.md).

Três pontos que a análise encontrou:

- **`PEDIDOS` responde por 89% de todas as fórmulas** e pelos 36 MB do arquivo — quatro
  colunas preenchidas até a linha 1.048.576.
- **As 151.385 fórmulas de `Planejamento Trançadeira` calculam um resto de divisão.**
  Verificado numericamente contra os dados reais.
- **A aba `Relação de referencias` está 100% quebrada** (`=#REF!`) e alimenta o relatório
  de teares através de `CONSUMO`, mascarada por `SEERRO`.

## Próximo passo

Responder as perguntas em aberto no fim da análise e então modelar os dados.

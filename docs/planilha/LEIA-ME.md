# Destino dos arquivos exportados

Esta pasta está **vazia de propósito**. É aqui que vão os arquivos gerados pela macro
`ferramentas/ExportarPlanilha.bas`.

## O que fazer

1. Siga o guia em [`../COMO-EXPORTAR.md`](../COMO-EXPORTAR.md).
2. Copie **todo o conteúdo** da pasta `_exportacao` para dentro desta pasta.
3. Confira os arquivos `50_AMOSTRA_*.csv` antes do commit — eles contêm dados reais.
4. Commit e push.

Depois disso a pasta deve conter algo parecido com:

```
docs/planilha/
├── 00_RESUMO.md
├── 10_FORMULAS_01_RELATORIO_TRANCADEIRAS.txt
├── 10_FORMULAS_02_Planejamento_Trancadeira.txt
├── 10_FORMULAS_03_Programacao_trancadeiras.txt
├── 20_NOMES.csv
├── 30_VALIDACOES.csv
├── 40_FORMATACAO_CONDICIONAL.csv
├── 50_AMOSTRA_01_RELATORIO_TRANCADEIRAS.csv
├── ...
├── 60_LIGACOES.txt
├── 70_OBJETOS.txt
├── 80_ESTRUTURA.txt
└── 99_ARQUIVOS.txt
```

Os nomes dos arquivos são sem acento de propósito, para não dar problema entre Windows,
Linux e macOS. O nome real de cada aba, com acentuação correta, está escrito dentro do
arquivo.

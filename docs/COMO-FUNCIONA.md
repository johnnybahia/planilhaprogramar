# Como o sistema funciona hoje

Documento construído a partir da explicação do usuário, **conferido contra os dados
exportados**. Cresce por partes.

---

# Parte 1 — Entrada do pedido e cadastros

## 1.1 De onde vêm os dados

O usuário **copia de outra planilha** (o sistema de origem, "Marfim") e cola em
`PEDIDOS`, no intervalo **A4:I4** para baixo. A origem já entrega essas colunas prontas
e na ordem certa — por isso a cópia é feita em bloco, mesmo que nem tudo seja útil.

| Col | Na origem | Em `PEDIDOS` | O que é |
|---|---|---|---|
| A | CÓD. MARFIM | CÓDIGO | Código do produto no sistema de origem |
| B | QTD. ABERTA | QUANTIDADE | Quantidade **em pares ou peças, dependendo do cliente** |
| C | ORD. COMPRA | OC | Número do pedido de compra |
| D | DATA RECEB. | Entrada | Data de entrada do pedido |
| E | DT. ENTREGA | Entrega | Data de entrega combinada |
| F | PRAZO | Prazo | Ver **1.5** — não é o que parece |
| G | DESCRIÇÃO | Descrição | O item |
| H | TAMANHO | Tamanho | `MM` = largura (produto contínuo) · `CM` = comprimento do corte |
| I | CÓD. FILIAL | CÓD. CLIENTE | Ver **1.6** — o nome muda no caminho |

A partir daí, tudo o que a planilha faz é derivado dessas nove colunas.

## 1.2 Dois tipos de produto, duas unidades

- **Produtos contínuos** — fitas e elásticos, vendidos por metro. O `TAMANHO` é a
  **largura** (`8MM`, `15MM`).
- **Produtos cortados** — vendidos por peça já no comprimento. O `TAMANHO` é o
  **comprimento do corte** (`110CM`, `125CM`).

Um produto cortado carrega as duas medidas: o código diz `92505125CM` (corte de 1,25 m)
e a descrição diz `ATAC M6034C 2441 8MM` (largura de 8 mm).

## 1.3 O botão "1 - criar código"

O sistema de origem usa **o mesmo código para o mesmo item em tamanhos diferentes** —
100 cm e 120 cm compartilham código. Para a produção isso não serve: são itens distintos.

O botão executa `JuntarDados` (Módulo2), que faz `A = A & H` — concatena o tamanho ao
código, criando a chave única:

```
CÓD. MARFIM   TAMANHO      código único
   16      +   15MM     →    1615MM
 69776     +   15MM     →    6977615MM
 92505     +  125CM     →    92505125CM
```

Confirmado nos dados exportados: todos os códigos de `PEDIDOS` terminam exatamente com o
conteúdo da coluna H.

> ⚠️ **A macro não é idempotente.** Rodar duas vezes na mesma linha gera `1615MM15MM`.
> Não há trava. No sistema novo o código único é **derivado**, nunca gravado — some o
> problema e o botão.

## 1.4 Os cadastros

### `DADOS GERAIS DE PRODUTOS` — o cadastro de itens

Toda chave única criada no passo anterior precisa existir aqui. É esta aba que responde
três perguntas sobre o item:

| Col | Conteúdo | Alimenta |
|---|---|---|
| B | Código único (chave de busca) | — |
| C | Referência / descrição normalizada | `PEDIDOS!J` |
| D | **Tamanho em metros** (0,5 · 1,25 · 1,4) | `PEDIDOS!L` |
| E | **Máquina**: TEAR · TRANÇADEIRA · ELÁSTICO | `PEDIDOS!K` |

As colunas J, K e L de `PEDIDOS` são três `PROCV` nesta aba. É o único ponto que decide
para onde o pedido vai.

### `PESOS DE FIOS` — o peso do fio por cor

O produto é feito de fios. Cada cor tem uma quantidade na ficha técnica e um peso.
Registrar isso é o que permite ao relatório responder **quanto de fio cada produto
consome** — um dos objetivos centrais do sistema.

Layout: uma linha por referência, com **até 5 pares cor/peso** lado a lado.

```
REFERENCIAS                          cor              Peso     cor2         peso3    cor4         peso5
ATAC M21020 821/1P DARK KHAKI 8MM    821/1 RECICLADO  0,0031
ATAC M10046 100 6MM                  100              0,0028   ENCHIMENTO   0,00087
ATAC 6ALF 100/460 6MM                100              0,0009   460          0,0022   ENCHIMENTO   0,00087
```

`ENCHIMENTO` aparece como uma "cor" — é o fio de enchimento, contabilizado junto.

> Os cabeçalhos estão numerados de forma inconsistente (`cor`, `cor2`, `cor4`, `cor6`,
> `cor8` / `Peso`, `peso3`, `peso5`, `peso7`). No modelo novo isso vira uma tabela
> filha — uma linha por cor —, sem limite de 5 e sem numeração.

### `DADOS DOS PRODUTOS` — a ficha técnica

Complementa com quantidade de espulas, quantidade de fio na espula e cor, por produto.
Aqui o cabeçalho da planilha está **deslocado uma coluna à direita** em relação aos
dados: o rótulo `QUANTIDADE DE FIO NA ESPULA` está sobre a coluna da cor.

## 1.5 ⚠️ O que `PRAZO` realmente é

Você descreveu `PRAZO` como "quantos dias da data de recebimento à data de entrega".
**Os dados não confirmam isso.** Testei as 57 linhas da amostra:

| Recebimento | Entrega | Entrega − Recebimento | `PRAZO` gravado |
|---|---|---|---|
| 2025-09-18 | 2025-09-30 | 12 | 6 |
| 2025-09-23 | 2025-10-03 | 10 | 9 |
| 2025-08-28 | 2025-09-29 | 32 | 5 |
| 2025-09-09 | 2025-10-06 | 27 | 12 |

Não há relação. Mas subtraindo `PRAZO` da data de entrega, **as 57 linhas caem na mesma
data**: `2025-09-24`.

```
PRAZO = data de entrega − 24/09/2025
```

Ou seja: **`PRAZO` são os dias que faltavam para a entrega no dia em que a origem foi
extraída** — não a janela entre recebimento e entrega. Por isso aparecem valores
negativos (`-7`, `-8`): são pedidos que já venceram.

**A consequência prática:** o `PRAZO` é um retrato congelado. Ele não se atualiza. Se a
cópia foi feita há dez dias, todo prazo na planilha está dez dias otimista, e o
sequenciamento da produção está sendo decidido com número velho.

No sistema novo isso deve ser **calculado ao vivo** (`entrega − hoje`) e nunca copiado.
A coluna some da importação e vira um campo derivado, sempre correto.

## 1.6 Outros pontos observados

- **`CÓD. FILIAL` → `CÓD. CLIENTE`.** A origem chama de filial, `PEDIDOS` chama de
  cliente, e `CADASTRO CLIENTES` traduz o número para o nome (2 = DASS CEARÁ,
  53 = GRENDENE, 225 = DILLY). Provavelmente é a filial do cliente. Precisa confirmar,
  porque define se o cadastro é de clientes ou de locais de entrega.

- **Três tipos de produto, duas máquinas.** O roteamento é feito assim:

  | `MÁQUINA` | Vai para |
  |---|---|
  | TRANÇADEIRA | `Programação trançadeiras` |
  | TEAR | `Programação Teares` |
  | ELÁSTICO | `Programação Teares` |

  O filtro dos teares é `=SE(K="TEAR";A;SE(K="elástico";A;" "))`. Ou seja, elástico é
  produzido em tear. Na amostra: 36 trançadeira, 17 tear, 4 elástico.

- **"JUNTAR MM" só funciona para milímetros.** A macro `Juntar` (Módulo8) anexa o tamanho
  à descrição apenas quando o tamanho contém `mm`. Produtos em `CM` nunca recebem esse
  tratamento. Pode ser intencional (a largura em mm já está na descrição) — vale
  confirmar.

- **Itens sem código no cadastro.** Nas 60 linhas amostradas de
  `DADOS GERAIS DE PRODUTOS`, 40 estavam com a coluna B (o código) vazia — só descrição,
  tamanho e máquina. Um item sem código **nunca é encontrado** pelo `PROCV`, e o pedido
  correspondente sai com `#N/A`. **Ressalva:** são só as 60 primeiras de 15.824 linhas, e
  as primeiras podem ser resíduo antigo. Não dá para concluir que o cadastro inteiro está
  assim.

---

## Perguntas desta parte

1. **Pares ou peças** — como o sistema sabe qual é qual? Se "par" significa duas unidades,
   isso **dobra o consumo** de metros e de fio. É a informação de maior impacto no cálculo,
   e hoje não vi nada na planilha que faça essa distinção.
2. **`PRAZO`** — confirma que é "dias restantes na data da extração"? E concorda em
   calculá-lo ao vivo no sistema novo?
3. **`CÓD. FILIAL`** — é a filial do cliente ou o próprio cliente?
4. **Peso do fio** — `0,0031` está em que unidade? Quilos por metro de produto?
5. **"JUNTAR MM"** — os produtos em CM ficarem sem o tamanho na descrição é intencional?

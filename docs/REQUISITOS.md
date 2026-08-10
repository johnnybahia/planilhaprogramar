# Requisitos do sistema novo

Decisões já tomadas, para serem implementadas. Separado de
[`COMO-FUNCIONA.md`](COMO-FUNCIONA.md), que descreve a planilha **como ela é hoje**.

---

## R1 — Unidade de compra por cliente

**Decidido na Parte 2.**

Substitui os nomes de clientes escritos dentro de fórmulas (hoje em três células, com
três grafias, uma delas com espaço duplo que nunca casa).

**Tabela `clientes`**

| campo | tipo | observação |
|---|---|---|
| `codigo` | número | vem do `CÓD. CLIENTE` da importação |
| `nome` | texto | ex.: PAQUETÁ ITAPAJÉ |
| `unidade` | `PAR` \| `PECA` | padrão `PAR` |

A regra é avaliada **por linha de pedido**, pelo código do cliente — não mais por lote.
Um lote pode misturar clientes.

Clientes que hoje usam peças: **PAQUETÁ ITAPAJÉ, PAQUETÁ PENTECOSTE, PAQUETÁ PROJEÇÃO**.

---

## R2 — Percentuais adicionais de produção

**Decidido nesta rodada.**

| Decisão | Escolha |
|---|---|
| Onde ficam os percentuais | **Valores globais**, numa tela de configuração |
| Produção e fio | **O mesmo percentual** — o fio é consequência dos metros |
| Como combinam | **Um ou outro, nunca acumulam** |
| No relatório | **Separado**: base, adicional e total |
| Vale para | **Teares e trançadeiras**, igualmente |

### Configuração

**Tabela `configuracao`** — dois percentuais globais, editáveis numa tela:

| campo | tipo | padrão | quando se aplica |
|---|---|---|---|
| `percentual_producao_normal` | número (%) | `0` | pedidos comuns |
| `percentual_reposicao` | número (%) | `25` | pedidos marcados como reposição |

> Com `percentual_producao_normal = 0`, o resultado de um pedido comum é **idêntico ao da
> planilha atual**. O campo só altera número quando alguém o preenche de propósito.

### No pedido

**Campo `reposicao`** — sim/não. É ele que escolhe **qual** dos dois percentuais entra.

### O cálculo

```js
const fatorUnidade = cliente.unidade === 'PECA' ? 1 : 2;

// um ou outro, nunca os dois
const percentual = pedido.reposicao
  ? config.percentual_reposicao
  : config.percentual_producao_normal;

const metrosBase      = pedido.quantidade * produto.tamanho_m * fatorUnidade;
const metrosAdicional = metrosBase * (percentual / 100);
const metrosTotal     = metrosBase + metrosAdicional;
```

E o consumo de fio segue os metros, como hoje:

```js
// por cor da referência
consumoBase      = metrosBase      * peso_da_cor;   // kg
consumoAdicional = metrosAdicional * peso_da_cor;   // kg
consumoTotal     = metrosTotal     * peso_da_cor;   // kg
```

### Fatores resultantes

| Cliente | Reposição | Percentual aplicado | Fator total |
|---|---|---|---|
| Pares (padrão) | não | normal = 0 | **2,00** ← igual a hoje |
| Pares (padrão) | não | normal = 5 | 2,10 |
| Pares (padrão) | sim | reposição = 25 | **2,50** ← igual a hoje |
| Peças (Paquetá) | não | normal = 0 | **1,00** ← igual a hoje |
| Peças (Paquetá) | não | normal = 5 | 1,05 |
| Peças (Paquetá) | sim | reposição = 25 | 1,25 ← hoje daria 1,00 |

### O que isso corrige

**1. Reposição passa a funcionar nos teares.** Hoje não funciona: a fórmula testa
`SE($M$1="reposição";…)`, mas `M1` só pode devolver `"2"` ou `FALSO`. O ramo é
inalcançável — só as trançadeiras fazem reposição.

**2. Peças e reposição deixam de ser mutuamente exclusivos.** Hoje a fórmula é:

```excel
=SE($P$1="/2"; B3*G3; SE($F$1="REPOSIÇÃO"; B3*G3*2*1,25; G3*2*B3))
```

Os dois testes leem **a mesma célula** (`PEDIDOS!D2`), que só pode conter um valor: ou o
nome do cliente Paquetá, ou a palavra `REPOSIÇÃO`. Nunca os dois. Ou seja: **um pedido
Paquetá com reposição é impossível de representar hoje** — e mesmo que fosse, o teste de
peças vem primeiro e a reposição seria ignorada.

Na fórmula nova os dois fatores **se multiplicam** em vez de competir:

| Cliente | Reposição | Fator hoje | Fator novo |
|---|---|---|---|
| Padrão (pares) | não | 2 | 2 |
| Padrão (pares) | sim | 2,5 | 2,5 |
| Paquetá (peças) | não | 1 | 1 |
| Paquetá (peças) | **sim** | **1** ← perde a reposição | **1,25** |

**3. O percentual vira configuração.** Mudar de 25% para 30% passa a ser editar um campo,
em vez de achar e alterar `1,25` dentro de fórmulas em duas abas.

### No relatório

As três grandezas aparecem separadas, em metros e em quilos por cor. **O rótulo da linha
do meio acompanha a origem do percentual**, para não chamar de reposição o que não é:

```
FITA MFP 101 15MM — pedido de reposição
  Pedido           1.000,0 m     3,10 kg  (cor 100/1)
  Reposição 25%      250,0 m     0,78 kg
  ──────────────────────────────────────
  Total            1.250,0 m     3,88 kg


FITA MFP 101 15MM — pedido comum, com adicional global de 5%
  Pedido           1.000,0 m     3,10 kg  (cor 100/1)
  Adicional 5%        50,0 m     0,16 kg
  ──────────────────────────────────────
  Total            1.050,0 m     3,26 kg
```

Quando o percentual aplicado for `0`, a linha do adicional **não aparece** — o relatório
fica igual ao de hoje.

### Arredondamento

Os pesos de fio são números muito pequenos (`0,0031` kg/m) multiplicados por milhares de
metros, e o resultado é arredondado para exibição. Conferindo os exemplos acima, o
adicional de 5% dá exatamente `0,155` kg — que arredonda para `0,16` com a regra decimal
normal, mas para `0,15` se o cálculo usar ponto flutuante binário, porque `0,155` não tem
representação exata em binário.

Para um relatório que orienta compra de fio, essa diferença não pode depender de acaso:

- somar e arredondar com **aritmética decimal** (não `float`), com arredondamento
  **meio para cima**;
- arredondar **só na exibição**, nunca nos valores intermediários;
- somar as parcelas antes de arredondar o total, para o total nunca divergir da soma das
  linhas mostradas.

---

## R3 — Cores e pesos sem limite de quantidade

**Decidido nesta rodada.**

Hoje há **dois tetos**, um por cima do outro:

| Onde | Limite |
|---|---|
| `PESOS DE FIOS` | 5 pares cor/peso, nas colunas B/C, D/E, F/G, H/I, J/K |
| Fórmulas de programação | leem apenas **4** — as colunas `J` e `K` nunca são consultadas |

Ou seja, mesmo cadastrando a quinta cor ela seria ignorada, sem aviso. Hoje isso não
causa erro porque nenhum produto passa de 3 cores — é um limite latente.

### O modelo novo

Cores deixam de ser colunas e viram **linhas de uma tabela filha**. Sem limite.

**Tabela `produto_fios`**

| campo | tipo | observação |
|---|---|---|
| `referencia` | texto | liga ao produto |
| `ordem` | número | 1, 2, 3… — a ordem do fio na construção |
| `tipo` | `COR` \| `ESTRUTURAL` | ver 2.6 do `COMO-FUNCIONA.md` |
| `codigo` | texto | `100`, `58`, `ENCHIMENTO`, `BORRACHA PRETA 28` |
| `qualificador` | texto | opcional: `RECICLADO`, `LAVADO`, `(30-2)` |
| `peso_kg_por_m` | decimal | ex.: `0,0031` |

O campo `tipo` preserva uma distinção que hoje existe só implicitamente: fio de cor
(cujo código aparece na descrição) versus material estrutural (enchimento, borracha).
Isso permite, por exemplo, somar o consumo de borracha separado do consumo de fio
colorido — hoje impossível, porque tudo está na mesma fila de colunas.

### O cálculo

Deixa de ser quatro `PROCV` fixos e passa a ser uma soma sobre todas as linhas:

```js
const consumo = produto.fios.map(fio => ({
  codigo:     fio.codigo,
  qualificador: fio.qualificador,
  tipo:       fio.tipo,
  base:       metrosBase      * fio.peso_kg_por_m,
  adicional:  metrosAdicional * fio.peso_kg_por_m,
  total:      metrosTotal     * fio.peso_kg_por_m,
}));
```

Um produto com 12 cores funciona igual a um com 1.

### Na migração

Ao importar `PESOS DE FIOS`, cada par cor/peso preenchido vira uma linha, com `ordem`
seguindo a posição da coluna. O `tipo` é deduzido assim: `ENCHIMENTO` e tudo que começa
com `BORRACHA` viram `ESTRUTURAL`; o resto vira `COR`. **A lista de estruturais precisa
ser confirmada** — foi deduzida da amostra de 59 produtos e pode haver outros materiais
nas 6.043 linhas completas.

---

## Pendências que afetam estes requisitos

- **O que são os 25%?** Perda de processo, margem de segurança ou decisão comercial.
  Não muda a implementação, mas muda o nome do campo na tela e o texto de ajuda.
- **`PRAZO` calculado ao vivo** (`entrega − hoje`) em vez de importado congelado —
  proposto na Parte 1, aguardando confirmação.

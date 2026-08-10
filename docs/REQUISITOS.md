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

## R2 — Percentual de reposição

**Decidido nesta rodada.**

| Decisão | Escolha |
|---|---|
| Onde fica o percentual | **Um valor global**, numa tela de configuração |
| Produção e fio | **O mesmo percentual** — o fio é consequência dos metros |
| No relatório | **Separado**: base, adicional e total |
| Vale para | **Teares e trançadeiras**, igualmente |

### Configuração

**Tabela `configuracao`**

| campo | tipo | padrão |
|---|---|---|
| `percentual_reposicao` | número (%) | `25` |

Um único valor, editável numa tela. Hoje ele está fixo dentro da fórmula como `1,25`.

### No pedido

**Campo `reposicao`** — sim/não, por pedido. Quando ligado, aplica o percentual global.

### O cálculo

```js
const fatorUnidade   = cliente.unidade === 'PECA' ? 1 : 2;
const fatorReposicao = pedido.reposicao ? 1 + config.percentual_reposicao / 100 : 1;

const metrosBase     = pedido.quantidade * produto.tamanho_m * fatorUnidade;
const metrosAdicional= metrosBase * (fatorReposicao - 1);
const metrosTotal    = metrosBase + metrosAdicional;
```

E o consumo de fio segue os metros, como hoje:

```js
// por cor da referência
consumoBase      = metrosBase      * peso_da_cor;   // kg
consumoAdicional = metrosAdicional * peso_da_cor;   // kg
consumoTotal     = metrosTotal     * peso_da_cor;   // kg
```

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

As três grandezas aparecem separadas, em metros e em quilos por cor:

```
FITA MFP 101 15MM
  Pedido      1.000,0 m     3,10 kg  (cor 100/1)
  Reposição     250,0 m     0,78 kg
  ─────────────────────────────────
  Total       1.250,0 m     3,88 kg
```

---

## Pendências que afetam estes requisitos

- **O que são os 25%?** Perda de processo, margem de segurança ou decisão comercial.
  Não muda a implementação, mas muda o nome do campo na tela e o texto de ajuda.
- **`PRAZO` calculado ao vivo** (`entrega − hoje`) em vez de importado congelado —
  proposto na Parte 1, aguardando confirmação.

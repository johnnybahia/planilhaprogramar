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

## R3 — Peso dos fios, sem limite de quantidade

**Revisado nesta rodada.** A versão anterior juntava peso e estrutura numa tabela só,
com um campo `tipo`. Estava errado: **são fatos independentes**, e a própria planilha já
os separa — em abas diferentes.

| | Onde mora hoje | O que é |
|---|---|---|
| **Peso** | `PESOS DE FIOS` | todo fio que tem massa contribui para o consumo |
| **Estrutura** | `DADOS DOS PRODUTOS` | só o que ocupa fuso ou posição na urdidura |

Confirmado nos dados: `ENCHIMENTO` aparece **22 vezes** em `PESOS DE FIOS` e **zero**
vezes em `DADOS DOS PRODUTOS`. `BORRACHA`, 5 e zero. Eles pesam, mas não ocupam
estrutura — exatamente o "algo por fora" descrito pelo usuário.

**Tabela `produto_fios`** — um registro por fio com massa, sem limite de quantidade:

| campo | tipo | observação |
|---|---|---|
| `referencia` | texto | liga ao produto |
| `codigo` | texto | `100`, `58`, `ENCHIMENTO`, `BORRACHA PRETA 28` |
| `qualificador` | texto | opcional: `RECICLADO`, `LAVADO`, `(30-2)` |
| `peso_kg_por_m` | decimal | ex.: `0,0031` |

Hoje há **dois tetos sobrepostos**: a planilha comporta 5 pares cor/peso e as fórmulas
leem apenas 4 — a quinta cor seria ignorada sem aviso. Aqui não há teto.

O cálculo deixa de ser quatro `PROCV` fixos e passa a somar sobre todas as linhas:

```js
const consumo = produto.fios.map(fio => ({
  codigo:    fio.codigo,
  base:      metrosBase      * fio.peso_kg_por_m,
  adicional: metrosAdicional * fio.peso_kg_por_m,
  total:     metrosTotal     * fio.peso_kg_por_m,
}));
```

O enchimento entra nessa soma como qualquer outro fio — é o que garante a "base de peso
universal". E como **cada produto tem seu próprio peso de enchimento**, ele é um campo do
produto, não uma constante. (Na amostra: `0,00087` em 17 produtos e `0,0021` em 5.)

---

## R4 — Estrutura do produto, com duas formas

**Decidido nesta rodada.** A estrutura descreve **o que ocupa fuso ou posição**, e tem
forma diferente conforme a máquina (ver Parte 4 do `COMO-FUNCIONA.md`).

### Trançadeira — conjunto de cores com espulas

**Tabela `estrutura_trancadeira`**

| campo | tipo |
|---|---|
| `referencia` | texto |
| `codigo_fio` | texto — liga a `produto_fios` |
| `espulas` | número |
| `fios_por_espula` | número |

Mais dois campos no próprio produto, que hoje moram nas colunas `N` e `O`:
`voltas_na_espula` e `producao_metros`. São eles que fazem a conversão
`voltas = metros × voltas_na_espula ÷ producao_metros`, e **variam por produto**.

### Tear — sequência ordenada da urdidura

**Tabela `estrutura_tear`**

| campo | tipo |
|---|---|
| `referencia` | texto |
| `ordem` | número — 1, 2, 3… |
| `codigo_fio` | texto |
| `fios` | número |

`ATAC 6ALF 212/100 6MM` vira:

| ordem | fio | fios |
|---|---|---|
| 1 | 212 | 1 |
| 2 | 100 | 52 |
| 3 | 212 | 3 |

**A ordem é dado, não apresentação.** A mesma cor repete em posições diferentes porque
é a sequência de montagem da urdidura. O total (56 fios) passa a ser **calculado**.

Isso elimina o texto livre com três separadores diferentes (`.`, `,`, `-`) que existe
hoje, e com ele a possibilidade de uma quarta variação aparecer amanhã.

### O enchimento na estrutura

Em ~99% dos casos o enchimento **não** entra na estrutura: só pesa. Mas em alguns
produtos ele faz parte da construção. O modelo cobre os dois casos sem campo extra:
basta o fio ter, ou não, uma linha em `estrutura_trancadeira` / `estrutura_tear`.

- **só pesa** → aparece em `produto_fios`, ausente da estrutura
- **também estrutura** → aparece nos dois, e ocupa espulas normalmente

---

## R5 — Cadastro de máquinas e limite de fusos

**Decidido nesta rodada.**

Hoje existem trançadeiras de **16, 32 e 48 fusos**, e novos modelos podem surgir. Esse
número não pode ficar dentro de fórmula.

**Tabela `modelos_maquina`**

| campo | tipo | exemplo |
|---|---|---|
| `nome` | texto | `Trançadeira 48` |
| `tipo` | `TRANCADEIRA` \| `TEAR` | |
| `fusos` | número | `48` |

Um modelo novo de 64 fusos entra como **um registro**, sem tocar em cálculo nenhum.

### Validação

> **A soma das espulas de um produto não pode passar do total de fusos do modelo.**

```js
const ocupados = estrutura.reduce((s, e) => s + e.espulas, 0);
if (ocupados > modelo.fusos) {
  erro(`${ocupados} espulas não cabem em ${modelo.fusos} fusos`);
}
```

Hoje nada impede cadastrar uma ficha que não cabe na máquina — o erro só apareceria no
chão de fábrica, com a máquina parada. Essa checagem custa uma linha e evita isso.

---

## Pendências que afetam estes requisitos

- **O que são os 25%?** Perda de processo, margem de segurança ou decisão comercial.
  Não muda a implementação, mas muda o nome do campo na tela e o texto de ajuda.
- **`PRAZO` calculado ao vivo** (`entrega − hoje`) em vez de importado congelado —
  proposto na Parte 1, aguardando confirmação.

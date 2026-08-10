# Análise da planilha

Baseada na exportação de `PROGRAMAÇÃO CEARA 27-01-25 - Atualizado.xlsb` (36,1 MB,
23 abas), feita em 2026-08-10.

---

## Resumo em uma frase

A planilha tem **4.733.724 células com fórmula**, mas apenas **1.350 regras distintas** —
uma regra para cada 3.506 células. Quase tudo é repetição, e três decisões estruturais
respondem por praticamente todo o peso e toda a lentidão.

---

## O que o sistema faz

Produção de **fitas e elásticos têxteis** para calçadistas (DASS, Grendene, Dilly,
Copershoes). O fluxo é:

1. Pedidos entram na aba `PEDIDOS` com código, quantidade, cliente e prazo.
2. Cada produto é classificado como **TEAR** ou **TRANÇADEIRA** (coluna K).
3. O pedido é distribuído entre as máquinas disponíveis daquele tipo.
4. Calcula-se quanto tempo cada máquina leva, e sai o relatório de programação.

Existem **duas linhas de produção espelhadas**, cada uma com seu trio de abas:

| | Trançadeiras | Teares |
|---|---|---|
| Fila | `Programação trançadeiras` | `Programação Teares` |
| Cálculo | `Planejamento Trançadeira` | `Planejamento tear` |
| Saída | `RELATÓRIO TRANÇADEIRAS` | `REALTÓRIO TEARES` |

A capacidade está anotada na linha 1 de `PEDIDOS`:
**máximo 131 itens em trançadeiras, 100 itens em tear.**

### Topologia real

```
CADASTRO CLIENTES ─┐
                   ├─> PEDIDOS ─┬─> Programação trançadeiras ─> Planejamento Trançadeira ─┐
DADOS GERAIS DE ───┘            │                                        ↑                 │
PRODUTOS                        │                                        └─────────────────┤
                                │                                                          ↓
                                │                                          RELATÓRIO TRANÇADEIRAS
                                │
                                └─> Programação Teares ─> Planejamento tear ─> REALTÓRIO TEARES
                                                                                      ↑
PESOS DE FIOS ─> DADOS DOS PRODUTOS ──────────────────────────────────────────────────┤
                                                                                      │
CORES USADAS ─┐                                                                       │
              ├─> CONSUMO ────────────────────────────────────────────────────────────┘
Relação de ───┘
referencias
(MORTA)
```

Note o **ciclo**: `Planejamento` lê do `RELATÓRIO` e o `RELATÓRIO` lê do `Planejamento`.
Não é referência circular de célula — são colunas diferentes —, mas obriga o cálculo a
acontecer em duas passadas.

---

## Descoberta 1 — `PEDIDOS` é 89% da planilha e a causa dos 36 MB

`PEDIDOS` tem **4.197.138 fórmulas em apenas 10 padrões distintos**. Quatro colunas
inteiras (J, K, L e N) estão preenchidas **da linha 4 até a linha 1.048.576** — a última
linha que o Excel permite.

Só isso já explica o tamanho do arquivo:

| | |
|---|---|
| Células nas 4 colunas cheias | 4.194.292 |
| A 8 bytes por célula | **~32 MB** |
| Tamanho real do arquivo | 36,1 MB |

Pior que o tamanho é o custo de cálculo. A coluna N repete, em **cada uma das 1.048.573
linhas**, esta fórmula matricial de valores únicos:

```
=SEERRO(ÍNDICE($J$4:$J$2502; MENOR(SE(CORRESP(SE($J$4:$J$2502="";"";$J$4:$J$2502);
 SE($J$4:$J$2502="";"";$J$4:$J$2502);0)= LIN(INDIRETO("1:"&LINS($J$4:$J$2502)));
 CORRESP(...);"");  LIN(INDIRETO("1:"&LINS($J$4:$J$2502)))));"")
```

Cada célula varre 2.499 linhas comparando o intervalo inteiro contra ele mesmo: cerca de
**12,5 milhões de comparações por célula**. Multiplicado por 1.048.573 células, dá
aproximadamente **1,3 × 10¹³ comparações a cada recálculo completo**.

É por isso que a planilha está em **modo de cálculo manual**. Não é preferência — é a
única forma de ela abrir.

**O que essa fórmula toda faz:** devolve a lista de valores únicos da coluna J. No
Google Planilhas isso é `=UNIQUE(J4:J2502)`, uma fórmula só. Num banco de dados é
`SELECT DISTINCT`.

As outras três colunas (J, K, L) são `PROCV` em `DADOS GERAIS DE PRODUTOS`, também
estendidos por um milhão de linhas. São junções de tabela — no sistema novo, um `JOIN`.

> **Ganho rápido, sem mudar nada:** apagar as linhas vazias abaixo da última linha com
> dado real (~2.437) em `PEDIDOS` deve derrubar o arquivo para poucos MB e devolver o
> cálculo automático. Ver `docs/COMO-EXPORTAR.md`.

---

## Descoberta 2 — 151.385 fórmulas que calculam um resto de divisão

`Planejamento Trançadeira` ocupa **137 linhas × 1.103 colunas** (de C até APM). Cada linha
é uma máquina. As colunas se alternam em pares, com apenas duas fórmulas:

```
C1 = SE(B1>A1; A1; SE(B1<A1; 0; A1))     ← quanto produzir neste período
D1 = SE(C1=A1;  B1-C1; 0)                ← quanto sobra para o próximo
```

Isso é um **laço de simulação desenrolado ao longo das colunas**: a cada par, produz até o
limite da capacidade e passa o resto adiante, até a demanda zerar. São **551 períodos**
simulados por máquina.

Depois, o relatório lê o resultado com a fórmula do `Módulo17`:

```
=MENOR('Planejamento Trançadeira'!C1:FJ1; CONT.SE('Planejamento Trançadeira'!C1:FJ1;0)+1)
```

que significa "o menor valor diferente de zero da linha" — ou seja, **a sobra do último
período**.

**Verifiquei numericamente contra os dados reais exportados:**

| Máquina | Demanda (voltas) | Capacidade | Simulação das 1.103 colunas | `demanda % capacidade` | Valor no relatório |
|---|---|---|---|---|---|
| 2 | 10.809,1914893617 | 1.200 | 9,1914893617 | 9,1914893617 | 9,19148936170131 |
| 3 | 6.838,46808510638 | 1.200 | 838,4680851063 | 838,4680851063 | 838,468085106382 |

São idênticos. **As 151.385 fórmulas calculam exatamente `demanda módulo capacidade`.**
O número de períodos é `demanda / capacidade` — que já está na coluna "Máq." do relatório
(9,0077 e 5,6987, conferem).

Em JavaScript, a aba inteira é:

```js
const periodos = Math.ceil(demanda / capacidade);
const sobra    = demanda % capacidade;
```

O mesmo vale para `Planejamento tear` (138.700 fórmulas, 16 padrões).
**Somadas, as duas abas de planejamento são 290.085 fórmulas que viram duas linhas de código.**

### Um limite invisível

A planilha tem 551 pares de colunas. A 1.200 voltas por período, ela só consegue
representar demandas de até **661.200 voltas**. Acima disso o resultado fica errado e
**nada avisa** — a fórmula `MENOR` simplesmente devolve um valor de uma coluna que ainda
não zerou. Com a fórmula de módulo esse teto deixa de existir.

---

## Descoberta 3 — uma aba morta alimentando o relatório, em silêncio

A aba `Relação de referencias` tem **12.500 células contendo literalmente `=#REF!`**. A
origem dos dados foi apagada em algum momento, e as fórmulas ficaram órfãs. A aba está
100% quebrada.

O problema é que ela **não está isolada**. A aba `CONSUMO` a consulta 39 vezes:

```
=SEERRO(SE($A3="";"";PROCV($A3;'Relação de referencias'!$A$1:$E$500;5;0));"")
```

Repare no `SEERRO`. Ele captura o erro e devolve **vazio**. Então `CONSUMO` não mostra
erro nenhum — mostra células em branco, como se não houvesse dado para aquele item.
E `CONSUMO` alimenta `REALTÓRIO TEARES` e `REALTÓRIO TEARES NOVO`.

**Este é o achado mais preocupante da análise.** Um erro visível é um erro que alguém
conserta; este está mascarado há tempo indeterminado, e o relatório de teares pode estar
com informação de consumo faltando sem que ninguém perceba.

Vale conferir se as colunas de consumo do relatório de teares estão saindo preenchidas.

---

## Estado de saúde

### Erros nas amostras exportadas

| Aba | Células com erro |
|---|---|
| `Relação de referencias` | **100%** (aba morta) |
| `Planejamento Trançadeira` | 88% |
| `CORES USADAS` | 20% |
| `Programação trançadeiras` | 19% |
| `Programação Teares` | 15% |
| `RELATÓRIO TRANÇADEIRAS` | 6% |

Os 88% do planejamento **não são falha de cálculo**: apenas as linhas 2 a 8 estavam
ativas na exportação. Das 137 linhas de máquina, **7 tinham pedido**. As demais recebem
erro da linha de origem vazia e propagam por todas as 1.103 colunas. É desperdício e
ruído visual, não corrupção — mas obriga qualquer leitura da aba a ignorar erro como se
fosse normal, o que é exatamente como erros de verdade passam despercebidos.

### Outros pontos

- **`_xlfn.SINGLE` = `#NAME?`** e **`'RELATÓRIO TRANÇADEIRAS'!Extract` = `#REF!`** —
  nomes definidos quebrados.
- **`REALTÓRIO TEARES` / `REALTÓRIO TEARES NOVO`** — erro de digitação em "RELATÓRIO",
  em duas abas. A "NOVO" (oculta, 8.895 fórmulas) parece uma reescrita da antiga (11.308)
  que ficou pela metade. **Duas versões do mesmo relatório convivendo.**
- **`PEDIDOS!M1`** contém `=SE(K4:K1273="TEAR";"TEAR";"NÃO")`, uma fórmula matricial
  antiga que hoje devolve erro.
- **Abas vazias**: `RELAÇÃO DE CONSUMOS`, `Planilha1`, `Plan2`.
- **Nenhuma validação de dados** em toda a planilha. Nada impede digitar um código de
  produto inexistente ou uma máquina inválida — o erro só aparece lá na frente, como
  `#N/A` num `PROCV`.
- **Formatação condicional**: 557 regras, mas quase todas de destaque de duplicados e
  comparação com textos fixos (`"ERRADO"`, `"PAQUETA"`, `"TRANÇADEIRA"`). Eu esperava
  encontrar regras de prazo (atrasado / no prazo); **não existem**. As cores não carregam
  regra de negócio relevante.

### Botões e macros

| Aba | Botão | Macro |
|---|---|---|
| `PEDIDOS` | "1 - criar código" | `JuntarDados` (Módulo2) |
| `PEDIDOS` | "JUNTAR MM" | `Juntar` (Módulo8) |
| `PEDIDOS` | "APAGAR" | `APAGAR` (Módulo13) |
| `RELATÓRIO TRANÇADEIRAS` | "AJUSTAR" | `ajustaralturalinha` (Módulo14) |
| `PESOS DE FIOS` | "GERAR ITENS NOVOS" | `EstaPasta_de_trabalho.copiarDados` |

Isso confirma que os Módulos 2, 8 e 13 estão em uso — a rotina de trabalho é
*criar código → juntar MM → programar*.

⚠️ **`copiarDados` não está no repositório.** Ela vive no módulo `EstaPasta_de_trabalho`
(ThisWorkbook), que não foi exportado junto com os `Módulo*.bas`. É a única macro em uso
cujo código eu não vi.

---

## O que isso significa para o sistema novo

O trabalho é **muito menor do que os 36 MB sugerem**:

| Aba | Fórmulas | Vira o quê |
|---|---|---|
| `PEDIDOS` (J, K, L) | 3.145.719 | `JOIN` com a tabela de produtos |
| `PEDIDOS` (N, Q) | 1.050.129 | `SELECT DISTINCT` |
| `Planejamento` ×2 | 290.085 | `demanda % capacidade` |
| `Programação` ×2 | 119.796 | `WHERE tipo = 'TEAR'` / `'TRANÇADEIRA'` |
| `CORES USADAS`, `CONSUMO` | 45.167 | consultas com junção |
| **Subtotal** | **4.650.896 (98%)** | **operações banais de banco de dados** |

Sobram ~83 mil fórmulas de cálculo real (metragem, voltas, espulas, pesos de fio) — essas
sim precisam ser lidas com atenção e reimplementadas com cuidado.

Sobre a decisão de arquitetura: **os volumes cabem tranquilamente no Google Planilhas.**
Os dados reais são ~2.437 pedidos, ~15.824 produtos, ~6.043 pesos de fio, 137+322 linhas
de planejamento. Tudo isso somado fica na casa das dezenas de milhares de células — bem
longe do limite de 10 milhões. O que estourava o Excel eram as colunas de um milhão de
linhas, que simplesmente não vão existir.

---

## Perguntas em aberto

Estas eu não consigo responder pelos arquivos — preciso de você:

1. **A capacidade de 1.200 voltas por período é fixa para toda máquina?** Vi 1.200 em
   todas as linhas amostradas, mas a coluna A do planejamento pode variar por máquina.
2. **O que é um "período" no planejamento?** Um turno? Um dia? Uma hora de produção?
   Isso define como o resultado vira data de entrega.
3. **`REALTÓRIO TEARES` ou `REALTÓRIO TEARES NOVO`** — qual está valendo hoje?
4. **As colunas de consumo do relatório de teares estão vindo preenchidas?** É o teste
   para saber há quanto tempo a `Relação de referencias` está quebrada.
5. **O que faz a macro `copiarDados`** do botão "GERAR ITENS NOVOS"?
6. **Quem digita o quê?** Alguém preenche `PEDIDOS` na mão, ou vem colado de um ERP?

---

## Sobre a rodada anterior

Vale registrar o que a exportação corrigiu do que eu havia deduzido só pelo VBA:

| Deduzi | Realidade |
|---|---|
| 3 abas principais | **23 abas**, duas linhas de produção espelhadas |
| Planejamento com 164 colunas (C:FJ) | **1.103 colunas** — C:FJ é só o trecho que o relatório lê |
| ~129 máquinas | 137 linhas de trançadeira + 322 de tear; **7 ativas** na exportação |
| Cores carregam regra de negócio | Não carregam — são destaque de duplicados |
| Arquivo grande por formatação | Não: por 4 colunas de fórmula com 1 milhão de linhas |

Duas previsões se confirmaram exatamente: as **129 fórmulas `MENOR`** individuais no
relatório, e o **passo de 6 linhas** por bloco de máquina.

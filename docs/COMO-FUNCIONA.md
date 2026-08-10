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

1. ~~**Pares ou peças**~~ — **respondido na Parte 2.**
2. **`PRAZO`** — confirma que é "dias restantes na data da extração"? E concorda em
   calculá-lo ao vivo no sistema novo?
3. **`CÓD. FILIAL`** — é a filial do cliente ou o próprio cliente?
4. **Peso do fio** — `0,0031` está em que unidade? Quilos por metro de produto?
5. **"JUNTAR MM"** — os produtos em CM ficarem sem o tamanho na descrição é intencional?

---

# Parte 2 — O cálculo de metros e o consumo de fio

Esta é a parte que justifica o sistema inteiro. Encontrei a cadeia completa nas abas
`Programação trançadeiras` e `Programação Teares`, e ela confirma exatamente a descrição
do usuário.

## 2.1 A fórmula central

Na coluna **H** das duas abas de programação:

```excel
=SE($P$1="/2";        B3*G3;           ← peças:     quantidade × tamanho
 SE($F$1="REPOSIÇÃO"; B3*G3*2*1,25;    ← reposição: quantidade × tamanho × 2 × 1,25
                      G3*2*B3))        ← padrão:    quantidade × tamanho × 2
```

Onde `B` = quantidade do pedido e `G` = tamanho em metros (vindo do `PROCV` em
`DADOS GERAIS DE PRODUTOS`).

Em forma limpa:

```
metros_da_linha = quantidade × tamanho_m × fator

fator = 2      → padrão: pedido em PARES (cada par são duas peças)
      = 1      → cliente que compra em PEÇAS
      = 2,5    → REPOSIÇÃO (2 × 1,25 — 25% a mais)
```

## 2.2 Da linha ao consumo de fio

```
    PEDIDOS (linha do pedido)
        │  quantidade × tamanho × fator
        ▼
    H = metros da linha
        │  SOMASES(H; F; I)   agrupa por referência
        ▼
    J = metros totais da referência
        │  × peso de cada cor  (PROCV em PESOS DE FIOS)
        ▼
    Q, R, S, T = consumo em kg de cada cor
```

As colunas se organizam assim nas duas abas de programação:

| Col | Conteúdo | Origem |
|---|---|---|
| `F` | Referência do item | `PROCV` em `DADOS GERAIS DE PRODUTOS` |
| `G` | Tamanho em metros | `PROCV` em `DADOS GERAIS DE PRODUTOS` |
| `H` | **Metros da linha** | a fórmula de 2.1 |
| `I` | Lista de referências únicas | fórmula matricial |
| `J` | **Metros totais por referência** | `=SOMASES(H$3:H$2501; F$3:F$2501; I3)` |
| `L, M, N, O` | Nome das cores 1 a 4 | `PROCV` em `PESOS DE FIOS`, colunas 2/4/6/8 |
| `Q, R, S, T` | **Consumo em kg das cores 1 a 4** | `PROCV` colunas 3/5/7/9 **× J** |

É exatamente a descrição do usuário: quantidade de pares × o dobro do tamanho → metros;
metros × peso de cada cor → consumo de fio em kg por cor por item.

## 2.3 ⚠️ Onde está hoje a lista de "clientes que usam peças"

O usuário pediu **um lugar para cadastrar os clientes que compram em peças**. Vale saber
onde essa informação mora hoje: **em três fórmulas diferentes, em três células
diferentes, com três grafias diferentes.**

| Aba | Célula | Fórmula |
|---|---|---|
| `Programação trançadeiras` | `P1` | `=SE(F1="PAQUETÁ ITAPAJÉ";"/2";SE(F1="PAQUETÁ PENTECOSTE";"/2";SE(F1="PAQUETÁ PROJEÇÃO";"/2")))` |
| `Programação Teares` | `M1` | `=SE(I1="paquetá itapajé";"2";SE(I1="PAQUETÁ PENTECOSTE";"2";SE(I1="PAQUETÁ PROJEÇÃO";"2")))` |
| `Programação Teares` | `Z4` | `=SE(S4="paquetá itapajé";"/2";SE(S4="paquetá  pentecoste";"/2";SE(S4="PAQUETÁ PROJEÇÃO";"/2")))` |

São três clientes: **PAQUETÁ ITAPAJÉ, PAQUETÁ PENTECOSTE e PAQUETÁ PROJEÇÃO**.

Três problemas concretos:

1. **`Z4` tem `"paquetá  pentecoste"` com dois espaços.** O Excel ignora maiúsculas na
   comparação, mas **não ignora espaço**. Esse ramo nunca casa.
2. **Incluir um quarto cliente exige editar fórmula em três lugares.** É o tipo de
   alteração que se faz em dois e se esquece do terceiro.
3. **O interruptor é global do lote, não por linha.** `P1`, `M1` e `Z4` são células
   únicas que valem para a aba inteira. Se um lote misturar Paquetá com outro cliente,
   **todos são calculados do mesmo jeito**.

### E as duas linhas ligam o interruptor de formas diferentes

| Aba | Lê de | Conteúdo na exportação |
|---|---|---|
| `Programação Teares` | `I1 = PEDIDOS!C2` | `DAKOTA` — `C2` tem `PROCV` no `CADASTRO CLIENTES`, é automático |
| `Programação trançadeiras` | `F1 = PEDIDOS!D2` | **vazio** — `D2` não tem fórmula, é digitação manual |

Ou seja: nos **teares** o modo peças é detectado sozinho a partir do cliente do primeiro
pedido. Nas **trançadeiras** ele depende de alguém digitar o nome do cliente (ou a palavra
`REPOSIÇÃO`) na célula `D2`, que estava vazia na exportação.

**O risco operacional:** se um pedido Paquetá for para trançadeira e ninguém preencher
`D2`, o cálculo usa o fator 2 em vez de 1 — **o dobro dos metros e o dobro do consumo de
fio**. E nada na tela avisa.

> O relatório de teares mostra o modo em `E7`:
> `=SE(PEDIDOS!L1="2";"PAQUETA OK PEÇAS";"PEDIDO EM PARES")`. O de trançadeiras não tem
> equivalente.

### Ramo morto

Na aba `Programação Teares`, a fórmula testa `SE($M$1="reposição";…)`, mas `M1` é a
própria fórmula acima, que só pode devolver `"2"` ou `FALSO` — **nunca** `"reposição"`.
Esse ramo é inalcançável: **reposição não funciona nos teares**, só nas trançadeiras.

## 2.4 Como isso fica no sistema novo

A resposta ao pedido do usuário é uma **tabela de clientes com a unidade de compra**:

| campo | exemplo |
|---|---|
| `codigo` | 133 |
| `nome` | PAQUETÁ ITAPAJÉ |
| `unidade` | `PECA` (padrão: `PAR`) |

E a regra passa a ser avaliada **por linha de pedido**, usando o `CÓD. CLIENTE` que já
vem na importação:

```js
const fator = pedido.reposicao          ? 2.5
            : cliente.unidade === 'PECA' ? 1
            :                              2;

const metros = pedido.quantidade * produto.tamanho_m * fator;
```

O que isso resolve de uma vez:

- cadastrar um novo cliente vira **uma linha numa tabela**, não três fórmulas;
- some a possibilidade de grafias divergentes — a ligação é pelo **código** do cliente,
  não pelo nome digitado;
- um lote **pode misturar clientes**, porque cada linha decide sozinha;
- `REPOSIÇÃO` passa a ser um campo do pedido e funciona nas duas linhas de produção.

## 2.5 Limite latente: a quinta cor

`PESOS DE FIOS` comporta **5 pares cor/peso** (colunas B/C, D/E, F/G, H/I, **J/K**). As
fórmulas de programação leem apenas as colunas 2/4/6/8 e 3/5/7/9 — ou seja, as
**4 primeiras**. As colunas `J` e `K` nunca são lidas.

Na amostra, nenhum produto usa a quinta cor (o máximo são 3), então **hoje isso não
causa erro**. Mas o dia em que alguém cadastrar um produto com 5 cores, o consumo sairá
subestimado sem aviso. No modelo novo, cores viram uma tabela filha, sem limite.

## 2.6 Como as cores se relacionam com a descrição

Testado contra as 59 linhas da amostra de `PESOS DE FIOS`, cobrindo 98 entradas de cor.
**A regra se confirma em 97 delas.**

Cada entrada de cor é um **fio usado no produto**, e a ordem das colunas reproduz a
**ordem dos fios na construção**. São dois tipos:

### (A) Fios de cor — o código está na descrição

Os códigos aparecem na descrição, **na mesma ordem**, separados por `/` quando há mais de
um:

| Descrição | cor 1 | cor 2 | cor 3 |
|---|---|---|---|
| `ATAC M10046 100 6MM` | `100` | ENCHIMENTO | |
| `ATAC 6ALF 100/460 6MM` | `100` | `460` | ENCHIMENTO |
| `ATAC M10030 278/102 6MM` | `278` | `102` | ENCHIMENTO |
| `ATAC 14163 199/100 6MM` | `199` | `100` | ENCHIMENTO |

Verificado: em **100% das linhas com duas ou mais cores**, a ordem das colunas bate com a
ordem em que os códigos aparecem na descrição.

Dois detalhes de escrita:

- **O zero à esquerda cai.** A descrição escreve o código com 3 dígitos; a coluna de cor
  guarda o número puro.

  | Descrição | Cor cadastrada |
  |---|---|
  | `ATAC 14163 0**58** 6MM` | `58` |
  | `ATAC 3000 0**04** 3MM` | `4` |
  | `ATAC 6ALF 100/0**38** 6MM` | `38` |

- **Pode haver qualificador** depois do código — mesmo código de cor, acabamento ou
  título de fio diferente, e por isso **peso diferente**:
  `102 LAVADO` · `2001 (30-2)` · `821/1 RECICLADO`

### (B) Materiais estruturais — nunca estão na descrição

Vêm sempre **depois** dos fios de cor, e não aparecem na descrição porque não são cor:
são construção.

| Material | Onde aparece |
|---|---|
| `ENCHIMENTO` | 22 ocorrências, em quase todo produto trançado |
| `BORRACHA PRETA 28` · `BORRACHA REVESTIDA PRETA` · `BORRACHA PRETA 38` | nos elásticos |

### Duas inconsistências encontradas nesta conferência

**1. Uma cor que não bate com a descrição.**

| Descrição | Cor cadastrada | Esperado pela regra |
|---|---|---|
| `FITA M12116 0**96**/LUREX PRATA 7MM` | `96` | `96` ✅ |
| `FITA M12116 **100**/LUREX PRATA 7MM` | `58` | `100` ❌ |

O produto irmão segue a regra exatamente. Tem cara de digitação trocada — vale conferir
se esse produto realmente usa a cor 58 ou se deveria ser 100.

**2. O LUREX não está sendo pesado.**

Os dois produtos acima têm `LUREX PRATA` na descrição, mas **`LUREX` nunca aparece como
fio cadastrado** em nenhuma linha da amostra. Se o lurex é um fio metálico de verdade que
entra na trama, o consumo dele não está sendo calculado em lugar nenhum.

## 2.7 Itens sem código

Confirmado pelo usuário: os itens sem código em `DADOS GERAIS DE PRODUTOS` são
**antigos e descontinuados**. Não é um defeito a corrigir — é histórico. Na migração
esses itens entram marcados como inativos, preservando o histórico sem poluir a busca.

---

## Perguntas desta parte

1. **O fator de reposição `1,25`** — os 25% a mais são perda de processo, margem de
   segurança, ou outra coisa? E ele varia por produto ou é fixo?
2. **Reposição hoje só funciona nas trançadeiras** (o ramo dos teares é inalcançável).
   Isso é uma limitação conhecida ou um defeito que passou despercebido?
3. **Um lote pode misturar clientes?** Se sim, o cálculo global de hoje já está errando
   nesses casos — vale saber se isso já foi notado.
4. **`CÓD. CLIENTE` identifica a filial** (Itapajé, Pentecoste, Projeção são filiais da
   Paquetá). Confirma que o cadastro deve ser por filial, e não por empresa?

---

# Parte 3 — A ficha técnica e a cadeia completa

## 3.1 A resposta curta

Sim, dá para ler a construção de um item de ponta a ponta. A cadeia tem **sete etapas** e
**duas chaves de ligação diferentes** — e é essa troca de chave no meio do caminho que
explica boa parte da fragilidade do sistema.

```
1  PEDIDOS                       código único = CÓD. MARFIM + TAMANHO
        │  PROCV pelo CÓDIGO
        ▼
2  DADOS GERAIS DE PRODUTOS      → referência · máquina · tamanho em metros
        │
        │  ⚠️ A CHAVE MUDA: daqui para baixo tudo se liga pela DESCRIÇÃO
        │
        ├──── PROCV pela REFERÊNCIA ────┐
        ▼                               ▼
3  PESOS DE FIOS                   4  DADOS DOS PRODUTOS
   cores + peso kg/m                  ficha técnica: espulas,
        │                             fios por espula, produção
        │                               │
        └───────────┬───────────────────┘
                    ▼
5  Programação (trançadeiras / teares)
   filtra por máquina · metros = qtd × tamanho × fator
   agrupa por referência · metros × peso = consumo kg
                    │
                    ▼
6  Planejamento                    simula período a período
                    │
                    ▼
7  RELATÓRIO                       períodos completos + sobra do último
```

## 3.2 ⚠️ A descrição do produto é usada como chave

Nas etapas 1 e 2 a ligação é pelo **código**. A partir da etapa 3, é pela **descrição**
(`ATAC 11628 100 7MM`). Ou seja: **um texto digitado à mão virou chave primária.**

Consequência direta: um espaço a mais, um zero à esquerda diferente ou uma letra trocada
na descrição **rompe silenciosamente** a ligação com os pesos de fio e com a ficha
técnica. O `PROCV` devolve `#N/A`, o `SEERRO` converte em vazio, e o consumo de fio
daquele item simplesmente não aparece.

É exatamente o mesmo mecanismo do problema da `Relação de referencias` (ver Análise), e a
mesma razão pela qual `paquetá  pentecoste` com dois espaços nunca casa.

**A própria planilha já sabe disso.** As colunas O, P e Q de `DADOS DOS PRODUTOS` são
verificadores automáticos:

```excel
O2 = SE(SEERRO(LOCALIZAR(L2;A2);0);"OK";"ERRADO")   ← a cor está na descrição?
P2 = SE(SEERRO(LOCALIZAR(M2;A2);0);"OK";"ERRADO")   ← o modelo está na descrição?
Q2 = SE(SEERRO(LOCALIZAR(N2;A2);0);"OK";"ERRADO")   ← a largura está na descrição?
```

Isso confirma, escrito na própria planilha, a regra de composição da descrição que
deduzimos na Parte 2.6: **descrição = modelo + cor + largura**.

> **Correção.** Eu havia escrito aqui que o verificador "foi desligado" em 6 linhas, onde
> aparecem `570` e `CONFERIDO 05/08/2020` no lugar de `OK`. **Isso estava errado.** Essas
> linhas são de trançadeira, e nelas as mesmas colunas guardam outra coisa: `570` é a
> **produção em metros**, não um verificador destruído. A explicação está na Parte 4 —
> a aba usa duas disposições de coluna diferentes. O verificador existe nas linhas de
> tear e simplesmente **não existe** nas de trançadeira.

## 3.3 A ficha técnica

`DADOS DOS PRODUTOS` guarda, por referência, **quais fios entram no item, em que
quantidade, e quanto a máquina produz**. Exemplo real da exportação:

```
A  ATAC 11628 100 7MM     ← referência (chave)
B  40                     ← quantidade
C  100                    ← cor
M  11628                  ← modelo
N  7MM                    ← largura
O  OK   P  OK   Q  OK     ← verificadores
```

Outro item, com espulas:

```
A  ATAC 3000 158 3MM
B  16                     ← nº de espulas
C  2                      ← fios por espula   → 16 × 2 = 32 fios
D  158                    ← cor
H  3000                   ← modelo
N  1200                   ← produção
```

O relatório consome essa ficha por `PROCV`, puxando as colunas **B a I, K a M e P**.

### Duas disposições de coluna na mesma aba

Como você disse, a aba tem dados de tear **e** de trançadeira — e as duas famílias
**usam colunas diferentes para a mesma informação**:

| Informação | Numa família | Na outra |
|---|---|---|
| Cor | `C` | `D` |
| Modelo | `M` | `H` |
| Conteúdo de `N` | largura (`7MM`) | produção (`1200`) |

Na amostra de 59 linhas: 56 na primeira disposição, 3 na segunda. É por isso que o
cabeçalho parece desalinhado — **ele descreve uma das duas disposições, não as duas**.

No modelo novo, tear e trançadeira ganham **campos próprios e nomeados**, e nada depende
de qual coluna a informação ocupa.

## 3.4 ⚠️ A ligação com `PESOS DE FIOS` é posicional, e escorregou 23 vezes

`DADOS DOS PRODUTOS` não busca a referência em `PESOS DE FIOS` — ele **copia por
posição**:

```excel
A2   = 'PESOS DE FIOS'!A2       desvio 0
A40  = 'PESOS DE FIOS'!A41      desvio +1
A43  = 'PESOS DE FIOS'!A45      desvio +2
A140 = 'PESOS DE FIOS'!A145     desvio +5
A240 = 'PESOS DE FIOS'!A251     desvio +11
...
```

São **24 trechos distintos**, com desvios de **0 a 23 — todos os inteiros da sequência**.

A leitura é direta: alguém inseriu uma linha em `PESOS DE FIOS` sem inserir a
correspondente aqui, o alinhamento quebrou dali para baixo, e a fórmula foi remendada a
partir daquele ponto. **Isso aconteceu 23 vezes.**

Enquanto os remendos estiverem certos, funciona. Mas cada nova inserção em
`PESOS DE FIOS` desalinha tudo abaixo dela, e o sintoma é a ficha técnica de um produto
aparecer associada ao **produto errado** — sem nenhum erro na tela.

No modelo novo isso desaparece: as duas informações passam a viver na **mesma tabela de
produto**, ligadas por identificador, não por posição de linha.

## 3.5 A outra metade do cálculo de tempo

Na Parte 2 vimos o `MENOR(...)`, que devolve a **sobra do último período**. Faltava a
outra metade, que está na coluna `W` do relatório:

```excel
=CONT.SE('Planejamento Trançadeira'!C30:APM30; 'Planejamento Trançadeira'!A30)
```

Conta quantas células da linha de planejamento são **iguais à capacidade** — ou seja,
**quantos períodos completos** a máquina roda antes de sobrar o resto.

```
tempo total = (períodos completos)  +  (a fração do último)
                    CONT.SE                    MENOR
```

Com a fórmula de módulo proposta na Parte 2, as duas saem de uma vez:

```js
const periodosCompletos = Math.floor(metrosTotal / capacidade);
const sobra             = metrosTotal % capacidade;
```

> Aqui também há desalinhamento: `W67` lê a linha 9 e `W193` lê a linha 30 — ambos
> batem com a regra de 6 linhas por bloco. Mas `W451` deveria ler a linha 73 e lê a
> **81**. O relatório e o planejamento também escorregaram um em relação ao outro.

## 3.6 Rastreamento completo, conferido número a número

Item: **`ATAC M15101 5334`**. Todos os valores abaixo saíram dos arquivos exportados.

### Etapas 1 e 2 — os pedidos

Três linhas em `PEDIDOS`, código `87192120CM`, todas TRANÇADEIRA, tamanho 1,2 m:
quantidades **480 + 492 + 792 = 1.764 pares**.

### Etapa 5 — metros

```
1.764 × 1,2 m × 2  =  4.233,6 m
```
`Programação trançadeiras!J` mostra **4.233,6** ✅

### Etapa 3 — consumo de fio

A cor `5334` pesa `0,0029` kg/m:
```
4.233,6 × 0,0029  =  12,27744 kg
```
`Programação trançadeiras!Q` mostra **12,27744** ✅

### Etapa 4 — ficha técnica

O relatório traz, por `PROCV` em `DADOS DOS PRODUTOS`:

| Campo | Valor |
|---|---|
| N° espulas | 24 |
| N° de fios | 2 → **48 fios** |
| Voltas na espula | 1.200 |
| Produção em metros | 470 |

### Etapa 7 — a conversão de metros para voltas

Esta era a peça que faltava. O relatório converte metros em voltas pela ficha técnica:

```
voltas = metros × (voltas na espula ÷ produção em metros)
       = 4.233,6 × 1.200 ÷ 470
       = 10.809,1914893617
```
O relatório mostra **10.809,1914893617** ✅

E o número de máquinas é a mesma conta em outra unidade:
```
Máq. = metros ÷ produção = 4.233,6 ÷ 470 = 9,00765957446808
```
O relatório mostra **9,00765957446808**, e `N° Máquinas` = **9**, a parte inteira ✅

### Etapa 6 — o planejamento

A linha 2 de `Planejamento Trançadeira` recebe capacidade `1.200` e demanda
`10.809,1914893617` — exatamente as voltas calculadas acima ✅

### Etapa 7 — a sobra

```
10.809,1914893617  mod  1.200  =  9,1914893617
```
`RELATÓRIO!H22`, a fórmula `MENOR/CONT.SE` escrita pelo Módulo17, mostra
**9,19148936170131** ✅

### As sete etapas fecham, sem divergência

```
1.764 pares
   × 1,2 m × 2     →   4.233,60 m          ✅
   × 0,0029 kg/m   →      12,27744 kg      ✅
   × 1200/470      →  10.809,19 voltas     ✅
   ÷ 470           →       9,0077 máquinas ✅
   mod 1200        →       9,1914893617    ✅
```

**Esta é a prova de que a migração é viável.** As 4,7 milhões de fórmulas se reproduzem
com cinco operações aritméticas, e o resultado bate com o da planilha até a décima casa
decimal.

## 3.7 O que "1.200" significa

A pergunta aberta da Parte 1 — *"o que é um período no planejamento?"* — tem agora uma
resposta aritmética, ainda que a interpretação física precise da sua confirmação.

O `1.200` da capacidade do planejamento é o mesmo `1.200` de **voltas na espula** da
ficha técnica. E como `voltas = metros × 1200/470`, dividir as voltas por 1.200 dá
exatamente o mesmo número que dividir os metros por 470:

```
10.809,19 ÷ 1.200  =  9,0077
 4.233,60 ÷   470  =  9,0077
```

São **a mesma grandeza em duas unidades**. O planejamento conta em voltas, o relatório
conta em metros, e `1200/470` é a taxa de conversão — que vem da ficha técnica do
produto, ou seja, **varia de produto para produto**.

Isso derruba uma suposição minha da Parte 2: eu tratei `1.200` como capacidade fixa. Ela
é fixa **por produto**, não global.

O que ainda não sei é o nome físico: cada "período" de 1.200 voltas é **uma espula
consumida** ou **uma máquina rodando um turno** que produz 470 m? O relatório rotula a
coluna como `Máq.` e `N° Máquinas`, o que sugere máquinas; mas o divisor é literalmente
as voltas de uma espula.

---

# Parte 4 — A ficha técnica serve a duas pessoas diferentes

Esta é a chave que faltava para entender `DADOS DOS PRODUTOS`. A aba não tem "duas
disposições de coluna" por desorganização: ela descreve **dois processos de fabricação
diferentes**, e cada relatório vai para **um profissional diferente**.

| Máquina | Relatório vai para | O que essa pessoa precisa saber |
|---|---|---|
| Trançadeira | **espulador** | quantas espulas encher, de cada cor |
| Tear | **urdidor** | em que ordem montar os fios na urdidura |

São necessidades distintas, e por isso a mesma coluna `B` guarda coisas distintas.

## 4.1 Trançadeira — espulas por cor

As máquinas têm **16, 32 ou 48 fusos** hoje, e novos modelos podem surgir.

Um produto pode distribuir as cores entre os fusos de qualquer maneira: uma cor pode
ocupar metade das espulas, outra apenas 1, 2 ou 3, ou uma única cor pode ocupar todas.
Por isso a ficha tem **um trio de colunas por cor**:

```
ATAC 3000 158 3MM      B=16    C=2     D=158    H=3000   N=1200   O=570
                       ↑       ↑       ↑        ↑        ↑        ↑
                    espulas  fios   cor 1    modelo   voltas   produção
                            por                       na       em metros
                          espula                    espula
```

| Coluna | Conteúdo |
|---|---|
| `B` | **nº de espulas** dessa cor |
| `C` | nº de fios em cada espula |
| `D` | a cor |
| `E`, `F`, `G` | o mesmo trio para a segunda cor |
| `H` | modelo |
| `N` | **voltas na espula** (1.200) |
| `O` | **produção em metros** (470/570) |

São `N` e `O` que alimentam a conversão descoberta na Parte 3.6:
`voltas = metros × N ÷ O`. **Elas vêm da ficha técnica, logo variam por produto.**

## 4.2 Tear — a sequência da urdidura

Aqui a coluna `B` não conta espulas: ela **descreve a estrutura do item**, como duas
listas paralelas.

```
ATAC 6ALF 212/100 6MM     B = 1 . 52 . 3        C = 212 . 100 . 212
                              └─┬─┘  └┬┘ └┬┘        └─┬─┘  └─┬─┘  └┬┘
                                1     52   3          212    100   212
```

Lê-se: **1 fio da cor 212, depois 52 fios da cor 100, depois 3 fios da cor 212.**
Total de 56 fios.

**A ordem carrega informação.** A cor `212` aparece duas vezes, em posições diferentes —
não é uma soma por cor, é a **sequência de montagem da urdidura**, da borda ao centro da
fita. É exatamente o que o urdidor precisa para montar a máquina.

Isso distingue radicalmente as duas fichas:

| | Trançadeira | Tear |
|---|---|---|
| Estrutura | **conjunto** — quantas espulas por cor | **sequência ordenada** de fios |
| A mesma cor repete? | não | **sim**, em posições diferentes |
| Colunas | um trio por cor | duas listas paralelas em `B` e `C` |

### ⚠️ A notação não é padronizada

Nas 9 linhas amostradas com sequência, **os separadores variam**:

| Produto | Quantidades | Cores | Separadores |
|---|---|---|---|
| `ATAC 6ALF 100/460 6MM` | `1 - 52 - 3` | `100 , 460 , 100` | hífen / vírgula |
| `ATAC M10030 100/222 6MM` | `10 . 4 . 10 . 4 . 14` | `100 , 222 , 100 , 222 , 100` | ponto / vírgula |
| `ATAC 14163 148/100 6MM` | `1 . 52 . 3` | `148 . 100 . 148` | ponto / ponto |

Três separadores em uso — `.`, `,` e `-` —, às vezes diferentes na mesma linha. Como é
texto livre, nada impede uma quarta variação amanhã.

O que **está** consistente: as duas listas sempre têm o mesmo número de elementos (9 de 9
na amostra). A estrutura é sólida; só a escrita é que não é.

Há ainda divergência de zero à esquerda entre abas: `ATAC 6ALF 100/038 6MM` registra a
cor como `038` aqui e como `38` em `PESOS DE FIOS`.

## 4.3 ⚠️ A coluna "Confirmado" do relatório está quebrada em 645 células

```excel
=SE($B$19="";"";PROCV($B$19;'DADOS DOS PRODUTOS'!$A$1:$O$9999;16;0))
```

Dois defeitos na mesma fórmula:

1. **O intervalo `A:O` tem 15 colunas, e a fórmula pede a 16ª.** Isso devolve erro
   sempre, em toda linha. Confirmado na amostra: a coluna `S` do relatório mostra `#ERRO`
   em **todas** as linhas conferidas.
2. **`$B$19` está travado na linha 19.** As outras fórmulas da mesma família usam `$B19`,
   que acompanha a linha. Mesmo que o intervalo fosse corrigido, as 645 células buscariam
   todas **o mesmo produto** — o da linha 19 — em vez de cada uma o seu.

A coluna provavelmente pretendia trazer o `OK`/`ERRADO` do verificador de
`DADOS DOS PRODUTOS`, que fica na coluna `O` (índice 15). Trocar `16` por `15` e `$B$19`
por `$B19` faria a checagem de consistência aparecer no relatório — que é justamente
onde ela seria útil.

## 4.4 O que isso muda no sistema novo

A ficha técnica deixa de ser "colunas que significam coisas diferentes conforme a linha"
e passa a ter **duas formas explícitas**, escolhidas pelo tipo de máquina:

**Trançadeira** — lista de cores, cada uma com sua quantidade de espulas:

| cor | espulas | fios por espula |
|---|---|---|
| 158 | 16 | 2 |

Mais os campos `voltas_na_espula` e `producao_metros`, que hoje moram em `N` e `O`.

**Tear** — lista **ordenada** de segmentos:

| ordem | cor | fios |
|---|---|---|
| 1 | 212 | 1 |
| 2 | 100 | 52 |
| 3 | 212 | 3 |

Vira uma tabela de verdade, com ordem explícita. Some o texto livre, somem os três
separadores, e o total de fios passa a ser calculado — não digitado.

E como as máquinas de 16, 32 e 48 fusos podem mudar, o número de fusos vira **cadastro de
máquina**, não número fixo em fórmula. Um modelo novo de 64 fusos entra como um registro,
sem tocar em cálculo nenhum.

---

# Parte 5 — O alinhamento das espulas

## 5.1 O problema físico

Numa trançadeira, cada espula pode carregar **1, 2, 3, 4 ou mais fios**. Um mesmo produto
costuma misturar: parte das espulas com 2 fios, parte com 1.

Espulas com quantidades diferentes de fios **esvaziam em ritmos diferentes**. Se todas
forem enchidas com o mesmo número de voltas, umas acabam antes das outras, e a máquina
para para troca parcial.

A solução é encher cada grupo com um número de voltas **diferente e proporcional**, de
modo que **todas terminem juntas**. Quantas voltas cada grupo precisa é algo que
**se mede na produção** — não se calcula.

## 5.2 A fórmula que faz isso

Na coluna `H` do `RELATÓRIO TRANÇADEIRAS`, em **516 células**:

```excel
=SE(M25=$C$2;QUOCIENTE(H25;0,833);
 SE(M25=$C$3;MULT(H25;1,235);
 ... 12 modelos ...
 H25))
```

Ela lê o **modelo** do produto (coluna `M`), procura na lista `$C$2:$C$13` e aplica o
fator correspondente ao valor da **linha de cima**. É uma cascata.

### Conferido com dados reais — `ATAC M15101 5334`, fator 1,25

| Linha | Espulas | Fios | Voltas | |
|---|---|---|---|---|
| 25 | 24 | **2** | **1.200** | base |
| 26 | 24 | **1** | 1.500 | ×1,25 |
| 27 | 0 | 0 | 1.875 | ×1,25 |
| 28 | 24 | **2** | **9,19148936170131** | a sobra (do `MENOR`) |
| 29 | 24 | 1 | 11,4893617021266 | ×1,25 |
| 30 | 0 | 0 | 14,3617021276583 | ×1,25 |

Isso explica finalmente o **bloco de 6 linhas**: são **dois grupos de 3**. O primeiro
trata das máquinas que rodam cheias; o segundo, da última máquina, que roda só a sobra.
Dentro de cada grupo, uma linha por conjunto de espulas.

O espulador lê: *"encha 24 espulas de 2 fios com 1.200 voltas, e 24 espulas de 1 fio com
1.500 voltas — assim acabam juntas."*

### O fator é medido, não deduzido

Se ele fosse a razão entre os fios, 2 fios → 1 fio daria **×2** (2.400 voltas). A planilha
usa **×1,25** (1.500). Confirma que o valor vem de medição na produção, como descrito.

## 5.3 Quatro fragilidades nessa construção

### a) Metade dado, metade código

A **lista de modelos** está em células (`C2:C13`), mas os **fatores** estão dentro da
fórmula. Incluir um 13º modelo exige editar o `SE` aninhado em **516 células**.

### b) Modelo fora da lista passa sem ajuste, em silêncio

O último argumento é `H25` — sem fator. Um produto cujo modelo não esteja entre os 12
recebe **as mesmas voltas em todos os grupos de espulas**, que é exatamente a situação
que a fórmula existe para evitar. E nada na tela indica isso.

O usuário reconhece que a lista pode estar incompleta ("se tem fórmulas faltando é erro
meu no preenchimento"). **É por isso que o silêncio é o problema, não a lacuna:** um
cadastro incompleto deveria avisar, não produzir número plausível e errado.

### c) `QUOCIENTE` onde deveria haver `MULT`

`1/0,833 = 1,2005`. Ou seja, **M6034 e M13745 significam a mesma coisa: ×1,2.** Mas
M6034 usa `QUOCIENTE`, que é divisão inteira e trunca:

| Voltas | `QUOCIENTE(H;0,833)` | `MULT(H;1,2)` | Perda |
|---|---|---|---|
| 1.200 | 1.440 | 1.440,00 | — |
| 9,19148936 | **11** | 11,0298 | 0,0298 |
| 838,468085 | **1.006** | 1.006,1617 | 0,1617 |

Nas máquinas cheias não faz diferença. Nas **linhas de sobra**, que são fracionárias por
natureza, ele descarta a parte decimal — sempre para menos — e o erro se propaga pelas
duas multiplicações seguintes.

### d) A terceira linha calcula para um grupo que não existe

No exemplo, o terceiro conjunto tem **0 espulas e 0 fios**, mas a cascata ainda produz
`1.875` e `14,36`. São números sem significado num relatório que vai para o chão de
fábrica.

## 5.4 Pergunta em aberto

O fator está ligado ao **modelo** (`M15101`), não à referência. Isso funciona enquanto
todos os produtos de um mesmo modelo tiverem a mesma distribuição de espulas e fios — o
que é verdade na amostra, onde todo `M15101` usa 24×2 + 24×1.

**Isso é garantido?** Se dois produtos do mesmo modelo puderem ter distribuições
diferentes, eles precisariam de fatores diferentes e hoje receberiam o mesmo.

---

## 5.5 Anatomia completa do bloco (lida das fórmulas)

Com as fórmulas à vista, o bloco de 6 linhas fica inteiramente decifrado.
Notação: `C`=Metros · `D`=Voltas · `E`=Máq. · `G`=N° Máquinas · `H`=Voltas ·
`J`=N° espulas · `Q`=voltas na espula · `R`=produção em metros · `W`=`CONT.SE`.

### As três colunas de cabeçalho do item

```excel
C25 = 'Programação trançadeiras'!J4      ← metros totais da referência
D25 = (Q25*C25)/R25                      ← voltas
E25 = D25/Q25                            ← máquinas
```

`D = (Q × C) ÷ R` é **literalmente** a conversão deduzida na Parte 3.6
(`voltas = metros × voltas_na_espula ÷ produção`). Conferido: `(1200 × 4233,6) ÷ 470 =
10.809,1914893617`, e `E = D ÷ Q = 9,00765957446808`. Ambos batem com a planilha.

### As 6 linhas

| Linha | `G` — N° Máquinas | `H` — Voltas |
|---|---|---|
| 25 | `=SE(H25<Q25;1;W25)` | `=SE(D25>Q25;Q25;SE(D25<Q25;D25;D25))` |
| 26 | `=SE(J26>0;G25;" ")` | `=SE(M25=$C$2;QUOCIENTE(H25;0,833);…)` |
| 27 | `=SE(J27>0;G26;" ")` | cascata sobre `H26` |
| 28 | `=SE(H28>0;1;" ")` | `=MENOR('Planejamento Trançadeira'!C2:FJ2;CONT.SE(…)+1)` |
| 29 | `=SE(J29>0;G28;" ")` | cascata sobre `H28` |
| 30 | `=SE(J30>0;G29;" ")` | cascata sobre `H29` |

Ou seja:

- **Linha 25** — as máquinas cheias. `H25` é o **mínimo entre as voltas totais e a
  capacidade da espula**: se o pedido inteiro couber em menos de uma espula, usa o
  próprio pedido; senão, enche a espula. `G25` é `1` quando o pedido cabe numa espula,
  e `W25` (o `CONT.SE`, número de períodos completos) quando não cabe.
- **Linha 28** — a última máquina, com a sobra do `MENOR`. `G28` é sempre `1`.
- **Linhas 26, 27, 29, 30** — os demais grupos de espulas, com a cascata de fator.

> **Correção.** Na Parte 5.2 eu descrevi `H25 = 1.200` como "o valor base vindo da ficha".
> Está mais preciso dizer que é `MÍNIMO(voltas totais; voltas na espula)`. No exemplo dá
> 1.200 porque o pedido é grande; num pedido pequeno daria o próprio pedido.

### Duas simplificações à vista

**1. Um `SE` com os dois ramos idênticos.**

```excel
=SE(D25>Q25; Q25; SE(D25<Q25; D25; D25))
                        ↑ verdadeiro e falso devolvem a MESMA coisa
```

O `SE` interno é inútil: some qualquer que seja a comparação. A fórmula equivale a
`=MÍNIMO(D25;Q25)`.

**2. A guarda existe numa coluna e falta na outra.**

`G26 = SE(J26>0; G25; " ")` só mostra o número de máquinas **se o grupo tiver espulas**.
A coluna `H` não tem guarda equivalente — por isso, na linha 27 do exemplo,
`G` sai em branco (grupo vazio, corretamente) mas `H` mostra **1.875 voltas** para um
grupo de **0 espulas**.

A lógica de "não mostrar grupo vazio" já existe na planilha; só não foi aplicada à
coluna que mais importa para o espulador.

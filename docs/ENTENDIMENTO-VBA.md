# O que os 18 módulos VBA fazem

Análise dos arquivos `Módulo1.bas` … `Módulo18.bas` que estão na raiz do repositório.

**Conclusão principal:** dos 18 módulos, apenas **4** carregam regra de negócio de
verdade. Os outros 14 são utilitários de interface do Excel ou macros gravadas com o
gravador — e a maioria deles simplesmente deixa de existir num sistema web, porque o
problema que resolviam é do Excel, não do negócio.

> Os nomes de abas aparecem corrompidos nos arquivos (`RELAT�RIO TRAN�ADEIRAS`) porque o
> VBA grava `.bas` em ANSI. Aqui eles estão escritos corretamente.

---

## Resumo por módulo

| Módulo | O que faz | Destino |
|---|---|---|
| 1 | Salva a planilha automaticamente a cada 10 min | **Descartar** — banco de dados salva sozinho |
| 2 | `JuntarDados`: concatena coluna A + coluna H | **Migrar** (regra de dados) |
| 3 | `CopiarNaoRepetidos`: lista única de J para N | **Substituir** — é um `DISTINCT` |
| 4 | `CopiarNaoRepetidos1`: variante da anterior, colunas C/A | **Substituir** — e contém um bug, veja abaixo |
| 5 | Seleciona a última célula da coluna C | **Descartar** — navegação |
| 6 | Seleciona a última célula da coluna B | **Descartar** — navegação |
| 7 | `Macro1/2/3`: macros gravadas; só a 3 faz algo | **Descartar** (ver ressalva) |
| 8 | `Juntar`: concatena G + H quando H contém "mm" | **Migrar** (regra de dados) |
| 9 | Oculta linhas por intervalo digitado | **Substituir** — filtro da interface |
| 10 | Oculta colunas por intervalo digitado | **Substituir** — filtro da interface |
| 11 | Reexibe todas as colunas | **Substituir** — filtro da interface |
| 12 | Versão otimizada do Módulo10 | **Substituir** — duplicata do 10 |
| 13 | `APAGAR`: limpa `A4:I` inteira | **Substituir** — "limpar" da interface |
| 14 | Autoajuste de altura das linhas 19:1266 | **Descartar** — layout do Excel |
| 15 | Insere fórmulas ligando Planejamento → Relatório | **Migrar** (mapeamento estrutural) |
| 16 | Insere fórmulas ligando Programação → Relatório | **Migrar** (mapeamento estrutural) |
| 17 | Insere a fórmula `MENOR/CONT.SE` na coluna H | **Migrar** — é o coração do cálculo |
| 18 | Copia H20:H21 a cada 3 linhas | **Migrar** com cuidado — conflita com o 17 |

---

## Os módulos que importam

### Módulo 17 — o coração do sistema

```vba
wsRelatorio.Cells(linhaRelatorio, 8).Formula = _
    "=MENOR('Planejamento Trançadeira'!C" & linhaPlanejamento & ":FJ" & linhaPlanejamento & _
    ",CONT.SE('Planejamento Trançadeira'!C" & linhaPlanejamento & ":FJ" & linhaPlanejamento & ",0)+1)"
```

Percorre o relatório da **linha 22 até a 790, de 6 em 6**, e para cada uma pega a linha
correspondente do planejamento (1, 2, 3, …).

São exatamente **129 células** (`22, 28, 34, … 790`).

A fórmula `MENOR(intervalo; CONT.SE(intervalo;0)+1)` é um idioma clássico do Excel para
**"o menor valor da linha, ignorando os zeros"**. Funciona assim: `CONT.SE(...;0)` conta
quantos zeros existem; se há 5 zeros, `MENOR(...; 6)` devolve o 6º menor valor — ou seja,
o primeiro que não é zero.

Traduzindo para o negócio: **para cada máquina, qual é a próxima data/ordem válida**,
desconsiderando os períodos vazios. Em SQL seria `MIN(valor) WHERE valor <> 0`; em
JavaScript, `Math.min(...linha.filter(v => v !== 0))`.

O intervalo `C:FJ` vai da coluna 3 à coluna 166 — **164 colunas por máquina**. Descobrir
o que essas colunas representam (dias? turnos? ordens?) é a pergunta central da rodada 2.

### Módulos 15 e 16 — o mapa entre as abas

Não calculam nada; apenas **conectam as abas**. Juntos revelam a topologia do sistema:

| Origem | Destino | Regra |
|---|---|---|
| `RELATÓRIO!D{19 + (i-1)*6}` | `Planejamento!B{i}` | Módulo15, i de 1 a 137 |
| `Programação!J{(i-19)/6 + 3}` | `RELATÓRIO!C{i}` | Módulo16, i de 19 a 787, passo 6 |

Ou seja, o fluxo é circular:

```
Programação trançadeiras ──J──> RELATÓRIO (col. C)
                                    │
                                    D
                                    ↓
                        Planejamento Trançadeira (col. B)
                                    │
                                 C:FJ
                                    ↓
                              RELATÓRIO (col. H)   ← Módulo17
```

**Um bloco de 6 linhas no relatório = 1 linha no planejamento = 1 máquina.** Os três
módulos concordam nesse ponto, o que dá boa confiança na leitura.

Note a divergência de contagem: o Módulo15 vai até 137 máquinas, o 16 e o 17 vão até
~129. Isso sugere que o planejamento tem folga para máquinas ainda não cadastradas — ou
que um dos laços está desatualizado. **Vale confirmar.**

### Módulos 2 e 8 — regras de concatenação

- **Módulo2** faz `A = A & H` para toda linha a partir da 4.
- **Módulo8** faz `G = G & " " & H`, **mas só quando H contém "mm"**.

O filtro por `"mm"` é claramente uma regra de negócio: parece juntar uma **medida em
milímetros** (bitola? diâmetro?) à descrição do produto. Num sistema web isso vira campo
calculado, e nem precisa ser gravado — dá para exibir na hora.

Atenção: ambos são **destrutivos e não idempotentes**. Rodar duas vezes gera
`ABCABC` e `descrição 5mm 5mm`. É o tipo de armadilha que o sistema novo elimina de
graça, mantendo os campos separados e concatenando só na exibição.

### Módulo 18 — conflito com o Módulo 17

Copia `H20:H21` e cola **a cada 3 linhas**, de 23 até 791. Mas o Módulo17 escreve na
mesma coluna H **a cada 6 linhas**. Os dois não podem estar corretos ao mesmo tempo.

Provavelmente um substituiu o outro e o antigo ficou para trás. **Preciso saber qual dos
dois você usa hoje** — o arquivo `10_FORMULAS_*.txt` vai responder isso sozinho, mostrando
o que de fato está na coluna H no momento.

---

## Ponto de atenção no Módulo 4

```vba
Set celulaDestino = Range("c2")
For Each celula In Range("c2:c" & ...)
    If WorksheetFunction.CountIf(Range("a2:a" & celulaDestino.Row - 1), celula.Value) = 0 Then
        celulaDestino.Value = celula.Value
```

Ele **lê** da coluna C, **verifica duplicidade** na coluna A, mas **escreve** na coluna C
— sobrescrevendo a própria fonte enquanto a percorre. O `CopiarNaoRepetidos` original
(Módulo3) é coerente: lê de J, verifica em N, escreve em N.

Isso tem toda a cara de uma cópia do Módulo3 adaptada pela metade. Provavelmente é código
morto. Não vou reproduzir esse comportamento no sistema novo sem você confirmar que ele é
usado — se for, o correto é a versão do Módulo3.

---

## Ressalva sobre o Módulo 7

`Macro1` e `Macro2` só fazem seleções e rolagem de tela: não têm efeito nenhum. Já a
`Macro3` termina com `Selection.ClearContents` sobre `A4:I2423` e `N4:N2423` — **apaga
dados de verdade**. É praticamente a mesma coisa que o `APAGAR` do Módulo13, com um
intervalo um pouco diferente e um nome que não avisa nada.

Vale saber se alguém usa `Macro3` no dia a dia, porque o nome não deixa óbvio que ela
limpa duas faixas de dados.

---

## O que isso significa para a migração

1. **A lógica de negócio no VBA é pequena** — cabe em um parágrafo: junte campos, encontre
   o menor valor não-zero por máquina, e mantenha o mapeamento de 6 linhas por bloco.
2. **A lógica de verdade está nas fórmulas das células**, que ninguém consegue ver ainda.
   Por isso a exportação vem antes de qualquer código.
3. **Cerca de 12 dos 18 módulos desaparecem** na migração. Ocultar colunas, ajustar altura
   de linha, salvar automaticamente, navegar até a última célula — nada disso é problema
   num sistema web. Isso reduz bastante o tamanho real do trabalho.
4. **Duas inconsistências já apareceram** (Módulo17 × Módulo18, e o Módulo4) só de ler o
   código. É provável que a exportação revele mais. Migrar é uma boa oportunidade para
   resolvê-las em vez de reproduzi-las.

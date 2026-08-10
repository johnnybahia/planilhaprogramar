# Como exportar a planilha para o Git

A planilha tem ~36 MB e não cabe no GitHub. Mas o que importa dela — **as fórmulas e a
estrutura** — cabe em poucas centenas de KB de texto. Este guia mostra como extrair isso
em cerca de 5 minutos.

Você não precisa instalar nada.

---

## Passo 1 — Importar a macro

1. Abra a sua planilha no Excel.
2. Pressione **`Alt + F11`** (abre o Editor do VBA).
3. No menu: **Arquivo ▸ Importar Arquivo…**
4. Escolha o arquivo **`ferramentas/ExportarPlanilha.bas`** deste repositório.
5. Feche o Editor do VBA (`Alt + Q`).

> Se o Excel bloquear macros, vá em **Arquivo ▸ Opções ▸ Central de Confiabilidade ▸
> Configurações da Central de Confiabilidade ▸ Configurações de Macro** e habilite as
> macros. Como o arquivo veio da internet, pode ser necessário também clicar com o botão
> direito no `.bas`, ir em **Propriedades** e marcar **Desbloquear**.

## Passo 2 — Rodar

1. Pressione **`Alt + F8`**.
2. Selecione **`ExportarTudo`** e clique em **Executar**.
3. Aguarde. Em uma planilha desse tamanho leva de alguns segundos a poucos minutos.
4. No final aparece uma janela com o caminho da pasta e a lista de arquivos gerados.

**A macro não altera nada na sua planilha.** Ela só lê. Não escreve em células, não apaga
nada e nunca salva o arquivo. Se quiser conferir antes de rodar, o código está comentado
em português.

## Passo 3 — Conferir a privacidade

Antes de subir para o Git, **abra os arquivos `50_AMOSTRA_*.csv`**. Eles contêm as
primeiras 60 linhas de dados reais de cada aba e podem incluir nomes de clientes, números
de pedido ou qualquer outra informação da empresa.

Se preferir não expor dado nenhum:

1. `Alt + F11`, abra o módulo `ExportarPlanilha`.
2. No topo, troque:
   ```vba
   Private Const EXPORTAR_AMOSTRA As Boolean = True
   ```
   por `False`.
3. Rode `ExportarTudo` de novo.

Você perde um pouco de contexto (eu deixo de ver o formato real dos dados), mas as
fórmulas e a estrutura continuam sendo exportadas normalmente.

## Passo 4 — Subir para o repositório

1. Copie **todo o conteúdo** da pasta `_exportacao` para a pasta **`docs/planilha/`**
   deste repositório.
2. Faça commit e push.
3. Me avise que subiu.

O `.gitignore` já bloqueia arquivos `.xlsm` / `.xlsx`, então não há risco de o arquivo de
36 MB entrar no repositório por acidente.

---

## O que cada arquivo gerado contém

| Arquivo | Para que serve |
|---|---|
| `00_RESUMO.md` | Visão geral: abas, tamanhos, quantas fórmulas cada uma tem, diagnóstico de inchaço. **Comece por aqui.** |
| `10_FORMULAS_*.txt` | **O arquivo mais importante.** Todas as fórmulas, agrupadas por padrão. |
| `20_NOMES.csv` | Intervalos nomeados. |
| `30_VALIDACOES.csv` | Listas suspensas e regras de entrada de dados. |
| `40_FORMATACAO_CONDICIONAL.csv` | As regras por trás das cores (atrasado, no prazo, parado). |
| `50_AMOSTRA_*.csv` | Primeiras 60 linhas de dados reais de cada aba. |
| `60_LIGACOES.txt` | Vínculos externos e quais abas alimentam quais. |
| `70_OBJETOS.txt` | Botões e a macro que cada um dispara, gráficos, tabelas dinâmicas. |
| `80_ESTRUTURA.txt` | Células mescladas, linhas/colunas ocultas, larguras — o layout. |
| `99_ARQUIVOS.txt` | Lista do que foi gerado, com tamanhos. |

### Por que os arquivos de fórmula são pequenos

A aba de planejamento pode ter mais de 100 mil células com fórmula. Despejar uma por
linha geraria um arquivo inutilizável.

A macro usa **notação R1C1**, na qual uma fórmula copiada para baixo é *idêntica* à
original. Agrupando por esse padrão, as 100 mil fórmulas viram algumas dezenas de blocos
como este:

```
----------------------------------------------------------------
Ocorrencias : 129
Primeira    : H22
Ultima      : H790
Distribuicao: uma a cada 6 linhas, de 22 a 790
R1C1        : =SMALL('Planejamento Trançadeira'!R[-21]C3:R[-21]C166;COUNTIF(...)+1)
A1-PT       : =MENOR('Planejamento Trançadeira'!C1:FJ1;CONT.SE('Planejamento Trançadeira'!C1:FJ1;0)+1)
A1-EN       : =SMALL('Planejamento Trançadeira'!C1:FJ1,COUNTIF('Planejamento Trançadeira'!C1:FJ1,0)+1)
```

Cada bloco traz três versões da mesma fórmula, e cada uma serve para uma coisa:

- **R1C1** — mostra o padrão, revelando que é a mesma fórmula repetida.
- **A1-PT** — como você vê no Excel, em português.
- **A1-EN** — com os nomes de função em inglês. **É esta que serve para o Google
  Planilhas**, que usa `SMALL` e `COUNTIF`, não `MENOR` e `CONT.SE`.

A linha `Distribuicao` é a que revela a estrutura do relatório: "uma a cada 6 linhas"
significa que cada máquina ocupa um bloco de 6 linhas.

---

## Por que a planilha tem 36 MB (e como diminuir)

Quase sempre a causa é o **intervalo usado inflado**: em algum momento uma formatação ou
uma seleção alcançou linhas lá embaixo, e o Excel passou a guardar centenas de milhares
de linhas vazias. Há indício disso no próprio VBA atual — o `Módulo13` faz:

```vba
Range("A4:I1048576").Select
```

Isso toca a última linha da planilha inteira.

**Como confirmar:** no `00_RESUMO.md`, compare as colunas *Intervalo usado* e *Última
célula real*. Se uma aba diz que usa até a linha 500.000 mas o último dado real está na
linha 800, é exatamente esse o problema. A seção *Diagnóstico de tamanho* do resumo já
lista as abas suspeitas automaticamente.

**Como corrigir** (faça uma cópia de segurança antes):

1. Vá até a aba com o problema.
2. Clique na linha logo abaixo do último dado real.
3. `Ctrl + Shift + ↓` para selecionar até o fim.
4. Botão direito ▸ **Excluir** (excluir a linha inteira, não apenas limpar o conteúdo).
5. Repita para as colunas: `Ctrl + Shift + →` ▸ Excluir.
6. **Salve e feche o arquivo.** O intervalo usado só é recalculado ao reabrir.

É comum um arquivo cair de 36 MB para 2–5 MB assim. Isso não é obrigatório para a
migração, mas deixa a planilha muito mais rápida enquanto ela ainda estiver em uso.

---

## Se der erro

A macro mostra uma janela com o número e a descrição do erro, além da pasta onde os
arquivos que deram certo foram gravados. **Me mande essa mensagem e os arquivos que
foram gerados** — mesmo uma exportação parcial já ajuda bastante.

Dois casos conhecidos:

- **Planilha no OneDrive ou SharePoint.** Nesse caso o Excel não consegue criar a pasta
  ao lado do arquivo. A macro detecta isso sozinha e grava na sua **Área de Trabalho**,
  numa pasta chamada `_exportacao`.
- **Erro de permissão ao criar a pasta.** Mova uma cópia da planilha para uma pasta local
  (por exemplo `C:\temp`) e rode a macro a partir dali.

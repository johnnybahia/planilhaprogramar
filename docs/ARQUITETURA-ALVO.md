# Arquitetura alvo

Documento curto para registrar as decisões já tomadas e as perguntas que ainda não têm
resposta. **Nenhum código de sistema foi escrito ainda** — isso é assunto da rodada 2.

---

## Decisões

| Tema | Decisão |
|---|---|
| Banco de dados | **Google Planilhas** |
| Back-end | **Google Apps Script** (API entre o front e a planilha) |
| Front-end | **HTML + CSS + JavaScript** |
| Usuários | **2 a 5 pessoas** |
| Hospedagem | Apps Script Web App (URL pronta, sem servidor para administrar) |

### Por que Apps Script

Você pediu "app web com servidor, vou usar o Google Planilhas". Google Planilhas sozinho
é um banco de dados — não atende requisições HTTP nem serve páginas. O Apps Script é a
peça que a Google oferece para isso: um script que roda na infraestrutura dela, lê e
escreve na planilha e responde a chamadas do navegador.

Na prática:

```
Navegador (HTML + JS)
        │  fetch / google.script.run
        ▼
Google Apps Script  ← o "servidor"
        │  SpreadsheetApp
        ▼
Google Planilhas    ← o banco de dados
```

O que isso te dá, sem custo e sem administrar servidor:

- Login pelo Google que a empresa já usa, com controle de quem acessa.
- Várias pessoas ao mesmo tempo, sem "arquivo bloqueado por outro usuário".
- Histórico de versões automático.
- Os dados continuam numa planilha, que você pode abrir e conferir a qualquer momento.

**Limites que precisamos respeitar** (relevantes para o dimensionamento):

- ~10 milhões de células por planilha — folgado para 137 máquinas × 164 colunas ≈ 22 mil.
- 6 minutos por execução de script — exige processar em lotes se um recálculo for pesado.
- Leitura/escrita em massa é rápida; **célula a célula é lento**. O código precisa usar
  `getValues()` / `setValues()` em blocos, não em laço.

> Se em algum momento isso apertar, o caminho natural é trocar só a camada de banco
> (Supabase, por exemplo) mantendo o front-end. Vale registrar que essa porta fica aberta.

### Uma diferença importante de fórmulas

O Google Planilhas usa os nomes de função **em inglês** e vírgula como separador:

| Excel (pt-BR) | Google Sheets |
|---|---|
| `MENOR(intervalo; k)` | `SMALL(range, k)` |
| `CONT.SE(intervalo; criterio)` | `COUNTIF(range, criterion)` |
| `PROCV` | `VLOOKUP` |
| `SE` | `IF` |

Por isso a macro exportadora grava **as duas versões** de cada fórmula (`A1-PT` e
`A1-EN`). A versão em inglês é a que vai ser usada de fato.

---

## Perguntas que só a exportação pode responder

Nenhuma delas dá para responder lendo o VBA. São o motivo de a extração vir primeiro.

1. **O que são as 164 colunas (C:FJ) do planejamento?**
   Dias corridos? Turnos? Ordens de produção enfileiradas? Isso decide o modelo de dados
   inteiro — é a pergunta mais importante da lista.

2. **De onde vêm os dados de entrada?**
   Alguém digita? Vem colado de um ERP? É importado de outro arquivo?
   O `60_LIGACOES.txt` mostra se há vínculos externos.

3. **O que a aba `Programação trançadeiras` calcula?**
   Ela alimenta a coluna C do relatório, mas nenhum módulo VBA revela sua lógica.

4. **São 129 ou 137 máquinas?**
   O Módulo15 vai até 137; os Módulos 16 e 17 param em ~129. Um dos dois está
   desatualizado, ou existe folga proposital.

5. **Qual regra está valendo na coluna H — o Módulo17 ou o Módulo18?**
   Eles se contradizem (a cada 6 linhas × a cada 3 linhas). O arquivo de fórmulas mostra
   o que realmente está na planilha hoje.

6. **O que as cores significam?**
   O `40_FORMATACAO_CONDICIONAL.csv` costuma revelar regras que ninguém documentou:
   atrasado, no prazo, máquina parada, em manutenção.

7. **Por que 6 linhas por máquina?**
   Seis linhas para quê? Turnos? Etapas do processo? Produtos diferentes na mesma máquina?

---

## Fases previstas

| Fase | Conteúdo | Situação |
|---|---|---|
| 1 | Extrair e documentar a planilha | **Esta rodada** |
| 2 | Ler os arquivos exportados, modelar os dados e validar com você | A seguir |
| 3 | Montar a planilha Google que serve de banco + Apps Script de leitura | |
| 4 | Front-end: tela de planejamento e tela de relatório | |
| 5 | Migrar os dados reais e rodar em paralelo com o Excel | |
| 6 | Desligar a planilha | |

A fase 5 importa mais do que parece: rodar os dois em paralelo por algumas semanas é o
que permite **comparar os números** e provar que o sistema novo dá o mesmo resultado
antes de alguém depender dele.

Attribute VB_Name = "ExportarPlanilha"
Option Explicit

'==============================================================================
'  ExportarPlanilha
'  ---------------------------------------------------------------------------
'  Extrai a ESTRUTURA e a LOGICA da planilha para arquivos de texto pequenos,
'  que cabem no Git. O arquivo .xlsm original nao precisa ser versionado.
'
'  IMPORTANTE: esta macro e SOMENTE LEITURA.
'  Ela nao altera nenhuma celula, nao apaga nada e nunca salva a planilha.
'
'  COMO USAR
'    1. Alt+F11  ->  Arquivo ->  Importar Arquivo...  ->  ExportarPlanilha.bas
'    2. Alt+F8   ->  executar "ExportarTudo"
'    3. Ao final aparece uma mensagem com o caminho da pasta gerada.
'
'  POR QUE OS COMENTARIOS NAO TEM ACENTO
'    O editor do VBA le arquivos .bas como ANSI. Um arquivo salvo em UTF-8 chega
'    com a acentuacao corrompida (e o que aconteceu com os Modulo*.bas deste
'    repositorio, que aparecem como "M?dulo1"). Mantendo o codigo em ASCII puro
'    o modulo importa corretamente em qualquer maquina. Os arquivos GERADOS pela
'    macro sao gravados em UTF-8 de verdade, com acentuacao correta.
'==============================================================================


'------------------------------------------------------------------------------
' CONFIGURACAO - ajuste aqui se precisar
'------------------------------------------------------------------------------

' Quantas linhas de dados reais exportar por aba (arquivos 50_AMOSTRA_*).
Private Const LINHAS_AMOSTRA As Long = 60

' Quantas colunas exportar na amostra. 0 = todas as colunas realmente usadas.
Private Const COLUNAS_AMOSTRA As Long = 0

' False = nao exporta dado nenhum, apenas estrutura e formulas.
' Use False se a planilha tiver nomes de clientes ou informacao sensivel.
Private Const EXPORTAR_AMOSTRA As Boolean = True

' Separador dos arquivos .csv. ";" e o padrao brasileiro.
Private Const SEPARADOR As String = ";"

' Nome da pasta de saida, criada ao lado da planilha.
Private Const NOME_PASTA As String = "_exportacao"

' Limite de celulas lidas de uma vez (protege a memoria em abas gigantes).
Private Const MAX_CELULAS_BLOCO As Long = 100000

' Acima deste numero de padroes distintos a ordenacao e pulada (seria lenta).
Private Const MAX_PADROES_ORDENAR As Long = 4000

' Ate quantas linhas procurar celulas mescladas (o layout fica no topo da aba).
Private Const MAX_LINHAS_MESCLA As Long = 200


'------------------------------------------------------------------------------
' Estado interno
'------------------------------------------------------------------------------
Private mBuf() As String
Private mBufN As Long
Private mArquivos As Object      ' Dictionary: nome do arquivo -> tamanho em bytes
Private mPasta As String


'==============================================================================
' PONTO DE ENTRADA
'==============================================================================
Public Sub ExportarTudo()

    Dim calcAnterior As Long
    Dim eventosAnterior As Boolean
    Dim telaAnterior As Boolean
    Dim t0 As Single

    t0 = Timer

    calcAnterior = Application.Calculation
    eventosAnterior = Application.EnableEvents
    telaAnterior = Application.ScreenUpdating

    On Error GoTo Falha

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set mArquivos = CreateObject("Scripting.Dictionary")

    mPasta = PrepararPasta()
    If Len(mPasta) = 0 Then GoTo Limpeza

    Application.StatusBar = "Exportando: resumo..."
    GerarResumo

    Application.StatusBar = "Exportando: formulas..."
    GerarFormulas

    Application.StatusBar = "Exportando: nomes definidos..."
    GerarNomes

    Application.StatusBar = "Exportando: validacoes de dados..."
    GerarValidacoes

    Application.StatusBar = "Exportando: formatacao condicional..."
    GerarFormatacaoCondicional

    If EXPORTAR_AMOSTRA Then
        Application.StatusBar = "Exportando: amostras de dados..."
        GerarAmostras
    End If

    Application.StatusBar = "Exportando: ligacoes..."
    GerarLigacoes

    Application.StatusBar = "Exportando: objetos e botoes..."
    GerarObjetos

    Application.StatusBar = "Exportando: estrutura visual..."
    GerarEstrutura

    GerarListaArquivos

    MsgBox Relatorio(t0), vbInformation, "Exportacao concluida"
    GoTo Limpeza


Falha:
    MsgBox "Ocorreu um erro durante a exportacao." & vbCrLf & vbCrLf & _
           "Numero: " & Err.Number & vbCrLf & _
           "Descricao: " & Err.Description & vbCrLf & vbCrLf & _
           "Os arquivos gerados ate aqui estao em:" & vbCrLf & mPasta & vbCrLf & vbCrLf & _
           "Envie esta mensagem junto com os arquivos que conseguiu gerar.", _
           vbCritical, "Erro na exportacao"

Limpeza:
    On Error Resume Next
    Application.StatusBar = False
    Application.Calculation = calcAnterior
    Application.EnableEvents = eventosAnterior
    Application.ScreenUpdating = telaAnterior

End Sub


'==============================================================================
' PASTA DE SAIDA
'==============================================================================
Private Function PrepararPasta() As String

    Dim base As String
    Dim alvo As String

    base = ThisWorkbook.Path

    ' Planilha no OneDrive / SharePoint devolve uma URL https, e MkDir falha nela.
    ' Nesse caso a saida vai para a area de trabalho.
    If Len(base) = 0 Or LCase$(Left$(base, 4)) = "http" Then
        base = Environ$("USERPROFILE") & "\Desktop"
        If Len(Dir$(base, vbDirectory)) = 0 Then
            base = Environ$("USERPROFILE") & "\Area de Trabalho"
        End If
        If Len(Dir$(base, vbDirectory)) = 0 Then
            base = Environ$("USERPROFILE")
        End If
    End If

    alvo = base & "\" & NOME_PASTA

    If Len(Dir$(alvo, vbDirectory)) = 0 Then
        On Error Resume Next
        MkDir alvo
        On Error GoTo 0
    End If

    If Len(Dir$(alvo, vbDirectory)) = 0 Then
        MsgBox "Nao foi possivel criar a pasta de saida:" & vbCrLf & vbCrLf & alvo & _
               vbCrLf & vbCrLf & "Verifique se voce tem permissao de escrita nesse local.", _
               vbCritical, "Erro"
        PrepararPasta = ""
    Else
        PrepararPasta = alvo
    End If

End Function


'==============================================================================
' GRAVACAO EM UTF-8
'==============================================================================
' comBOM = True  -> usado nos .csv, para o Excel abrir com acentuacao correta.
' comBOM = False -> usado nos .md e .txt, para o arquivo ficar limpo no Git.
Private Sub Gravar(ByVal nomeArquivo As String, ByVal conteudo As String, ByVal comBOM As Boolean)

    Dim caminho As String
    Dim stm As Object
    Dim bin As Object

    caminho = mPasta & "\" & nomeArquivo

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                      ' adTypeText
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText conteudo

    If comBOM Then
        stm.SaveToFile caminho, 2     ' adSaveCreateOverWrite
        stm.Close
    Else
        ' Reabre o mesmo conteudo como binario pulando os 3 bytes do BOM.
        stm.Position = 0
        stm.Type = 1                  ' adTypeBinary
        stm.Position = 3

        Set bin = CreateObject("ADODB.Stream")
        bin.Type = 1
        bin.Open
        stm.CopyTo bin
        stm.Close
        bin.SaveToFile caminho, 2
        bin.Close
    End If

    mArquivos(nomeArquivo) = FileLen(caminho)

End Sub


'------------------------------------------------------------------------------
' Acumulador de texto. Concatenar com & dentro de laco e O(n^2) e trava o Excel
' em arquivos grandes; aqui as linhas sao guardadas num array e unidas no fim.
'------------------------------------------------------------------------------
Private Sub SbInit()
    ReDim mBuf(0 To 1023)
    mBufN = 0
End Sub

Private Sub SbAdd(ByVal s As String)
    If mBufN > UBound(mBuf) Then
        ReDim Preserve mBuf(0 To (UBound(mBuf) + 1) * 2 - 1)
    End If
    mBuf(mBufN) = s
    mBufN = mBufN + 1
End Sub

Private Function SbTexto() As String
    If mBufN = 0 Then
        SbTexto = ""
    Else
        ReDim Preserve mBuf(0 To mBufN - 1)
        SbTexto = Join(mBuf, vbCrLf)
    End If
End Function


'==============================================================================
' 00_RESUMO.md
'==============================================================================
Private Sub GerarResumo()

    Dim ws As Worksheet
    Dim idx As Long
    Dim ultLin As Long, ultCol As Long
    Dim nFormulas As Double, nConstantes As Double
    Dim ur As String
    Dim inflado As String

    SbInit
    SbAdd "# Resumo da planilha"
    SbAdd ""
    SbAdd "Gerado por `ExportarPlanilha.bas` em " & Format$(Now, "yyyy-mm-dd hh:nn")
    SbAdd ""
    SbAdd "| Item | Valor |"
    SbAdd "|---|---|"
    SbAdd "| Arquivo | " & ThisWorkbook.Name & " |"
    SbAdd "| Pasta | " & ThisWorkbook.Path & " |"
    SbAdd "| Tamanho | " & TamanhoArquivo() & " |"
    SbAdd "| Versao do Excel | " & Application.Version & " |"
    SbAdd "| Modo de calculo | " & NomeCalculo(Application.Calculation) & " |"
    SbAdd "| Planilhas | " & ThisWorkbook.Worksheets.Count & " |"
    SbAdd "| Graficos (abas) | " & ThisWorkbook.Charts.Count & " |"
    SbAdd ""
    SbAdd "## Abas"
    SbAdd ""
    SbAdd "| # | Aba | Intervalo usado | Ultima celula real | Formulas | Constantes | Visivel | Protegida |"
    SbAdd "|---|---|---|---|---|---|---|---|"

    idx = 0
    For Each ws In ThisWorkbook.Worksheets
        idx = idx + 1
        Application.StatusBar = "Resumo: " & ws.Name

        ur = ""
        On Error Resume Next
        ur = ws.UsedRange.Address(False, False)
        On Error GoTo 0

        ultLin = UltimaLinhaReal(ws)
        ultCol = UltimaColunaReal(ws)

        nFormulas = ContarTipo(ws, xlCellTypeFormulas)
        nConstantes = ContarTipo(ws, xlCellTypeConstants)

        SbAdd "| " & idx & _
              " | " & ws.Name & _
              " | " & ur & _
              " | " & IIf(ultLin = 0, "(vazia)", Endereco(ultLin, IIf(ultCol = 0, 1, ultCol))) & _
              " | " & Fmt(nFormulas) & _
              " | " & Fmt(nConstantes) & _
              " | " & IIf(ws.Visible = xlSheetVisible, "sim", "NAO") & _
              " | " & IIf(ws.ProtectContents, "sim", "nao") & " |"
    Next ws

    ' Diagnostico do tamanho do arquivo.
    SbAdd ""
    SbAdd "## Diagnostico de tamanho"
    SbAdd ""
    SbAdd "Quando o ""intervalo usado"" vai muito alem da ""ultima celula real"", o Excel"
    SbAdd "esta guardando milhares de linhas vazias formatadas. Essa e a causa mais comum"
    SbAdd "de um arquivo inchado. Veja a secao correspondente em `docs/COMO-EXPORTAR.md`."
    SbAdd ""

    idx = 0
    For Each ws In ThisWorkbook.Worksheets
        idx = idx + 1
        ultLin = UltimaLinhaReal(ws)
        inflado = ""
        On Error Resume Next
        If ws.UsedRange.Rows.Count > ultLin + 100 And ultLin > 0 Then
            inflado = "- **" & ws.Name & "**: intervalo usado com " & _
                      Fmt(CDbl(ws.UsedRange.Rows.Count)) & " linhas, mas o ultimo dado real esta na linha " & _
                      Fmt(CDbl(ultLin)) & "."
        End If
        On Error GoTo 0
        If Len(inflado) > 0 Then SbAdd inflado
    Next ws

    Gravar "00_RESUMO.md", SbTexto(), False

End Sub


Private Function TamanhoArquivo() As String
    Dim n As Double
    On Error Resume Next
    n = FileLen(ThisWorkbook.FullName)
    On Error GoTo 0
    If n <= 0 Then
        TamanhoArquivo = "(indisponivel)"
    Else
        TamanhoArquivo = Format$(n / 1048576, "#,##0.0") & " MB"
    End If
End Function


Private Function NomeCalculo(ByVal v As Long) As String
    Select Case v
        Case xlCalculationAutomatic:            NomeCalculo = "automatico"
        Case xlCalculationManual:               NomeCalculo = "manual"
        Case xlCalculationSemiautomatic:        NomeCalculo = "automatico exceto tabelas"
        Case Else:                              NomeCalculo = "desconhecido (" & v & ")"
    End Select
End Function


Private Function ContarTipo(ws As Worksheet, ByVal tipo As Long) As Double
    Dim rng As Range
    Set rng = Nothing
    On Error Resume Next
    Set rng = ws.Cells.SpecialCells(tipo)
    On Error GoTo 0
    If rng Is Nothing Then
        ContarTipo = 0
    Else
        ContarTipo = ContarCelulas(rng)
    End If
End Function


'==============================================================================
' 10_FORMULAS_<aba>.txt   -- o arquivo mais importante da exportacao
'==============================================================================
' Estrategia: agrupar as formulas pela notacao R1C1.
' Uma formula copiada para baixo ou para o lado tem R1C1 IDENTICA, entao dezenas
' de milhares de celulas colapsam em poucas dezenas de padroes.
'==============================================================================
Private Sub GerarFormulas()

    Dim ws As Worksheet
    Dim idx As Long

    idx = 0
    For Each ws In ThisWorkbook.Worksheets
        idx = idx + 1
        Application.StatusBar = "Formulas: " & ws.Name
        ExportarFormulasDaAba ws, idx
    Next ws

End Sub


Private Sub ExportarFormulasDaAba(ws As Worksheet, ByVal idx As Long)

    Dim rng As Range
    Dim area As Range
    Dim dic As Object
    Dim chaves() As String
    Dim contagens() As Double
    Dim i As Long
    Dim k As Variant
    Dim reg As Variant
    Dim total As Double
    Dim nomeArq As String

    Set rng = Nothing
    On Error Resume Next
    Set rng = ws.Cells.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    Set dic = CreateObject("Scripting.Dictionary")

    For Each area In rng.Areas
        ColetarArea area, dic
    Next area

    If dic.Count = 0 Then Exit Sub

    ' Passa o dicionario para arrays paralelos e ordena por ocorrencias.
    ReDim chaves(0 To dic.Count - 1)
    ReDim contagens(0 To dic.Count - 1)

    i = 0
    total = 0
    For Each k In dic.Keys
        reg = dic(k)
        chaves(i) = CStr(k)
        contagens(i) = reg(0)
        total = total + reg(0)
        i = i + 1
    Next k

    If dic.Count <= MAX_PADROES_ORDENAR Then OrdenarDesc chaves, contagens

    SbInit
    SbAdd "================================================================"
    SbAdd "FORMULAS DA ABA: " & ws.Name
    SbAdd "================================================================"
    SbAdd ""
    SbAdd "Total de celulas com formula : " & Fmt(total)
    SbAdd "Padroes distintos (R1C1)     : " & Fmt(CDbl(dic.Count))
    SbAdd ""
    SbAdd "Como ler este arquivo:"
    SbAdd "  R1C1  = o padrao da formula. Igual para toda celula copiada."
    SbAdd "          R = linha atual, C = coluna atual. R[-1] = uma linha acima."
    SbAdd "  A1-PT = a formula como aparece no Excel em portugues."
    SbAdd "  A1-EN = a mesma formula com os nomes de funcao em ingles."
    SbAdd "          E esta que serve para o Google Sheets e para programacao."
    SbAdd ""
    If dic.Count > MAX_PADROES_ORDENAR Then
        SbAdd "AVISO: muitos padroes distintos; a lista nao foi ordenada por frequencia."
        SbAdd ""
    End If

    For i = LBound(chaves) To UBound(chaves)
        reg = dic(chaves(i))
        SbAdd "----------------------------------------------------------------"
        SbAdd "Ocorrencias : " & Fmt(reg(0))
        SbAdd "Primeira    : " & Endereco(CLng(reg(1)), CLng(reg(2)))
        SbAdd "Ultima      : " & Endereco(CLng(reg(3)), CLng(reg(4)))

        Dim passo As String
        passo = DescreverPasso(reg)
        If Len(passo) > 0 Then SbAdd "Distribuicao: " & passo

        SbAdd "R1C1        : " & chaves(i)
        SbAdd "A1-PT       : " & FormulaDaCelula(ws, CLng(reg(1)), CLng(reg(2)), True)
        SbAdd "A1-EN       : " & FormulaDaCelula(ws, CLng(reg(1)), CLng(reg(2)), False)
        SbAdd ""
    Next i

    nomeArq = "10_FORMULAS_" & Format$(idx, "00") & "_" & AsciiNome(ws.Name) & ".txt"
    Gravar nomeArq, SbTexto(), False

End Sub


' Descreve o espacamento entre as ocorrencias. E isto que revela estruturas do
' tipo "um bloco a cada 6 linhas", muito comum em relatorios montados a mao.
Private Function DescreverPasso(reg As Variant) As String

    Dim n As Double
    Dim linIni As Long, colIni As Long, linFim As Long, colFim As Long
    Dim passo As Double

    n = reg(0)
    linIni = CLng(reg(1)): colIni = CLng(reg(2))
    linFim = CLng(reg(3)): colFim = CLng(reg(4))

    If n < 3 Then Exit Function
    If colIni <> colFim Then Exit Function
    If linFim <= linIni Then Exit Function

    passo = (linFim - linIni) / (n - 1)

    If passo = Int(passo) And passo >= 1 Then
        If passo = 1 Then
            DescreverPasso = "linhas consecutivas de " & linIni & " a " & linFim
        Else
            DescreverPasso = "uma a cada " & CLng(passo) & " linhas, de " & linIni & " a " & linFim
        End If
    End If

End Function


Private Function FormulaDaCelula(ws As Worksheet, ByVal lin As Long, ByVal col As Long, _
                                 ByVal localizada As Boolean) As String
    Dim s As String
    On Error Resume Next
    If localizada Then
        s = ws.Cells(lin, col).FormulaLocal
    Else
        s = ws.Cells(lin, col).Formula
    End If
    On Error GoTo 0
    If Len(s) = 0 Then s = "(indisponivel)"
    FormulaDaCelula = s
End Function


' Percorre uma area em blocos, para nao carregar uma matriz gigante na memoria.
Private Sub ColetarArea(area As Range, dic As Object)

    Dim nCols As Long, nRows As Long
    Dim passo As Long, r As Long
    Dim altura As Long

    nCols = area.Columns.Count
    nRows = area.Rows.Count

    passo = MAX_CELULAS_BLOCO \ nCols
    If passo < 1 Then passo = 1

    For r = 1 To nRows Step passo
        altura = passo
        If r + altura - 1 > nRows Then altura = nRows - r + 1
        ColetarBloco area.Cells(r, 1).Resize(altura, nCols), dic
    Next r

End Sub


Private Sub ColetarBloco(bloco As Range, dic As Object)

    Dim v As Variant
    Dim i As Long, j As Long
    Dim f As String
    Dim linBase As Long, colBase As Long

    linBase = bloco.Row
    colBase = bloco.Column

    ' Uma unica celula devolve String, nao matriz. Caso classico de erro.
    If bloco.Cells.Count = 1 Then
        f = CStr(bloco.FormulaR1C1)
        If Left$(f, 1) = "=" Then RegistrarFormula dic, f, linBase, colBase
        Exit Sub
    End If

    v = bloco.FormulaR1C1

    For i = 1 To UBound(v, 1)
        For j = 1 To UBound(v, 2)
            If Not IsError(v(i, j)) Then
                f = CStr(v(i, j))
                If Left$(f, 1) = "=" Then
                    RegistrarFormula dic, f, linBase + i - 1, colBase + j - 1
                End If
            End If
        Next j
    Next i

End Sub


Private Sub RegistrarFormula(dic As Object, ByVal f As String, ByVal lin As Long, ByVal col As Long)

    Dim reg As Variant

    If dic.Exists(f) Then
        reg = dic(f)
        reg(0) = reg(0) + 1
        reg(3) = lin
        reg(4) = col
        dic(f) = reg
    Else
        dic.Add f, Array(CDbl(1), lin, col, lin, col)
    End If

End Sub


Private Sub OrdenarDesc(chaves() As String, contagens() As Double)

    Dim i As Long, j As Long
    Dim ts As String, td As Double

    For i = LBound(chaves) To UBound(chaves) - 1
        For j = i + 1 To UBound(chaves)
            If contagens(j) > contagens(i) Then
                ts = chaves(i): chaves(i) = chaves(j): chaves(j) = ts
                td = contagens(i): contagens(i) = contagens(j): contagens(j) = td
            End If
        Next j
    Next i

End Sub


'==============================================================================
' 20_NOMES.csv
'==============================================================================
Private Sub GerarNomes()

    Dim nm As Name
    Dim refere As String
    Dim escopo As String
    Dim visivel As String

    SbInit
    SbAdd CabecalhoCsv(Array("nome", "refere_a", "escopo", "visivel"))

    For Each nm In ThisWorkbook.Names
        refere = "(erro ao ler)"
        escopo = "pasta de trabalho"
        visivel = "sim"

        On Error Resume Next
        refere = nm.RefersTo
        If Not nm.Visible Then visivel = "nao"
        If TypeName(nm.Parent) = "Worksheet" Then escopo = nm.Parent.Name
        On Error GoTo 0

        SbAdd LinhaCsv(Array(nm.Name, refere, escopo, visivel))
    Next nm

    Gravar "20_NOMES.csv", SbTexto(), True

End Sub


'==============================================================================
' 30_VALIDACOES.csv  -- revela as listas suspensas e as regras de entrada
'==============================================================================
Private Sub GerarValidacoes()

    Dim ws As Worksheet
    Dim rng As Range
    Dim area As Range
    Dim tipo As String, f1 As String, f2 As String, msg As String

    SbInit
    SbAdd CabecalhoCsv(Array("aba", "intervalo", "tipo", "formula1", "formula2", "mensagem_entrada"))

    For Each ws In ThisWorkbook.Worksheets
        Set rng = Nothing
        On Error Resume Next
        Set rng = ws.Cells.SpecialCells(xlCellTypeAllValidation)
        On Error GoTo 0
        If Not rng Is Nothing Then
            For Each area In rng.Areas
                tipo = "": f1 = "": f2 = "": msg = ""
                On Error Resume Next
                tipo = NomeTipoValidacao(area.Validation.Type)
                f1 = area.Validation.Formula1
                f2 = area.Validation.Formula2
                msg = area.Validation.InputMessage
                On Error GoTo 0
                SbAdd LinhaCsv(Array(ws.Name, area.Address(False, False), tipo, f1, f2, msg))
            Next area
        End If
    Next ws

    Gravar "30_VALIDACOES.csv", SbTexto(), True

End Sub


Private Function NomeTipoValidacao(ByVal v As Long) As String
    Select Case v
        Case 0: NomeTipoValidacao = "qualquer valor"
        Case 1: NomeTipoValidacao = "numero inteiro"
        Case 2: NomeTipoValidacao = "decimal"
        Case 3: NomeTipoValidacao = "lista"
        Case 4: NomeTipoValidacao = "data"
        Case 5: NomeTipoValidacao = "hora"
        Case 6: NomeTipoValidacao = "comprimento do texto"
        Case 7: NomeTipoValidacao = "personalizada"
        Case Else: NomeTipoValidacao = "tipo " & v
    End Select
End Function


'==============================================================================
' 40_FORMATACAO_CONDICIONAL.csv
'==============================================================================
' As cores da planilha costumam carregar regra de negocio (atrasado, no prazo,
' maquina parada). Sem isto a regra se perde na migracao.
'==============================================================================
Private Sub GerarFormatacaoCondicional()

    Dim ws As Worksheet
    Dim rng As Range
    Dim area As Range
    Dim fc As Object
    Dim i As Long
    Dim vistos As Object
    Dim chave As String
    Dim aplicaA As String, tipo As String, f1 As String, f2 As String, cor As String

    Set vistos = CreateObject("Scripting.Dictionary")

    SbInit
    SbAdd CabecalhoCsv(Array("aba", "aplica_a", "tipo", "formula1", "formula2", "cor_fundo"))

    For Each ws In ThisWorkbook.Worksheets
        Set rng = Nothing
        On Error Resume Next
        Set rng = ws.Cells.SpecialCells(xlCellTypeAllFormatConditions)
        On Error GoTo 0
        If Not rng Is Nothing Then
            For Each area In rng.Areas
                For i = 1 To area.Cells(1, 1).FormatConditions.Count
                    Set fc = Nothing
                    On Error Resume Next
                    Set fc = area.Cells(1, 1).FormatConditions(i)
                    On Error GoTo 0
                    If Not fc Is Nothing Then

                        aplicaA = "": tipo = "": f1 = "": f2 = "": cor = ""
                        On Error Resume Next
                        aplicaA = fc.AppliesTo.Address(False, False)
                        tipo = NomeTipoFC(fc.Type)
                        f1 = fc.Formula1
                        f2 = fc.Formula2
                        cor = "RGB " & fc.Interior.Color
                        On Error GoTo 0

                        chave = ws.Name & "|" & aplicaA & "|" & tipo & "|" & f1 & "|" & f2
                        If Not vistos.Exists(chave) Then
                            vistos.Add chave, 1
                            SbAdd LinhaCsv(Array(ws.Name, aplicaA, tipo, f1, f2, cor))
                        End If

                    End If
                Next i
            Next area
        End If
    Next ws

    Gravar "40_FORMATACAO_CONDICIONAL.csv", SbTexto(), True

End Sub


Private Function NomeTipoFC(ByVal v As Long) As String
    Select Case v
        Case 1:  NomeTipoFC = "valor da celula"
        Case 2:  NomeTipoFC = "formula"
        Case 3:  NomeTipoFC = "escala de cor"
        Case 4:  NomeTipoFC = "barra de dados"
        Case 6:  NomeTipoFC = "conjunto de icones"
        Case 8:  NomeTipoFC = "texto especifico"
        Case 9:  NomeTipoFC = "acima ou abaixo da media"
        Case 10: NomeTipoFC = "primeiros ou ultimos"
        Case 12: NomeTipoFC = "valores unicos"
        Case 16: NomeTipoFC = "data ocorrendo"
        Case Else: NomeTipoFC = "tipo " & v
    End Select
End Function


'==============================================================================
' 50_AMOSTRA_<aba>.csv
'==============================================================================
Private Sub GerarAmostras()

    Dim ws As Worksheet
    Dim idx As Long

    idx = 0
    For Each ws In ThisWorkbook.Worksheets
        idx = idx + 1
        Application.StatusBar = "Amostra: " & ws.Name
        ExportarAmostraDaAba ws, idx
    Next ws

End Sub


Private Sub ExportarAmostraDaAba(ws As Worksheet, ByVal idx As Long)

    Dim ultLin As Long, ultCol As Long
    Dim nLin As Long, nCol As Long
    Dim v As Variant
    Dim i As Long, j As Long
    Dim campos() As String

    ultLin = UltimaLinhaReal(ws)
    ultCol = UltimaColunaReal(ws)
    If ultLin = 0 Or ultCol = 0 Then Exit Sub

    nLin = LINHAS_AMOSTRA
    If nLin > ultLin Then nLin = ultLin

    If COLUNAS_AMOSTRA > 0 Then
        nCol = COLUNAS_AMOSTRA
    Else
        nCol = ultCol
    End If
    If nCol > ultCol Then nCol = ultCol

    v = ws.Range(ws.Cells(1, 1), ws.Cells(nLin, nCol)).Value

    SbInit

    ' Uma unica celula nao devolve matriz.
    ' ParaTexto ja devolve o campo escapado; nao passar por LinhaCsv de novo.
    If nLin = 1 And nCol = 1 Then
        SbAdd ParaTexto(v)
    Else
        ReDim campos(1 To nCol)
        For i = 1 To nLin
            For j = 1 To nCol
                campos(j) = ParaTexto(v(i, j))
            Next j
            SbAdd Join(campos, SEPARADOR)
        Next i
    End If

    Gravar "50_AMOSTRA_" & Format$(idx, "00") & "_" & AsciiNome(ws.Name) & ".csv", SbTexto(), True

End Sub


'==============================================================================
' 60_LIGACOES.txt
'==============================================================================
Private Sub GerarLigacoes()

    Dim links As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim wsRef As Worksheet
    Dim rng As Range
    Dim encontrou As Boolean

    SbInit
    SbAdd "================================================================"
    SbAdd "LIGACOES"
    SbAdd "================================================================"
    SbAdd ""
    SbAdd "## Vinculos com outros arquivos"
    SbAdd ""

    links = Empty
    On Error Resume Next
    links = ThisWorkbook.LinkSources(1)     ' xlExcelLinks
    On Error GoTo 0

    If IsEmpty(links) Then
        SbAdd "(nenhum vinculo externo)"
    Else
        For i = LBound(links) To UBound(links)
            SbAdd "- " & links(i)
        Next i
    End If

    SbAdd ""
    SbAdd "## Quais abas referenciam quais"
    SbAdd ""
    SbAdd "Deduzido procurando o nome de cada aba dentro das formulas das demais."
    SbAdd ""

    For Each ws In ThisWorkbook.Worksheets

        ' Calculado uma vez por aba, e nao a cada par de abas.
        Set rng = Nothing
        On Error Resume Next
        Set rng = ws.Cells.SpecialCells(xlCellTypeFormulas)
        On Error GoTo 0

        If Not rng Is Nothing Then
            For Each wsRef In ThisWorkbook.Worksheets
                If ws.Name <> wsRef.Name Then

                    encontrou = False
                    On Error Resume Next
                    encontrou = Not rng.Find(What:=wsRef.Name, LookIn:=xlFormulas, _
                                             LookAt:=xlPart, MatchCase:=False) Is Nothing
                    On Error GoTo 0

                    If encontrou Then
                        SbAdd "- `" & ws.Name & "` usa dados de `" & wsRef.Name & "`"
                    End If

                End If
            Next wsRef
        End If

    Next ws

    Gravar "60_LIGACOES.txt", SbTexto(), False

End Sub


'==============================================================================
' 70_OBJETOS.txt  -- botoes, macros atribuidas, graficos, tabelas dinamicas
'==============================================================================
Private Sub GerarObjetos()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim pt As PivotTable
    Dim lo As ListObject
    Dim co As ChartObject
    Dim macro As String
    Dim texto As String

    SbInit
    SbAdd "================================================================"
    SbAdd "OBJETOS DA PLANILHA"
    SbAdd "================================================================"
    SbAdd ""
    SbAdd "## Botoes e formas com macro atribuida"
    SbAdd ""
    SbAdd "Isto mapeia a interface: qual botao dispara qual macro."
    SbAdd ""

    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            macro = ""
            texto = ""
            On Error Resume Next
            macro = shp.OnAction
            texto = shp.TextFrame.Characters.Text
            On Error GoTo 0

            If Len(macro) > 0 Then
                SbAdd "- [" & ws.Name & "] " & shp.Name & _
                      IIf(Len(texto) > 0, " (""" & texto & """)", "") & _
                      "  ->  macro: " & macro
            End If
        Next shp
    Next ws

    SbAdd ""
    SbAdd "## Tabelas dinamicas"
    SbAdd ""
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            SbAdd "- [" & ws.Name & "] " & pt.Name
            On Error Resume Next
            SbAdd "    origem: " & CStr(pt.SourceData)
            SbAdd "    area  : " & pt.TableRange1.Address(False, False)
            On Error GoTo 0
        Next pt
    Next ws

    SbAdd ""
    SbAdd "## Tabelas nomeadas"
    SbAdd ""
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            SbAdd "- [" & ws.Name & "] " & lo.Name & "  " & lo.Range.Address(False, False)
        Next lo
    Next ws

    SbAdd ""
    SbAdd "## Graficos"
    SbAdd ""
    For Each ws In ThisWorkbook.Worksheets
        For Each co In ws.ChartObjects
            SbAdd "- [" & ws.Name & "] " & co.Name
        Next co
    Next ws

    Gravar "70_OBJETOS.txt", SbTexto(), False

End Sub


'==============================================================================
' 80_ESTRUTURA.txt  -- pistas do layout: mesclagens, ocultos, larguras
'==============================================================================
Private Sub GerarEstrutura()

    Dim ws As Worksheet
    Dim ultLin As Long, ultCol As Long
    Dim i As Long
    Dim vistos As Object
    Dim cel As Range
    Dim ender As String
    Dim ocultas As String

    SbInit
    SbAdd "================================================================"
    SbAdd "ESTRUTURA VISUAL"
    SbAdd "================================================================"
    SbAdd ""

    For Each ws In ThisWorkbook.Worksheets

        Application.StatusBar = "Estrutura: " & ws.Name

        ultLin = UltimaLinhaReal(ws)
        ultCol = UltimaColunaReal(ws)
        ' Ambos precisam ser validos: ws.Cells(linha, 0) geraria erro 1004.
        If ultLin = 0 Or ultCol = 0 Then GoTo ProximaAba

        SbAdd "----------------------------------------------------------------"
        SbAdd "ABA: " & ws.Name
        SbAdd "----------------------------------------------------------------"

        ' --- celulas mescladas (so no topo, onde fica o cabecalho do relatorio)
        Set vistos = CreateObject("Scripting.Dictionary")
        SbAdd ""
        SbAdd "Celulas mescladas (primeiras " & MAX_LINHAS_MESCLA & " linhas):"

        Dim limite As Long
        Dim limiteCol As Long
        limite = ultLin
        If limite > MAX_LINHAS_MESCLA Then limite = MAX_LINHAS_MESCLA

        ' Varrer celula a celula custa uma chamada COM por celula; limitar a
        ' largura evita que uma aba muito larga deixe a exportacao lenta.
        limiteCol = ultCol
        If limiteCol > 200 Then limiteCol = 200

        ' Uma unica leitura resolve o caso comum: MergeCells devolve False quando
        ' nenhuma celula do bloco esta mesclada, e Null quando ha mistura. So
        ' vale varrer celula a celula (uma chamada COM por celula) no caso Null.
        Dim mescladas As Variant
        mescladas = ws.Range(ws.Cells(1, 1), ws.Cells(limite, limiteCol)).MergeCells

        If VarType(mescladas) = vbBoolean And mescladas = False Then
            SbAdd "  (nenhuma)"
        Else
            For Each cel In ws.Range(ws.Cells(1, 1), ws.Cells(limite, limiteCol))
                If cel.MergeCells Then
                    ender = cel.MergeArea.Address(False, False)
                    If Not vistos.Exists(ender) Then
                        vistos.Add ender, 1
                        SbAdd "  " & ender & "  = " & Left$(ValorTexto(cel.MergeArea.Cells(1, 1).Value), 60)
                    End If
                End If
            Next cel
            If vistos.Count = 0 Then SbAdd "  (nenhuma)"
        End If

        ' --- colunas ocultas
        SbAdd ""
        ocultas = ""
        For i = 1 To ultCol
            If ws.Columns(i).Hidden Then ocultas = ocultas & LetraColuna(i) & " "
        Next i
        SbAdd "Colunas ocultas: " & IIf(Len(ocultas) = 0, "(nenhuma)", ocultas)

        ' --- linhas ocultas
        ocultas = ""
        For i = 1 To limite
            If ws.Rows(i).Hidden Then ocultas = ocultas & i & " "
        Next i
        SbAdd "Linhas ocultas (ate a linha " & limite & "): " & _
              IIf(Len(ocultas) = 0, "(nenhuma)", ocultas)

        ' --- larguras
        SbAdd ""
        SbAdd "Largura das colunas:"
        Dim larg As String
        larg = ""
        For i = 1 To ultCol
            larg = larg & LetraColuna(i) & "=" & Format$(ws.Columns(i).ColumnWidth, "0.0") & "  "
            If i Mod 10 = 0 Then
                SbAdd "  " & larg
                larg = ""
            End If
        Next i
        If Len(larg) > 0 Then SbAdd "  " & larg

        SbAdd ""

ProximaAba:
    Next ws

    ' Painel congelado so pode ser lido da janela ativa, sem ativar cada aba
    ' (ativar abas mudaria o estado da planilha, o que esta macro evita).
    SbAdd "----------------------------------------------------------------"
    SbAdd "Painel congelado (apenas da aba ativa no momento da exportacao)"
    SbAdd "----------------------------------------------------------------"
    On Error Resume Next
    SbAdd "Aba ativa      : " & ActiveSheet.Name
    SbAdd "Congelado      : " & IIf(ActiveWindow.FreezePanes, "sim", "nao")
    SbAdd "Linhas fixas   : " & ActiveWindow.SplitRow
    SbAdd "Colunas fixas  : " & ActiveWindow.SplitColumn
    On Error GoTo 0

    Gravar "80_ESTRUTURA.txt", SbTexto(), False

End Sub


'==============================================================================
' 99_ARQUIVOS.txt
'==============================================================================
Private Sub GerarListaArquivos()

    Dim k As Variant

    SbInit
    SbAdd "Arquivos gerados por ExportarPlanilha.bas"
    SbAdd "Data: " & Format$(Now, "yyyy-mm-dd hh:nn")
    SbAdd ""

    For Each k In mArquivos.Keys
        SbAdd Format$(mArquivos(k) / 1024, "#,##0") & " KB  " & CStr(k)
    Next k

    Gravar "99_ARQUIVOS.txt", SbTexto(), False

End Sub


Private Function Relatorio(ByVal t0 As Single) As String

    Dim k As Variant
    Dim s As String
    Dim n As Long
    Dim totalBytes As Double

    n = 0
    totalBytes = 0
    For Each k In mArquivos.Keys
        totalBytes = totalBytes + mArquivos(k)
    Next k

    s = "Exportacao concluida em " & Format$(Timer - t0, "0.0") & " segundos." & vbCrLf & vbCrLf
    s = s & "Pasta:" & vbCrLf & mPasta & vbCrLf & vbCrLf
    s = s & mArquivos.Count & " arquivos, " & Format$(totalBytes / 1024, "#,##0") & " KB no total." & vbCrLf & vbCrLf

    For Each k In mArquivos.Keys
        n = n + 1
        If n > 20 Then
            s = s & "  ... e mais " & (mArquivos.Count - 20) & " arquivos" & vbCrLf
            Exit For
        End If
        s = s & "  " & CStr(k) & vbCrLf
    Next k

    s = s & vbCrLf & "Proximo passo: copie esta pasta para docs/planilha/ no repositorio."

    Relatorio = s

End Function


'==============================================================================
' UTILITARIOS
'==============================================================================

Private Function UltimaLinhaReal(ws As Worksheet) As Long
    Dim c As Range
    Set c = Nothing
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If c Is Nothing Then UltimaLinhaReal = 0 Else UltimaLinhaReal = c.Row
End Function


Private Function UltimaColunaReal(ws As Worksheet) As Long
    Dim c As Range
    Set c = Nothing
    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If c Is Nothing Then UltimaColunaReal = 0 Else UltimaColunaReal = c.Column
End Function


' Cells.Count estoura o limite de Long numa planilha inteira; CountLarge nao.
Private Function ContarCelulas(rng As Range) As Double
    Dim n As Double
    n = -1
    On Error Resume Next
    n = rng.CountLarge
    On Error GoTo 0
    If n < 0 Then
        On Error Resume Next
        n = rng.Count
        On Error GoTo 0
    End If
    If n < 0 Then n = 0
    ContarCelulas = n
End Function


Private Function Fmt(ByVal n As Double) As String
    Fmt = Format$(n, "#,##0")
End Function


Private Function Endereco(ByVal lin As Long, ByVal col As Long) As String
    Endereco = LetraColuna(col) & CStr(lin)
End Function


Private Function LetraColuna(ByVal col As Long) As String
    Dim s As String
    Dim n As Long
    Dim r As Long
    n = col
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr$(65 + r) & s
        n = (n - 1) \ 26
    Loop
    If Len(s) = 0 Then s = "?"
    LetraColuna = s
End Function


' Converte o nome da aba num nome de arquivo seguro em qualquer sistema.
' O nome real e acentuado continua gravado DENTRO do arquivo.
Private Function AsciiNome(ByVal s As String) As String

    Dim i As Long
    Dim c As Long
    Dim ch As String
    Dim r As String

    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        Select Case c
            Case 192, 193, 194, 195, 196, 197: ch = "A"
            Case 224, 225, 226, 227, 228, 229: ch = "a"
            Case 199:                          ch = "C"
            Case 231:                          ch = "c"
            Case 200, 201, 202, 203:           ch = "E"
            Case 232, 233, 234, 235:           ch = "e"
            Case 204, 205, 206, 207:           ch = "I"
            Case 236, 237, 238, 239:           ch = "i"
            Case 209:                          ch = "N"
            Case 241:                          ch = "n"
            Case 210, 211, 212, 213, 214:      ch = "O"
            Case 242, 243, 244, 245, 246:      ch = "o"
            Case 217, 218, 219, 220:           ch = "U"
            Case 249, 250, 251, 252:           ch = "u"
            Case 48 To 57, 65 To 90, 97 To 122: ch = Chr$(c)
            Case Else:                         ch = "_"
        End Select
        r = r & ch
    Next i

    Do While InStr(r, "__") > 0
        r = Replace(r, "__", "_")
    Loop

    Do While Left$(r, 1) = "_"
        r = Mid$(r, 2)
    Loop
    Do While Right$(r, 1) = "_" And Len(r) > 0
        r = Left$(r, Len(r) - 1)
    Loop

    If Len(r) = 0 Then r = "aba"
    If Len(r) > 50 Then r = Left$(r, 50)

    AsciiNome = r

End Function


Private Function CabecalhoCsv(campos As Variant) As String
    CabecalhoCsv = LinhaCsv(campos)
End Function


Private Function LinhaCsv(campos As Variant) As String
    Dim i As Long
    Dim saida() As String
    ReDim saida(LBound(campos) To UBound(campos))
    For i = LBound(campos) To UBound(campos)
        saida(i) = CampoCsv(CStr(campos(i)))
    Next i
    LinhaCsv = Join(saida, SEPARADOR)
End Function


Private Function CampoCsv(ByVal s As String) As String
    If InStr(s, SEPARADOR) > 0 Or InStr(s, """") > 0 _
       Or InStr(s, vbCr) > 0 Or InStr(s, vbLf) > 0 Then
        CampoCsv = """" & Replace(s, """", """""") & """"
    Else
        CampoCsv = s
    End If
End Function


' Converte um valor de celula em texto legivel, sem escape de CSV.
' Use esta versao em arquivos .txt e .md.
Private Function ValorTexto(ByVal v As Variant) As String

    If IsError(v) Then
        ValorTexto = "#ERRO"
    ElseIf IsEmpty(v) Then
        ValorTexto = ""
    ElseIf IsNull(v) Then
        ValorTexto = ""
    ElseIf VarType(v) = vbDate Then
        If Int(CDbl(v)) = CDbl(v) Then
            ValorTexto = Format$(v, "yyyy-mm-dd")
        Else
            ValorTexto = Format$(v, "yyyy-mm-dd hh:nn")
        End If
    ElseIf VarType(v) = vbBoolean Then
        ValorTexto = IIf(v, "VERDADEIRO", "FALSO")
    ElseIf IsNumeric(v) Then
        ' Decimal sempre com ponto, para o valor nao colidir com o separador ";".
        ValorTexto = Replace(CStr(v), ",", ".")
    Else
        ValorTexto = CStr(v)
    End If

End Function


' Mesma conversao, ja escapada para uso dentro de um arquivo .csv.
Private Function ParaTexto(ByVal v As Variant) As String
    ParaTexto = CampoCsv(ValorTexto(v))
End Function

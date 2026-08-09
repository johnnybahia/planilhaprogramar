Attribute VB_Name = "Módulo17"
 Sub InserirFormulas()
    Dim wsRelatorio As Worksheet
    Dim wsPlanejamento As Worksheet
    Dim linhaRelatorio As Long
    Dim linhaPlanejamento As Long
    Dim ultimaLinha As Long

    On Error GoTo ErrorHandler ' Adicionar tratamento de erro

    ' Definir as planilhas
    Set wsRelatorio = ThisWorkbook.Worksheets("RELATÓRIO TRANÇADEIRAS")
    Set wsPlanejamento = ThisWorkbook.Worksheets("Planejamento Trançadeira")

    ' Configurar as linhas iniciais e finais
    linhaRelatorio = 22
    linhaPlanejamento = 1
    ultimaLinha = 790

    ' Loop para inserir as fórmulas na coluna H
    Do While linhaRelatorio <= ultimaLinha
        wsRelatorio.Cells(linhaRelatorio, 8).Formula = "=MENOR('Planejamento Trançadeira'!C" & linhaPlanejamento & ":FJ" & linhaPlanejamento & ",CONT.SE('Planejamento Trançadeira'!C" & linhaPlanejamento & ":FJ" & linhaPlanejamento & ",0)+1)"
        
        ' Avançar para a próxima célula e linha
        linhaRelatorio = linhaRelatorio + 6
        linhaPlanejamento = linhaPlanejamento + 1
    Loop

    MsgBox "Fórmulas inseridas com sucesso!"
    Exit Sub

ErrorHandler:
    MsgBox "Ocorreu um erro: " & Err.Description, vbCritical
End Sub


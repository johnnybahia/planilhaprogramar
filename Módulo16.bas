Attribute VB_Name = "Módulo16"
Sub AtualizarValores()
    Dim i As Long
    Dim wsRelatorio As Worksheet
    Dim wsProgramacao As Worksheet
    Dim linha As Long

    ' Defina as planilhas
    Set wsRelatorio = ThisWorkbook.Sheets("RELATÓRIO TRANÇADEIRAS")
    Set wsProgramacao = ThisWorkbook.Sheets("Programação trançadeiras")

    ' Comece a partir da célula C19 e insira a fórmula em cada 6 células
    For i = 19 To 787 Step 6
        linha = (i - 19) / 6 + 3 ' Calcula a linha da planilha Programação trançadeiras (J3, J4, J5...)
        wsRelatorio.Cells(i, 3).Formula = "='Programação trançadeiras'!J" & linha
    Next i
End Sub


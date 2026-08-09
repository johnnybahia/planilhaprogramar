Attribute VB_Name = "Módulo18"
Sub CopiarColarFormulasAjustadas()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("RELATÓRIO TRANÇADEIRAS")
    
    Dim linhaDestino As Long
    linhaDestino = 23  ' Inicia na linha 23 (para colar em H23 e H24)
    
    Do While linhaDestino <= 791
        ' Copia o intervalo H20:H21
        ws.Range("H20:H21").Copy
        ' Cola como fórmulas na posição de destino, permitindo que o Excel ajuste as referências relativas
        ws.Range("H" & linhaDestino).PasteSpecial Paste:=xlPasteFormulas
        Application.CutCopyMode = False
        
        ' Incrementa 3 linhas para pular uma linha entre os grupos
        linhaDestino = linhaDestino + 3
    Loop
End Sub



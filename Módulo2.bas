Attribute VB_Name = "Módulo2"
Sub JuntarDados()
    Dim ultimaLinha As Long
    ultimaLinha = Cells(Rows.Count, "H").End(xlUp).Row ' determina a última linha da coluna H
    
    For i = 4 To ultimaLinha ' loop pelas linhas da coluna H a partir da linha 4 até a última linha
        Cells(i, 1).Value = Cells(i, 1).Value & Cells(i, 8).Value ' concatena o valor da célula A com o valor da célula H na mesma linha e armazena na coluna A
    Next i
    
    Range("A4").Select ' seleciona a célula A4
End Sub



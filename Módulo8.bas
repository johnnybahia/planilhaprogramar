Attribute VB_Name = "Módulo8"
Sub Juntar()
    Dim ultimaLinha As Long
    ultimaLinha = Cells(Rows.Count, "H").End(xlUp).Row ' determina a última linha da coluna H
    
    For i = 4 To ultimaLinha ' loop pelas linhas da coluna H a partir da linha 4 até a última linha
        If InStr(1, LCase(Cells(i, "H").Value), "mm") > 0 Then ' verifica se a célula na coluna H contém "mm" em qualquer posição
            Cells(i, "g").Value = Cells(i, "G").Value & " " & Cells(i, "H").Value ' concatena o valor da célula G com o valor da célula H na mesma linha e armazena na coluna A
        End If
    Next i
    
    Range("g4").Select ' seleciona a célula A4
End Sub


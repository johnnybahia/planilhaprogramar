Attribute VB_Name = "Módulo5"
Sub UltimaLinhaColunaC()
    Dim ultimaLinha As Long
    ultimaLinha = Cells(Rows.Count, "C").End(xlUp).Row
    Range("C" & ultimaLinha).Select
End Sub


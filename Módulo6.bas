Attribute VB_Name = "Módulo6"
Sub UltimaLinhaColunab()
    Dim ultimaLinha As Long
    ultimaLinha = Cells(Rows.Count, "b").End(xlUp).Row
    Range("b" & ultimaLinha).Select
End Sub


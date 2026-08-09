Attribute VB_Name = "Módulo11"
Sub Reexibir_Todas_Colunas()

    ' Percorre todas as colunas da planilha ativa
    For Each coluna In ActiveSheet.UsedRange.Columns
        
        ' Se a coluna estiver oculta, reexiba-a
        If coluna.Hidden Then
            coluna.Hidden = False
        End If
    Next coluna

End Sub




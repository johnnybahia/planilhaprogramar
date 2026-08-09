Attribute VB_Name = "Módulo4"
Sub CopiarNaoRepetidos1()

    Dim texto As String
    Dim celula As Range
    Dim celulaDestino As Range
    
    Set celulaDestino = Range("c2") 'define a primeira célula de destino como J4
    
    'percorre todas as células com texto da coluna F a partir de F4
    For Each celula In Range("c2:c" & Range("c" & Rows.Count).End(xlUp).Row)
    
        'verifica se o texto da célula já foi copiado para a coluna J
        If WorksheetFunction.CountIf(Range("a2:a" & celulaDestino.Row - 1), celula.Value) = 0 Then
        
            'copia o texto para a próxima célula de destino e atualiza a variável celulaDestino
            celulaDestino.Value = celula.Value
            Set celulaDestino = celulaDestino.Offset(1, 0)
        End If
    Next celula
    
End Sub


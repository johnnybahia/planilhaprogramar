Attribute VB_Name = "Módulo3"
Sub CopiarNaoRepetidos()

    Dim texto As String
    Dim celula As Range
    Dim celulaDestino As Range
    
    Set celulaDestino = Range("N4") 'define a primeira célula de destino como J4
    
     'apaga os dados da coluna J antes de copiar novos dados
    Range("N4:N" & Range("N" & Rows.Count).End(xlUp).Row).ClearContents
    
    'percorre todas as células com texto da coluna F a partir de F4
    For Each celula In Range("J4:J" & Range("J" & Rows.Count).End(xlUp).Row)
    
        'verifica se o texto da célula já foi copiado para a coluna J
        If WorksheetFunction.CountIf(Range("N4:N" & celulaDestino.Row - 1), celula.Value) = 0 Then
        
            'copia o texto para a próxima célula de destino e atualiza a variável celulaDestino
            celulaDestino.Value = celula.Value
            Set celulaDestino = celulaDestino.Offset(1, 0)
        End If
    Next celula
    
End Sub


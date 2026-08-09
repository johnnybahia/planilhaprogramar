Attribute VB_Name = "Módulo15"
Sub InserirFormulaSEISEMSEIS()

    Dim i As Integer
    Dim linha As Integer

    ' Ativa a pasta "Planejamento Trançadeira"
    Sheets("Planejamento Trançadeira").Activate
    
    ' Começa da linha 1 e vai até a linha 137
    For i = 1 To 137
        linha = (i - 1) * 6 + 19 ' Calcula a linha de referência no "RELATÓRIO TRANÇADEIRAS"
        
        ' Coloca a fórmula na célula da coluna A LETRA I CORRESPONDE A COLUNA QUE QUER ALTERAR
        Cells(i, 2).Formula = "='RELATÓRIO TRANÇADEIRAS'!D" & linha
    Next i

End Sub


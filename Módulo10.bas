Attribute VB_Name = "Módulo10"
Sub OcultarColunas()

' Declaração de variáveis
Dim ColunaInicial As Long
Dim ColunaFinal As Long

' Obter a coluna inicial do usuário
ColunaInicial = InputBox("Digite a coluna inicial:", "Ocultar Colunas")

' Obter a coluna final do usuário
ColunaFinal = InputBox("Digite a coluna final:", "Ocultar Colunas")

' Validar a entrada do usuário
If ColunaInicial > ColunaFinal Then
    MsgBox "Coluna inicial deve ser menor ou igual à coluna final.", vbExclamation
    Exit Sub
End If

' Ocultar as colunas
For i = ColunaInicial To ColunaFinal
    Columns(i).EntireColumn.Hidden = True
Next i

' Mensagem de confirmação
MsgBox "Colunas " & ColunaInicial & " a " & ColunaFinal & " foram ocultadas com sucesso.", vbInformation

End Sub


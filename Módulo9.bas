Attribute VB_Name = "Módulo9"
Sub OcultarLinhas()

' Declaração de variáveis
Dim LinhaInicial As Long
Dim LinhaFinal As Long
Dim rng As Range

' Obter a linha inicial do usuário
LinhaInicial = InputBox("Digite a linha inicial:", "Ocultar Linhas")

' Obter a linha final do usuário
LinhaFinal = InputBox("Digite a linha final:", "Ocultar Linhas")

' Validar a entrada do usuário
If LinhaInicial > LinhaFinal Then
    MsgBox "Linha inicial deve ser menor ou igual à linha final.", vbExclamation
    Exit Sub
End If

' Selecionar as linhas
Set rng = Range(Cells(LinhaInicial, 1), Cells(LinhaFinal, 1))
rng.Select

' Ocultar as linhas selecionadas
rng.EntireRow.Hidden = True

' Mensagem de confirmação
MsgBox "Linhas " & LinhaInicial & " a " & LinhaFinal & " foram ocultadas com sucesso.", vbInformation

End Sub


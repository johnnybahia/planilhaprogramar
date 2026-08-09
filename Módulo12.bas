Attribute VB_Name = "Módulo12"
Sub OcultarColunas_Otimizada()

' Declaração de variáveis
Dim ColunaInicial As Long
Dim ColunaFinal As Long
Dim rngColunas As Range

' Obter a coluna inicial e final do usuário
ColunaInicial = InputBox("Digite a coluna inicial:", "Ocultar Colunas")
ColunaFinal = InputBox("Digite a coluna final:", "Ocultar Colunas")

' Validar a entrada do usuário
If ColunaInicial > ColunaFinal Then
  MsgBox "Coluna inicial deve ser menor ou igual à coluna final.", vbExclamation
  Exit Sub
End If

' Desabilitar recursos desnecessários
Application.ScreenUpdating = False
Application.Calculation = xlManual
Application.DisplayStatusBar = False

' Ocultar as colunas
Set rngColunas = Range(Cells(1, ColunaInicial), Cells(Rows.Count, ColunaFinal))
rngColunas.EntireColumn.Hidden = True

' Habilitar recursos desnecessários
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.DisplayStatusBar = True

' Mensagem de confirmação
MsgBox "Colunas " & ColunaInicial & " a " & ColunaFinal & " foram ocultadas com sucesso.", vbInformation

End Sub


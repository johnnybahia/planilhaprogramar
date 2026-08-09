Attribute VB_Name = "Módulo1"
Private Sub Workbook_Open()
    Call AutoSave(10) 'intervalo de tempo em minutos
End Sub

Sub AutoSave(intervalo As Integer)
    Dim TimeToRun As Date
    TimeToRun = Now + TimeValue("00:" & intervalo & ":00")
    Application.OnTime TimeToRun, "AutoSaveNow"
End Sub

Sub AutoSaveNow()
    ThisWorkbook.Save
    Call AutoSave(10) 'intervalo de tempo em minutos
End Sub


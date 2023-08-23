VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
If InStr(1, ThisWorkbook.Name, "(No Links)", vbTextCompare) <> 0 Then
    Application.CalculateBeforeSave = False
Else
    Application.CalculateBeforeSave = True
End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
If InStr(1, ThisWorkbook.Name, "(No Links)", vbTextCompare) <> 0 Then
    Application.Calculation = xlCalculationAutomatic
Else
    Application.Calculation = xlCalculationManual
    Application.CalculateBeforeSave = False
End If
End Sub

Private Sub Workbook_Open()

If InStr(1, ThisWorkbook.Name, "(No Links)", vbTextCompare) <> 0 Then
    Application.Calculation = xlCalculationAutomatic
Else
    Application.Calculation = xlCalculationManual
End If
    
ActiveWorkbook.RefreshAll

End Sub

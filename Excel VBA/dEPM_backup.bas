Attribute VB_Name = "dEPM_backup"
Sub BackupDEPMfiles()

On Error Resume Next

Dim dt, dEPMfile, GetBook As String
Dim Loc As Variant

dt = Format((Now), "yyyy_mm_dd  hmm AM/PM")

GetBook = Replace(ThisWorkbook.Name, ".xlms", "")

dEPMfile = GetBook & "_" & dt & ".xlsm"
'save the workbook before backing up
Application.Calculation = xlCalculationManual
Application.CalculateBeforeSave = False
ThisWorkbook.Save

Dim dEPMSave(3) As String
    dEPMSave(0) = "\\prcfile\DeptFolders\FinAnalyst\Infor Cloudsuite\dEPM reports\Backups\" & dEPMfile
    dEPMSave(1) = "\\prcfile\Users\jgarcia\My Documents\Infor Cloudsuite\dEPM\Backups\" & dEPMfile
For Each Loc In dEPMSave()
    ThisWorkbook.SaveCopyAs (Loc)

'For Each Loc In DORSaveLocations()
'    Debug.Print (Loc)
Next
Application.CalculateBeforeSave = True
End Sub

Public Sub ConvertFormulas()
On Error Resume Next

Dim dEPMSheets(0 To 11) As String
Dim ws As Worksheet
Dim wb As Workbook
Dim sht As Variant
Dim rng As Range

dEPMSheets(0) = "Monthly Budget Detail"
dEPMSheets(1) = "Monthly Actual Detail"
dEPMSheets(2) = "Monthly Stats Detail"
dEPMSheets(3) = "Quarterly Stats Detail"
dEPMSheets(4) = "Balance Sheet"
dEPMSheets(5) = "Net BS Changes"
dEPMSheets(6) = "Quarterly Actual Detail"
dEPMSheets(7) = "Departments Hierarchy"
dEPMSheets(8) = "Account Hierarchy"
dEPMSheets(9) = "Scenario Hierarchy"
dEPMSheets(10) = "GL Systems Hierarchy"
dEPMSheets(11) = "Net Changes"

Beep
Convert = MsgBox("Conversion of dEPM formulas to static values will take place. Procced?", vbYesNo + vbDefaultButton2, "Convert formulas?")

' if yes then convert the formulas
If Convert = vbYes Then
    If InStr(1, ThisWorkbook.Name, "(No Links)") <> 0 Then
            Application.ScreenUpdating = False
            For Each sht In dEPMSheets
                Worksheets(sht).Activate
                Call ResetFilters
                Cells.Select
                Selection.Copy
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
            Next sht
            Beep
            MsgBox "Converted all dEPM sheets to static values!"
            Application.ScreenUpdating = True
    
    'stop the macro and tell user to change the filename
    Else
        Beep
        MsgBox "Please save and rename file with the suffix of '(No Links)' in the filename before breaking links!" _
        & Chr(13) & Chr(13) & "Example file name: 'dEPM - Cash FLow Formulas - January 2023 (No Links).xlsm'"
        Exit Sub
    End If
Else
    Exit Sub
End If

End Sub

Sub ResetFilters()
    On Error Resume Next
    ActiveSheet.ShowAllData
End Sub

Attribute VB_Name = "OpenDORFiles"
Option Explicit

Sub OpenFiles()

On Error Resume Next

Dim ws As Worksheet
Dim objCOMAddin As COMAddIn
Dim DORwb As Workbook
Dim cboDORFile As ComboBox
Dim DORfile, DORFilePath, DORDay As String
Dim DORDate, DORDate2 As Variant

Set ws = ThisWorkbook.Worksheets("DOR Central")
Set cboDORFile = ws.OLEObjects("OpenFile").Object

DORDay = Day(ws.Range("DOR_Date"))
DORfile = cboDORFile.Column(0)
DORFilePath = cboDORFile.Column(1)
DORDate = ws.Range("DOR_DATE")
'
'Debug.Print DORFile
'Debug.Print DORFilePath
'Debug.Print DORDay
'Debug.Print DORDate

Select Case DORfile

        Case Is = "Daily Flash Report"
            Set DORwb = Workbooks.Open(filename:=DORFilePath, UpdateLinks:=3, ReadOnly:=True)
            DORwb.Worksheets(DORDay).Select
            Range("A1").Select

        Case Is = "Daily Labor Report"
            Set DORwb = Workbooks.Open(filename:=DORFilePath)
                'to disable and re-enable the IBM data transfer Excel Extension
                Application.COMAddIns("DataTransfer.Addin.1").Connect = False
                Application.COMAddIns("DataTransfer.Addin.1").Connect = True

        Case Else
            Set DORwb = Workbooks.Open(filename:=DORFilePath, UpdateLinks:=3)


    End Select
End Sub



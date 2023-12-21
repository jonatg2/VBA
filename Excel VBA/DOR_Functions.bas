Attribute VB_Name = "DOR_Functions"
Option Explicit

Function Hold_Drop_Var(Actual, Budget As Variant) As Variant

Hold_Drop_Var = (Actual / Budget) - 1

End Function

Sub BackupDORfiles()

On Error Resume Next

Dim dt, DORfile As String
Dim Loc As Variant

dt = Format((Now), "yyyy_mm_dd  hmm AM/PM")

DORfile = "DOR Macros_PROD_" & dt & ".xlsm"

Dim DORSaveLocations(3) As String
    DORSaveLocations(0) = "\\prcfile\DeptFolders\FinAnalyst\DOR\Excel Macros\DOR Central Backups\" & DORfile
    DORSaveLocations(1) = "\\prcfile\DeptFolders\FinAnalyst\DOR Central Backups\" & DORfile
    DORSaveLocations(2) = "\\prcfile\Users\jgarcia\My Documents\Excel\DOR Central Backups\" & DORfile
For Each Loc In DORSaveLocations()
    ThisWorkbook.SaveCopyAs (Loc)

'For Each Loc In DORSaveLocations()
'    Debug.Print (Loc)
Next
End Sub
Function DORDateConvert(DORDate As Variant) As Variant

Select Case Format(DORDate, "MMMM")
    Case "October", "November", "December"
        DORDateConvert = DateAdd("yyyy", 1, DORDate)
    Case Else
        DORDateConvert = DORDate
End Select

End Function
Function DORDateSubtract(DORDate As Variant) As Variant

Select Case Format(DORDate, "MMMM")
    Case "October", "November", "December"
        DORDateSubtract = DateAdd("yyyy", -1, DORDate)
    Case Else
        DORDateSubtract = DORDate
End Select

End Function

Function ListLinksMinDate() As Date

Dim alinks As Variant
Dim DateArray(1 To 5), MinDate As Date
Dim i As Integer

alinks = ActiveWorkbook.LinkSources(xlExcelLinks)
    If Not IsEmpty(alinks) Then
        For i = 1 To UBound(alinks)
            DateArray(i) = CLng(DORFileDate(CStr(alinks(i))))
        Next i
    End If
MinDate = CDate(Application.WorksheetFunction.Min(DateArray))
ListLinksMinDate = MinDate
End Function

Function ListLinksMaxDate() As Date

Dim alinks As Variant
Dim DateArray(1 To 5), MaxDate As Date
Dim i As Integer

alinks = ActiveWorkbook.LinkSources(xlExcelLinks)
    If Not IsEmpty(alinks) Then
        For i = 1 To UBound(alinks)
            DateArray(i) = CLng(DORFileDate(CStr(alinks(i))))
        Next i
    End If
MaxDate = CDate(Application.WorksheetFunction.Max(DateArray))
ListLinksMaxDate = MaxDate
End Function


Function MonthCheck(StringVal As String) As Integer
Dim i As Integer, MonthStr As String
Dim Months(1 To 12) As String
    Months(1) = "January"
    Months(2) = "February"
    Months(3) = "March"
    Months(4) = "April"
    Months(5) = "May"
    Months(6) = "June"
    Months(7) = "July"
    Months(8) = "August"
    Months(9) = "September"
    Months(10) = "October"
    Months(11) = "November"
    Months(12) = "December"

For i = 1 To 12
    If InStr(StringVal, Months(i)) > 1 Then
        MonthStr = Month(DateValue("01 " & Months(i) & " 2019"))
    End If
Next i
MonthCheck = MonthStr

End Function

Function YearCheck(StringVal As String) As Integer
Dim i, j, YearStr As Integer
    i = 0
    j = 0
Do While j <> 1
YearStr = (Year(Now) + 1) - i
   If StringVal Like "*" & YearStr & "*" Then
        YearCheck = YearStr
        j = 1
    End If
   i = i + 1
Loop
End Function

Function DORFileDate(StringVal As String) As Date
DORFileDate = DateValue(MonthCheck(StringVal) & " 01 " & YearCheck(StringVal))

End Function
Function GetDesktop() As String
    Dim oWSHShell As Object

    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop")
    Set oWSHShell = Nothing
End Function

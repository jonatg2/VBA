Attribute VB_Name = "ListAllLinks"
Sub UpdateLinks()

Set ExcelFilePath_Daily = Worksheets("Setup").Range("FilePath_PROD")
Set ExcelFilePath_Daily_PREV = Worksheets("Setup").Range("FilePath_PROD_PREV")

    Dim alinks As Variant
    Dim DateArray(1 To 5), SingleDate As Date
    alinks = ActiveWorkbook.LinkSources(xlExcelLinks)
    
    If Not IsEmpty(alinks) Then
        For i = 1 To UBound(alinks)
            DateArray(i) = CLng(DORDateSubtract(DORFileDate(CStr(alinks(i)))))
        Next i
    End If

MaxDate = CDate(Application.WorksheetFunction.Max(DateArray))
MinDate = CDate(Application.WorksheetFunction.Min(DateArray))

Application.Wait (Now + TimeValue("00:00:05"))

' to loop through external links and update with new links
    If Not IsEmpty(alinks) Then
        For i = 1 To UBound(alinks)
            If CDate(CLng(DORDateSubtract(DORFileDate(CStr(alinks(i)))))) = MaxDate Then
                ActiveWorkbook.ChangeLink alinks(i), ExcelFilePath_Daily, xlLinkTypeExcelLinks
            ElseIf CDate(CLng(DORDateSubtract(DORFileDate(CStr(alinks(i)))))) = MinDate Then
                ActiveWorkbook.ChangeLink alinks(i), ExcelFilePath_Daily_PREV, xlLinkTypeExcelLinks

                
            End If
            Next i
   End If
    
'Debug.Print MaxDate
'Debug.Print MinDate
End Sub

Sub UpdateLookups()
Application.ScreenUpdating = False

Dim DORCurrentLink_NEW, DORCurrentLink_OLD, _
    DORCurrentLinkWeekly_NEW, DORCurrentLinkWeekly_OLD, _
    DORPreviousLink_NEW, DORPreviousLink_OLD As Range
    
Set DORCurrentLink_NEW = Worksheets("Setup").Range("DORCurrentLink_NEW")
Set DORCurrentLink_OLD = Worksheets("Setup").Range("DORCurrentLink_OLD")
Set DORPreviousLink_NEW = Worksheets("Setup").Range("DORPreviousLink_NEW")
Set DORPreviousLink_OLD = Worksheets("Setup").Range("DORPreviousLink_OLD")
Set DORCurrentLinkWeekly_NEW = Worksheets("Setup").Range("DORCurrentLinkWeekly_NEW")
Set DORCurrentLinkWeekly_OLD = Worksheets("Setup").Range("DORCurrentLinkWeekly_OLD")

' to replace current month formulas
Worksheets("Lookups").Columns("B").Replace _
What:=DORCurrentLink_OLD, Replacement:=DORCurrentLink_NEW, _
LookAt:=xlPart, SearchOrder:=xlByColumns

Worksheets("Lookups").Range("WeeklyDOR_ActualCheck").Replace _
What:=DORCurrentLinkWeekly_OLD, Replacement:=DORCurrentLinkWeekly_NEW, _
LookAt:=xlPart

Worksheets("Lookups").Range("WeeklyDOR_BudgetCheck").Replace _
What:=DORCurrentLinkWeekly_OLD, Replacement:=DORCurrentLinkWeekly_NEW, _
LookAt:=xlPart

Worksheets("DOR Central").Range("DOR_DATE_SS").Replace _
What:=DORCurrentLink_OLD, Replacement:=DORCurrentLink_NEW, _
LookAt:=xlPart, SearchOrder:=xlByColumns

Worksheets("DOR Central").Range("DOR_DATE_SS_WEEKLY").Replace _
What:=DORCurrentLinkWeekly_OLD, Replacement:=DORCurrentLinkWeekly_NEW, _
LookAt:=xlPart, SearchOrder:=xlByColumns

' to replace previous month formulas
Worksheets("Lookups").Columns("C").Replace _
What:=DORPreviousLink_OLD, Replacement:=DORPreviousLink_NEW, _
LookAt:=xlPart, SearchOrder:=xlByColumns

Application.ScreenUpdating = True
End Sub

Attribute VB_Name = "CreateDORFiles"
Option Explicit

Sub DORCreateFilesDaily()
On Error GoTo ErrHandler
'to disable the screen updating
Application.ScreenUpdating = False

Dim ExternalLinks As Variant
Dim DORwb, wb As Workbook
Dim DORws, DORCentralWS, ws As Worksheet
Dim x As Long
Dim WorkbookPath As Variant
Dim filename As Variant
Dim DORExcelSavePath, DORExcelSavePath_Desktop, DORExcelHyperLink, DORExcelSaveLocation, ExcelStaticDaily  As Variant
Dim DORPDFSavePath, DORPDFSaveLocation As Variant
Dim DORDate, DORDate_SS As Variant
Dim CreateDORFiles, OverwriteDORFiles As Integer
Dim HideYTD As Object


Set DORDate = Application.Range("DOR_Date")
Set DORDate_SS = Application.Range("DOR_Date_SS")
Set DORExcelSavePath = Worksheets("Setup").Range("DORSavePath")
Set DORCentralWS = ThisWorkbook.Worksheets("DOR Central")
Set HideYTD = DORCentralWS.OLEObjects("HideYTD").Object

Beep
CreateDORFiles = MsgBox("PDF and Excel DOR Daily Files will be created for " & Format(DORDate, "DDDD") & ", " & DORDate & ". Do you wish to continue?", vbYesNoCancel, "Create DOR Files?")
    If CreateDORFiles = vbYes Then
        If DORDate <> DORDate_SS Then
                Beep
                MsgBox "The date on the DOR file does not match the date you wish to generate the file for. " _
                & "Please change the date!", vbCritical, "DOR Dates don't match!"
                Exit Sub
        ElseIf FileExists(DORExcelSavePath) Then
            Beep
            OverwriteDORFiles = MsgBox("The file already exists for " & Format(DORDate, "DDDD") & ", " & DORDate & ". Do you wish to overwrite these files?", vbYesNoCancel, "Overwrite DOR Files?")
                 If OverwriteDORFiles = vbYes Then
                
                'Code executes to overwrite the DOR file
                    Set WorkbookPath = Worksheets("Setup").Range("FilePath_PROD")
                    Set DORExcelSavePath = Worksheets("Setup").Range("DORSavePath")
                    Set DORExcelSavePath_Desktop = Worksheets("Setup").Range("DORSavePath_Desktop")
                    Set DORExcelSaveLocation = Worksheets("Setup").Range("DORExcelSaveLocation")
                    Set DORExcelHyperLink = Worksheets("Setup").Range("DORHyperlink")
                    Set DORPDFSavePath = Worksheets("Setup").Range("PDF_FileSavePath")
                    Set DORPDFSaveLocation = Worksheets("Setup").Range("PDFSaveLocation")
                    Set ExcelStaticDaily = Worksheets("Setup").Range("ExcelStatic_Daily")
                    
                   'to delete any previous copies of the DOR file that may already be in the directory
                    Call DeleteDORFile(DORPDFSavePath)
                    Call DeleteDORFile(DORExcelSavePath)
                    
                    
                    Set DORwb = Workbooks.Open(WorkbookPath, UpdateLinks:=0)
                    Set DORws = DORwb.Sheets("DOR")
                    
                                'Copy the DOR sheet to new workbook and close DOR workbook without saving
                                    
                                    ActiveSheet.Copy
                                    Set ws = ActiveSheet
                                    Set wb = ActiveWorkbook
                                    DORwb.Close (False)
                                    Application.DisplayAlerts = False
                                    'wb.UpdateLink Name:=wb.LinkSources, Type:=1
                           
                                'Create an Array of all External Links stored in Workbook
                                      ExternalLinks = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
                                    
                                'Loop Through each External Link in ActiveWorkbook and Break it
                                      For x = 1 To UBound(ExternalLinks)
                                        wb.BreakLink Name:=ExternalLinks(x), Type:=xlLinkTypeExcelLinks
                                      Next x
                                      
                              'Hide uneeded columns
                                    Columns("BM:DA").EntireColumn.Hidden = True
                                    Columns("AS").EntireColumn.Hidden = True

                                    
                                    If HideYTD = True Then
                                        Columns("BC:BF").EntireColumn.Hidden = True
                                    End If
                                'Collapse all rows and columns
                                    ActiveSheet.Outline.ShowLevels RowLevels:=1 ' to collapse the rows
                                    ActiveSheet.Outline.ShowLevels ColumnLevels:=1 'to collapse the columns
                                    
                                'to remove conditional formatting
                                    With ws
                                       Range("A1:ZZ1000").FormatConditions.Delete
                                    End With
                                    
                                'to set page break preview
                                        With ws
                                            ActiveWindow.View = xlPageBreakPreview
                                        End With
                                    
                                'to set selected cell to proper location and set zoom level
                                    Application.ScreenUpdating = True
                                        With ws
                                            .Range("AH4").Select
                                            ActiveWindow.Zoom = 80
                                            Application.SendKeys "^{HOME}", 3
                                        End With
                                    Application.ScreenUpdating = False
                                        
                                'to uncheck workbook compatibility
                                        With wb
                                            .CheckCompatibility = False
                                        End With
                                            
                                    
                                'to save a copy of the file as a pdf
                                    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                        filename:=DORPDFSavePath
                                        
                                        
                                'to save a copy of the file as xls (Excel 97 version)
                                Shell "explorer.exe " & DORExcelSaveLocation, vbNormalFocus  'open Excel file location
                                      With wb
                                            .CheckCompatibility = False
                                            .SaveAs DORExcelSavePath, 56
                                            .Close
                                        End With
                                        
                                'to open the file locations and the actual files for the DOR PDF and Excel files
                                    Beep
                                    MsgBox "PDF and Excel DOR files have been created! Opening file locations and files now.", vbInformation, "Files Created"
                                    Shell "explorer.exe " & DORPDFSaveLocation, vbNormalFocus  'open PDF file location
                                    Shell "explorer.exe " & DORExcelSaveLocation, vbNormalFocus  'open Excel file location
                                    Shell "explorer.exe " & DORPDFSavePath & "", vbNormalFocus  'open PDF file
                                    Workbooks.Open DORExcelSavePath 'open Excel file

                           Else
                                 Exit Sub
                        End If
            Else 'Code executes to create the DOR file
                Set WorkbookPath = Worksheets("Setup").Range("FilePath_PROD")
                    Set DORExcelSavePath = Worksheets("Setup").Range("DORSavePath")
                    Set DORExcelSavePath_Desktop = Worksheets("Setup").Range("DORSavePath_Desktop")
                    Set DORExcelSaveLocation = Worksheets("Setup").Range("DORExcelSaveLocation")
                    Set DORExcelHyperLink = Worksheets("Setup").Range("DORHyperlink")
                    Set DORPDFSavePath = Worksheets("Setup").Range("PDF_FileSavePath")
                    Set DORPDFSaveLocation = Worksheets("Setup").Range("PDFSaveLocation")
                    Set ExcelStaticDaily = Worksheets("Setup").Range("ExcelStatic_Daily")
                    
                   'to delete any previous copies of the DOR file that may already be in the directory
                    Call DeleteDORFile(DORPDFSavePath)
                    Call DeleteDORFile(DORExcelSavePath)
                    
                    
                    Set DORwb = Workbooks.Open(WorkbookPath, UpdateLinks:=0)
                    Set DORws = DORwb.Sheets("DOR")
                    
                                'Copy the DOR sheet to new workbook and close DOR workbook without saving
                                    
                                    ActiveSheet.Copy
                                    Set ws = ActiveSheet
                                    Set wb = ActiveWorkbook
                                    DORwb.Close (False)
                                    Application.DisplayAlerts = False
                                    'wb.UpdateLink Name:=wb.LinkSources, Type:=1
                           
                                'Create an Array of all External Links stored in Workbook
                                      ExternalLinks = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
                                    
                                'Loop Through each External Link in ActiveWorkbook and Break it
                                      For x = 1 To UBound(ExternalLinks)
                                        wb.BreakLink Name:=ExternalLinks(x), Type:=xlLinkTypeExcelLinks
                                      Next x
                                      
                                'Hide uneeded columns
                                    Columns("BM:DA").EntireColumn.Hidden = True
                                    Columns("AS").EntireColumn.Hidden = True
                                    
                                    
                                    If HideYTD = True Then
                                        Columns("BC:BF").EntireColumn.Hidden = True
                                    End If
                                    
                                    'Collapse all rows and columns
                                    ActiveSheet.Outline.ShowLevels RowLevels:=1 ' to collapse the rows
                                    ActiveSheet.Outline.ShowLevels ColumnLevels:=1 'to collapse the columns
                                    
                                'to remove conditional formatting
                                    With ws
                                       Range("A1:ZZ1000").FormatConditions.Delete
                                    End With
                                    
                                'to set page break preview
                                        With ws
                                            ActiveWindow.View = xlPageBreakPreview
                                        End With
                                    
                                'to set selected cell to proper location and set zoom level
                                    Application.ScreenUpdating = True
                                        With ws
                                            .Range("AH4").Select
                                            ActiveWindow.Zoom = 80
                                            Application.SendKeys "^{HOME}", 3
                                        End With
                                    Application.ScreenUpdating = False
                                        
                                'to uncheck workbook compatibility
                                        With wb
                                            .CheckCompatibility = False
                                        End With
                                            
                                    
                                'to save a copy of the file as a pdf
                                    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                        filename:=DORPDFSavePath
                                        
                                        
                              'to save a copy of the file as xls (Excel 97 version)
                                Shell "explorer.exe " & DORExcelSaveLocation, vbNormalFocus  'open Excel file location
                                      With wb
                                            .CheckCompatibility = False
                                            .SaveAs DORExcelSavePath, 56
                                            .Close
                                        End With
                                        
                                'to open the file locations and the actual files for the DOR PDF and Excel files
                                    Beep
                                    MsgBox "PDF and Excel DOR files have been created! Opening file locations and files now.", vbInformation, "Files Created"
                                    Shell "explorer.exe " & DORPDFSaveLocation, vbNormalFocus  'open PDF file location
                                    Shell "explorer.exe " & DORExcelSaveLocation, vbNormalFocus  'open Excel file location
                                    Shell "explorer.exe " & DORPDFSavePath & "", vbNormalFocus  'open PDF file
                                    Workbooks.Open DORExcelSavePath 'open Excel file

                Exit Sub
            End If
        Else
             Exit Sub
    End If
 Application.ScreenUpdating = True
ErrHandler:
    If Err.Number = 1004 Then
        With wb
            .CheckCompatibility = False
            .SaveAs DORExcelSavePath_Desktop, 56
            .Close
        End With
    End If
    Resume Next
End Sub


Sub DORCreateFilesWeekly()

'to disable the screen updating
Application.ScreenUpdating = False

Dim ExternalLinks As Variant
Dim DORwb, wb As Workbook
Dim DORws, ws As Worksheet
Dim x As Long
Dim WorkbookPath As Variant
Dim filename As Variant
Dim DORExcelSavePath, DORExcelHyperLink, DORExcelSaveLocation  As Variant
Dim DORPDFSavePath, DORPDFSaveLocation As Variant
Dim DORDate, DORDate_SS, ActualCheck, BudgetCheck As Variant
Dim CreateDORFiles, OverwriteDORFiles As Integer

Set DORDate = Application.Range("DOR_Date_Weekly")
Set DORDate_SS = Application.Range("DOR_Date_SS_WEEKLY")
Set ActualCheck = Application.Range("WeeklyDOR_ActualCheck")
Set BudgetCheck = Application.Range("WeeklyDOR_BudgetCheck")
Set DORExcelSavePath = Worksheets("Setup").Range("DORSavePath_Weekly")

Beep
CreateDORFiles = MsgBox("PDF and Excel DOR Weekly Files will be created for " & Format(DORDate, "DDDD") & ", " & DORDate & ". Do you wish to continue?", vbYesNoCancel, "Create DOR Files?")
    If CreateDORFiles = vbYes Then
        If DORDate <> DORDate_SS Then
            Beep
            MsgBox "The date on the DOR file does not match the date you wish to generate the file for. " _
            & "Please change the date!", vbCritical, "DOR Dates don't match!"
            Exit Sub
        ElseIf ActualCheck <> "OK" Or BudgetCheck <> "OK" Then
            Beep
            MsgBox "Check the Weekly DOR MTD Actual and Budgeted Amounts! They do not match with the DOR Daily MTD amounts" _
            & " OR the Daily DOR Date does not match the Weekly DOR Date.", vbCritical, "Check MTD Amounts"
            Exit Sub
        ElseIf FileExists(DORExcelSavePath) Then
            Beep
            OverwriteDORFiles = MsgBox("The file already exists for " & Format(DORDate, "DDDD") & ", " & DORDate & ". Do you wish to overwrite these files?", vbYesNoCancel, "Overwrite DOR Files?")
                 If OverwriteDORFiles = vbYes Then
                   'to overwrite any previous DOR files
                    Set WorkbookPath = Worksheets("Setup").Range("FilePath_Weekly_PROD")
                    Set DORExcelSavePath = Worksheets("Setup").Range("DORSavePath_Weekly")
                    Set DORExcelSaveLocation = Worksheets("Setup").Range("DORExcelSaveLocation")
                    Set DORExcelHyperLink = Worksheets("Setup").Range("DORHyperlink_Weekly")
                    Set DORPDFSavePath = Worksheets("Setup").Range("PDF_FileSavePath_Weekly")
                    Set DORPDFSaveLocation = Worksheets("Setup").Range("PDFSaveLocation")
                    
                    'to delete any previous copies of the DOR file that may already be in the directory
                    Call DeleteDORFile(DORPDFSavePath)
                    Call DeleteDORFile(DORExcelSavePath)
                    
                    Set DORwb = Workbooks.Open(WorkbookPath, UpdateLinks:=3)
                    Set DORws = DORwb.Sheets("DOR")
                    
                                'Copy the DOR sheet to new workbook and close DOR workbook without saving
                                    
                                    ActiveSheet.Copy
                                    Set ws = ActiveSheet
                                    Set wb = ActiveWorkbook
                                    DORwb.Close (False)
                                    Application.DisplayAlerts = False
                                    wb.UpdateLink Name:=wb.LinkSources, Type:=1
                           
                                'Create an Array of all External Links stored in Workbook
                                      ExternalLinks = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
                                    
                                'Loop Through each External Link in ActiveWorkbook and Break it
                                      For x = 1 To UBound(ExternalLinks)
                                        wb.BreakLink Name:=ExternalLinks(x), Type:=xlLinkTypeExcelLinks
                                      Next x

                                'Collapse all rows and columns
                                    ActiveSheet.Outline.ShowLevels RowLevels:=1 ' to collapse the rows
                                    ActiveSheet.Outline.ShowLevels ColumnLevels:=1 'to collapse the columns
                                    
                                'to remove conditional formatting
                                    With ws
                                       Range("A1:ZZ1000").FormatConditions.Delete
                                    End With
                                    
                                'to set page break preview
                                        With ws
                                            ActiveWindow.View = xlPageBreakPreview
                                        End With
                                    
                                'to set selected cell to proper location and set zoom level
                                        With ws
                                            .Range("BM5").Select
                                            ActiveWindow.Zoom = 80
                                            Application.SendKeys "^{HOME}"
                                        End With
                                        
                                'to uncheck workbook compatibility
                                        With wb
                                            .CheckCompatibility = False
                                        End With
                                            
                                    
                                'to save a copy of the file as a pdf
                                    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                        filename:=DORPDFSavePath
                                        
                                        
                                'to save a copy of the file as xls (Excel 97 version) and delete any previous versions of file
                                        With wb
                                            .CheckCompatibility = False
                                            .SaveAs DORExcelSavePath, 56
                                            .Close
                                        End With
                                        
                                'to open the file locations and the actual files for the DOR PDF and Excel files
                                    Beep
                                    MsgBox "PDF and Excel DOR files have been created! Opening file locations and files now.", vbInformation, "Files Created"
                                    Shell "explorer.exe " & DORPDFSaveLocation, vbNormalFocus  'open PDF file location
                                    Shell "explorer.exe " & DORExcelSaveLocation, vbNormalFocus  'open Excel file location
                                    Shell "explorer.exe " & DORPDFSavePath & "", vbNormalFocus  'open PDF file
                                    Workbooks.Open DORExcelSavePath 'open Excel file
                            
                        Else
                            Exit Sub
                        End If
        Else
                 'to create DOR files
                    Set WorkbookPath = Worksheets("Setup").Range("FilePath_Weekly_PROD")
                    Set DORExcelSavePath = Worksheets("Setup").Range("DORSavePath_Weekly")
                    Set DORExcelSaveLocation = Worksheets("Setup").Range("DORExcelSaveLocation")
                    Set DORExcelHyperLink = Worksheets("Setup").Range("DORHyperlink_Weekly")
                    Set DORPDFSavePath = Worksheets("Setup").Range("PDF_FileSavePath_Weekly")
                    Set DORPDFSaveLocation = Worksheets("Setup").Range("PDFSaveLocation")
                    
                    'to delete any previous copies of the DOR file that may already be in the directory
                    Call DeleteDORFile(DORPDFSavePath)
                    Call DeleteDORFile(DORExcelSavePath)
                    
                    Set DORwb = Workbooks.Open(WorkbookPath, UpdateLinks:=3)
                    Set DORws = DORwb.Sheets("DOR")
                    
                                'Copy the DOR sheet to new workbook and close DOR workbook without saving
                                    
                                    ActiveSheet.Copy
                                    Set ws = ActiveSheet
                                    Set wb = ActiveWorkbook
                                    DORwb.Close (False)
                                    Application.DisplayAlerts = False
                                    wb.UpdateLink Name:=wb.LinkSources, Type:=1
                           
                                'Create an Array of all External Links stored in Workbook
                                      ExternalLinks = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
                                    
                                'Loop Through each External Link in ActiveWorkbook and Break it
                                      For x = 1 To UBound(ExternalLinks)
                                        wb.BreakLink Name:=ExternalLinks(x), Type:=xlLinkTypeExcelLinks
                                      Next x
    
                                'Collapse all rows and columns
                                    ActiveSheet.Outline.ShowLevels RowLevels:=1 ' to collapse the rows
                                    ActiveSheet.Outline.ShowLevels ColumnLevels:=1 'to collapse the columns
                                    
                                'to remove conditional formatting
                                    With ws
                                       Range("A1:ZZ1000").FormatConditions.Delete
                                    End With
                                    
                                'to set page break preview
                                        With ws
                                            ActiveWindow.View = xlPageBreakPreview
                                        End With
                                    
                                'to set selected cell to proper location and set zoom level
                                        With ws
                                            .Range("BM5").Select
                                            ActiveWindow.Zoom = 80
                                            Application.SendKeys "^{HOME}"
                                        End With
                                        
                                'to uncheck workbook compatibility
                                        With wb
                                            .CheckCompatibility = False
                                        End With
                                            
                                    
                                'to save a copy of the file as a pdf
                                    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
                                        filename:=DORPDFSavePath
                                        
                                        
                                'to save a copy of the file as xls (Excel 97 version) and delete any previous versions of file
                                        With wb
                                            .CheckCompatibility = False
                                            .SaveAs DORExcelSavePath, 56
                                            .Close
                                        End With
                                        
                                'to open the file locations and the actual files for the DOR PDF and Excel files
                                    Beep
                                    MsgBox "PDF and Excel DOR files have been created! Opening file locations and files now.", vbInformation, "Files Created"
                                    Shell "explorer.exe " & DORPDFSaveLocation, vbNormalFocus  'open PDF file location
                                    Shell "explorer.exe " & DORExcelSaveLocation, vbNormalFocus  'open Excel file location
                                    Shell "explorer.exe " & DORPDFSavePath & "", vbNormalFocus  'open PDF file
                                    Workbooks.Open DORExcelSavePath 'open Excel file
                            
                            
                        End If
            Else
                Exit Sub
     End If
 Application.ScreenUpdating = True
End Sub


Sub FileSelector()

 If Worksheets("DOR Central").Excel_Daily.Value = True Then
    Call DORCreateFilesDaily
    
ElseIf Worksheets("DOR Central").Weekly_Excel_PDF.Value = True Then
    Call DORCreateFilesWeekly
 
            Else
                Exit Sub
End If

End Sub




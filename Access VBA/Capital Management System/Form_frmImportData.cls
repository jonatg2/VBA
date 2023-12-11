VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cboDivision_Change()
If IsNull(Me.cboDivision) = False Then
    CurrentDb().QueryDefs("qryBudgetDataRead").SQL = "SELECT * FROM [excel_BudgetDataExport_" & Me.cboDivision.Value & "]"
    CurrentDb().QueryDefs("qryProjectDataRead").SQL = "SELECT * FROM [excel_ProjectDataExport_" & Me.cboDivision.Value & "] WHERE [CarNumber] <> 'Total'"
End If
End Sub

Private Sub cmdCreateEmail_Click()

If IsNull(Me.cboDivision) = True Then
    Beep
    MsgBox "Please select a division to export!", vbCritical, "Select a Division"
    Me.cboDivision.SetFocus
    Me.cboDivision.Dropdown
    Exit Sub
ElseIf IsNull(Me.cboFiscalYear) = True Then
    Beep
    MsgBox "Please select a fiscal year!", vbCritical, "Select a Fiscal Year"
    Me.cboFiscalYear.SetFocus
    Me.cboFiscalYear.Dropdown
    Exit Sub
ElseIf IsNull(Me.cboType) = True Then
    Beep
    MsgBox "Please select a project type!", vbCritical, "Select a Project Type"
    Me.cboType.SetFocus
    Me.cboType.Dropdown
    Exit Sub
End If

Call CreateDivisionEmail(Me.cboDivision.Value, Me.cboFiscalYear.Value, Me.cboType.Value)

End Sub

Private Sub cmdExportData_Click()
On Error GoTo ErrHandler

Dim rootFolder, rootFolderProject, stagingFolder, strFilePath, strFilePathProject, msg As String
Dim fs As Object
Dim recordCount, recordCountProject As Integer
Dim tbl As ListObject
Dim appExcel As Excel.Application
Dim myWorkbook As Excel.Workbook
Dim rng As Range
Dim Cell As Variant

'Budget Variables
rootFolder = "\\kiicha\DavWWWRoot\sites\leadership\PA\Shared Documents\Capital Budget by Division\"
stagingFolder = "\\prcfile\DeptFolders\FinAnalyst\Capital\Capital Tracking Database\Staging\Division Budget Files\"
strFilePath = stagingFolder & "qryBudgetDataExport_" & Me.cboDivision.Value & ".xlsx"
recordCount = DCount("[budgetReference]", "qryBudgetDataExport")

'Project Variables

rootFolderProject = "\\kiicha\DavWWWRoot\sites\leadership\PA\Shared Documents\Capital Projects by Division\"
strFilePathProject = stagingFolder & "qryProjectDataExport_" & Me.cboDivision.Value & ".xlsx"
recordCountProject = DCount("[CarNumber]", "qryProjectDataExport")

If IsNull(Me.cboImportExportFiles) = True Then
    Beep
    MsgBox "Please select a value from the Files dropdown.", vbInformation, "Select Value"
    Me.cboImportExportFiles.SetFocus
    Me.cboImportExportFiles.Dropdown
End If

Select Case Me.cboImportExportFiles.Value
    'Export project information to Sharepoint
    Case "Project Updates by Division (Excel)"
        If IsNull(Me.cboDivision) = True Then
            Beep
            MsgBox "Please select a division to export!", vbCritical, "Select a Division"
            Me.cboDivision.SetFocus
            Me.cboDivision.Dropdown
            Exit Sub
        ElseIf IsNull(Me.cboFiscalYear) = True And Me.cboType.Value <> "Approved Projects" Then
            Beep
            MsgBox "Please select a fiscal year!", vbCritical, "Select a Fiscal Year"
            Me.cboFiscalYear.SetFocus
            Me.cboFiscalYear.Dropdown
            Exit Sub
        ElseIf IsNull(Me.cboType) = True Then
            Beep
            MsgBox "Please select a project type!", vbCritical, "Select a Project Type"
            Me.cboType.SetFocus
            Me.cboType.Dropdown
            Exit Sub
        End If
        
        Select Case Me.cboType.Value
            Case "Budgeted Projects"
                'Import data before sending out
                Call UpdateBudgetNotes(userName(), "Budget Projects")
                
                'Export budget query to staging folder
                DoCmd.OutputTo acOutputQuery, "qryBudgetDataExport", acFormatXLSX, stagingFolder & "qryBudgetDataExport_" & Me.cboDivision.Value & ".xlsx", Autostart:=False
                
                'Open Excel file and apply table formatting
                
                Set appExcel = CreateObject("Excel.Application")
                Set myWorkbook = appExcel.Workbooks.Open(strFilePath, , False)
                appExcel.Visible = False
                appExcel.Application.ScreenUpdating = False
                Set rng = myWorkbook.Worksheets("qryBudgetDataExport").Range("A1:J" & recordCount + 1)
                rng.ClearFormats
                myWorkbook.Worksheets("qryBudgetDataExport").Range("C2:E" & recordCount + 1).NumberFormat = "$#,##0.00"
                myWorkbook.Worksheets("qryBudgetDataExport").Range("I2:I" & recordCount + 1).NumberFormat = "mm/dd/yyyy"
                Set tbl = myWorkbook.ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
                tbl.TableStyle = "TableStyleMedium16"
                myWorkbook.Save
                appExcel.Application.ScreenUpdating = True
                Set rng = Nothing
                myWorkbook.Close
                Set myWorkbook = Nothing
                appExcel.Quit
                Set appExcel = Nothing
                'Set tbl = Nothing
        
                
                'Close Excel file and move to Sharepoint folder
                Set fs = CreateObject("Scripting.FileSystemObject")
                fs.CopyFile stagingFolder & "qryBudgetDataExport_" & Me.cboDivision.Value & ".xlsx", rootFolder & "qryBudgetDataExport_" & Me.cboDivision.Value & ".xlsx"
                Beep
                MsgBox Me.cboDivision.Value & " Export Complete!"
            
             Case "Approved Projects"
                'Import data before sending out
                Call UpdateBudgetNotes(userName(), "Approved Projects")
                
                'Export budget query to staging folder
                DoCmd.OutputTo acOutputQuery, "qryProjectDataExport", acFormatXLSX, stagingFolder & "qryProjectDataExport_" & Me.cboDivision.Value & ".xlsx", Autostart:=False
                
                'Open Excel file and apply table formatting
                
                Set appExcel = CreateObject("Excel.Application")
                Set myWorkbook = appExcel.Workbooks.Open(strFilePathProject, , False)
                appExcel.Visible = False
                appExcel.Application.ScreenUpdating = False
                Set rng = myWorkbook.Worksheets("qryProjectDataExport").Range("A1:L" & recordCountProject + 1)
                rng.ClearFormats
                For Each Cell In myWorkbook.Worksheets("qryProjectDataExport").Range("L2:L" & recordCountProject + 1)
                    Cell.Hyperlinks.Add Anchor:=Cell, Address:=Cell.Value, TextToDisplay:=Cell.Offset(0, -11).Value & " Approved CAR PDF"
                Next Cell
                
                myWorkbook.Worksheets("qryProjectDataExport").Range("C2:E" & recordCountProject + 1).Select
                With myWorkbook.Selection
                    .NumberFormat = "General"
                    .Value = .Value
                End With
                myWorkbook.Worksheets("qryProjectDataExport").Range("C2:E" & recordCountProject + 1).NumberFormat = "$#,##0.00"
                myWorkbook.Worksheets("qryProjectDataExport").Range("H2:J" & recordCountProject + 1).NumberFormat = "mm/dd/yyyy"
                Set tbl = myWorkbook.ActiveSheet.ListObjects.Add(xlSrcRange, rng, , xlYes)
                tbl.TableStyle = "TableStyleMedium16"
                tbl.ShowTotals = True
                tbl.ListColumns("Car Amount").TotalsCalculation = xlTotalsCalculationSum
                tbl.ListColumns("Spend LTD").TotalsCalculation = xlTotalsCalculationSum
                tbl.ListColumns("Remaining Balance").TotalsCalculation = xlTotalsCalculationSum
                tbl.ListColumns("ProjectDescription").TotalsCalculation = xlTotalsCalculationCount
                myWorkbook.Save
                appExcel.Application.ScreenUpdating = True
                Set rng = Nothing
                myWorkbook.Close
                Set myWorkbook = Nothing
                appExcel.Quit
                Set appExcel = Nothing
                'Set tbl = Nothing
        
                
                'Close Excel file and move to Sharepoint folder
                Set fs = CreateObject("Scripting.FileSystemObject")
                fs.CopyFile stagingFolder & "qryProjectDataExport_" & Me.cboDivision.Value & ".xlsx", rootFolderProject & "qryProjectDataExport_" & Me.cboDivision.Value & ".xlsx"
                Beep
                MsgBox Me.cboDivision.Value & " Export Complete!"
            
        End Select
ErrHandler:
 If Err.Number = -2147024676 Then
  msg = "Someone in Sharepoint online or in the P&A Sharepoint folder is currently editing the " & Me.cboDivision.Value & " Budget file." _
  & " Please notify the user(s) to save any edits and close the file."
   Beep
   MsgBox msg, , "Error", Err.HelpFile, Err.HelpContext
   Exit Sub
ElseIf Err.Number = 70 Then
    Beep
    MsgBox "The View SharePoint Data query is still open. Please close out the query."
    Exit Sub
 ElseIf Err.Number = 0 Then
    'Resume Next
 Else
    msg = "Error # " & Str(Err.Number) & " was generated by " _
         & Err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & Err.Description
   Beep
   MsgBox msg, , "Error", Err.HelpFile, Err.HelpContext
    myWorkbook.Close
    Set myWorkbook = Nothing
    appExcel.Quit
    Set appExcel = Nothing
   Exit Sub
End If

End Select
End Sub

Private Sub cmdImport_Click()

Dim ReadytoImport As Integer

If IsNull(Me.cboImportExportFiles) = True Then
    Beep
    MsgBox "Please select a value from the Files dropdown.", vbInformation, "Select Value"
    Me.cboImportExportFiles.SetFocus
    Me.cboImportExportFiles.Dropdown
    Exit Sub
End If

Select Case Me.cboImportExportFiles.Value

    'Capital Spending File data import from ISD
    Case "Capital Spending File (ISD)"
        If IsNull(Me.txtImportDate) = True Then
            Beep
            MsgBox "Please enter an import date and set for the end of the month.", vbInformation, "Enter Import Date"
            Exit Sub
        End If
        
        Beep
        ReadytoImport = MsgBox("Capital Spending records for " & Me.txtImportDate.Value & " will be imported. Continue?", vbYesNo + vbDefaultButton2, "Import Data?")
        If ReadytoImport = vbYes Then
            DoCmd.RunSQL ("DELETE FROM PECHANGA\jgarcia_factCapitalSpending WHERE MonthEndDate = #" & Me.txtImportDate.Value & "#")
            DoCmd.OpenQuery ("qryAppendCapitalSpending")
            DoCmd.OpenQuery ("qryAppendTransferstoFA")
            Beep
            MsgBox "CapEx records have been imported!", vbInformation, "Import Completed"
        Else
            Exit Sub
        End If
    
    'Project Updates import from Excel (Sharepoint)
    Case "Project Updates by Division (Excel)"
        If IsNull(Me.cboDivision) = True Then
            Beep
            MsgBox "Please select a division to import!", vbCritical, "Please select a Division"
            Me.cboDivision.SetFocus
            Me.cboDivision.Dropdown
            Exit Sub
        ElseIf IsNull(Me.cboType) = True Then
            Beep
            MsgBox "Please select a project type!", vbCritical, "Please select a Project Type"
            Me.cboType.SetFocus
            Me.cboType.Dropdown
            Exit Sub
        End If
        
        Select Case Me.cboType.Value
            Case "Budgeted Projects"
                Beep
                ReadytoImport = MsgBox("Budget dates and notes will be imported for " & Me.cboDivision.Value & ". Continue?", vbYesNo + vbDefaultButton2, "Import Budget Notes?")
                If ReadytoImport = vbYes Then
                    Call UpdateBudgetNotes(userName(), "Budget Projects")
                    Beep
                    MsgBox "Data imported for " & Me.cboDivision.Value & "!", vbInformation, "Budget Notes Imported"
                Else
                    Exit Sub
                End If
            Case "Approved Projects"
                ReadytoImport = MsgBox("Project dates and notes will be imported for " & Me.cboDivision.Value & ". Continue?", vbYesNo + vbDefaultButton2, "Import Project Notes?")
                If ReadytoImport = vbYes Then
                    Call UpdateBudgetNotes(userName(), "Approved Projects")
                    Beep
                    MsgBox "Data imported for " & Me.cboDivision.Value & "!", vbInformation, "Project Notes Imported"
                Else
                    Exit Sub
                End If
        End Select
    
    'Capital Budget Import from Excel
    Case "Capital Budget Import (Excel)"
        If IsNull(Me.cboFiscalYear) = True Then
            Beep
            MsgBox "Please select/enter a fiscal year!", vbCritical, "Missing Fiscal Year"
            Me.cboFiscalYear.SetFocus
            Me.cboFiscalYear.Dropdown
            Exit Sub
        Else
            Beep
            ReadytoImport = MsgBox("Capital budget will be imported for FY " & Me.cboFiscalYear.Value & ". Continue?", vbYesNo + vbDefaultButton2, "Import Capital Budget?")
                    If ReadytoImport = vbYes Then
                        DoCmd.RunSQL ("DELETE FROM PECHANGA\jgarcia_factCapitalBudget WHERE FiscalYear = " & Me.cboFiscalYear.Value)
                        DoCmd.OpenQuery "qryAppendCapitalProjectsBudget"
                        Beep
                        MsgBox "Capital Budget Data imported for " & Me.cboFiscalYear.Value & "!", vbInformation, "Capital Budget Imported"
                    Else
                        Exit Sub
                    End If
        End If
    
    'Add new Capital Carryover Projects
    Case "Capital Carryover Update"
        If IsNull(Me.txtImportDate) = True Then
            Beep
            MsgBox "Please enter an import date and set for the end of the month.", vbInformation, "Enter Import Date"
            Me.txtImportDate.SetFocus
            Exit Sub
        ElseIf IsNull(Me.cboFiscalYear) = True Then
            Beep
            MsgBox "Please enter the new fiscal year!", vbCritical, "Missing Fiscal Year"
            Me.cboFiscalYear.SetFocus
        Else
            Beep
            ReadytoImport = MsgBox("CarryOver projects will be created for FY " & Me.cboFiscalYear.Value & ". Continue?", vbYesNo + vbDefaultButton2, "Add CarryOver Projects?")
                    If ReadytoImport = vbYes Then
                        DoCmd.OpenForm "frmReportNavigation", acNormal, , , , acHidden
                        Forms!frmReportNavigation!txtMonthEndDate.Value = Me.txtImportDate.Value
                        Forms!frmReportNavigation!cboFiscalYear.Value = Me.cboFiscalYear.Value - 1
                        DoCmd.RunSQL ("DELETE FROM PECHANGA\jgarcia_factCapitalCarryOver WHERE FiscalYearCarryOver = " & Me.cboFiscalYear.Value)
                        DoCmd.OpenQuery "qryAppendCapitalCarryOver"
                        DoCmd.Close acForm, "frmReportNavigation"
                        Beep
                        MsgBox "CarryOver Capital Projects have been added for FY " & Me.cboFiscalYear.Value & "!", vbInformation, "CarryOver Projects Added"
                    Else
                        Exit Sub
                    End If
        End If
End Select
        
End Sub

Private Sub cmdOpenSharepoint_Click()
If IsNull(Me.cboDivision) = True Then
    Beep
    MsgBox "Please select a division!", vbCritical, "Select a Divsion"
    Me.cboDivision.SetFocus
    Exit Sub
End If

Select Case Me.cboType.Value

'Open Sharepoint website
    Case "Budgeted Projects"
        CreateObject("Shell.Application").Open "http://kiicha/sites/leadership/PA/_layouts/15/WopiFrame.aspx?sourcedoc=/sites/leadership/PA/Shared%20Documents/Capital%20Budget%20by%20Division/qryBudgetDataExport_" & Me.cboDivision.Value & ".xlsx&action=default"
    Case "Approved Projects"
        CreateObject("Shell.Application").Open "http://kiicha/sites/leadership/PA/_layouts/15/WopiFrame.aspx?sourcedoc=/sites/leadership/PA/Shared%20Documents/Capital%20Projects%20by%20Division/qryProjectDataExport_" & Me.cboDivision.Value & ".xlsx&action=default"
    End Select
End Sub

Private Sub cmdViewSPdata_Click()
If IsNull(Me.cboDivision) = True Then
    Beep
    MsgBox "Please select a division!", vbCritical, "Select a Divsion"
    Me.cboDivision.SetFocus
    Me.cboDivision.Dropdown
    Exit Sub
End If

If IsNull(Me.cboType) = True Then
    Beep
    MsgBox "Please select a project type!", vbCritical, "Select a Project Type"
    Me.cboType.SetFocus
    Me.cboType.Dropdown
    Exit Sub
End If

Select Case Me.cboType.Value
    Case "Budgeted Projects"
        CurrentDb.TableDefs("excel_BudgetDataExport_" & Me.cboDivision.Value).RefreshLink
        DoCmd.Close acForm, "frmBudgetDataRead", acSaveNo
        DoCmd.OpenForm "frmBudgetDataRead", acFormDS
    
    Case "Approved Projects"
        CurrentDb.TableDefs("excel_ProjectDataExport_" & Me.cboDivision.Value).RefreshLink
        DoCmd.Close acForm, "frmProjectDataRead", acSaveNo
        DoCmd.OpenForm "frmProjectDataRead", acFormDS
        
    
End Select

End Sub

Private Sub Form_Load()
Select Case User
    Case "jgarcia"
        Me.cboImportExportFiles.RowSource = "SELECT [PECHANGA\jgarcia_dimCapExLookups].LookupValue" _
 & " FROM [PECHANGA\jgarcia_dimCapExLookups]" _
 & " WHERE ((([PECHANGA\jgarcia_dimCapExLookups].Section) = 'Import_Export_Files'))" _
 & " ORDER BY [PECHANGA\jgarcia_dimCapExLookups].SortOrder;"
    DoCmd.OpenForm "frmImportData_check", acNormal
    Case Else
        Me.cboImportExportFiles.RowSource = "SELECT [PECHANGA\jgarcia_dimCapExLookups].LookupValue" _
 & " FROM [PECHANGA\jgarcia_dimCapExLookups]" _
 & " WHERE ((([PECHANGA\jgarcia_dimCapExLookups].Section)='Import_Export_Files') AND (([PECHANGA\jgarcia_dimCapExLookups].PermissionGroup)='User'))" _
 & " ORDER BY [PECHANGA\jgarcia_dimCapExLookups].SortOrder;"
    End Select
End Sub

Private Sub txtImportDate_AfterUpdate()
If IsNull(Me.txtImportDate) = False Then
    Me.txtImportDate.Value = GetNowLast(Me.txtImportDate)
End If
End Sub


Attribute VB_Name = "CustomSubs"
Option Compare Database

Option Explicit

Public Function CreateTempVars()

TempVars("User") = Environ("USERNAME")
TempVars("Workstation") = Environ("COMPUTERNAME")
TempVars("ReportSQL") = "Start"

End Function
Public Sub CreateDivisionEmail(Division As String, FiscalYear As String, ProjectType As String)

Dim DIV_email, objOutlook As Object
Dim signature, imgEditExcel, imgInstructions, urlSharePoint, urlTableau, imgPDFfile As String
Dim CAR, Cell As Variant

'variables
Set objOutlook = CreateObject("Outlook.Application")
Set DIV_email = objOutlook.CreateItem(0)

Select Case ProjectType
    Case "Budgeted Projects"
            urlSharePoint = "http://kiicha/sites/leadership/PA/_layouts/15/WopiFrame.aspx?sourcedoc=/sites/leadership/PA" _
                            & "/Shared%20Documents/Capital%20Budget%20by%20Division/qryBudgetDataExport_" & Replace(Division, " ", "%20") & ".xlsx&action=default"
            imgEditExcel = "\\prcfile\DeptFolders\FinAnalyst\Capital\Capital Tracking Database\Images\imgEditExcel.png"
            imgInstructions = "\\prcfile\DeptFolders\FinAnalyst\Capital\Capital Tracking Database\Images\imgInstructions.png"
            imgEditExcel = "\\prcfile\DeptFolders\FinAnalyst\Capital\Capital Tracking Database\Images\imgEditExcel.png"
            urlTableau = "https://prctableau.pechanga.com/#/views/CapitalBudgetTracking_16715568671250/BudgetTrackingDB?:iid=7"
    Case "Approved Projects"
            urlSharePoint = "http://kiicha/sites/leadership/PA/_layouts/15/WopiFrame.aspx?sourcedoc=/sites/leadership/PA" _
                            & "/Shared%20Documents/Capital%20Projects%20by%20Division/qryProjectDataExport_" & Replace(Division, " ", "%20") & ".xlsx&action=default"
            imgEditExcel = "\\prcfile\DeptFolders\FinAnalyst\Capital\Capital Tracking Database\Images\imgEditExcel.png"
            imgInstructions = "\\prcfile\DeptFolders\FinAnalyst\Capital\Capital Tracking Database\Images\imgInstructionsProject.png"
            imgEditExcel = "\\prcfile\DeptFolders\FinAnalyst\Capital\Capital Tracking Database\Images\imgEditExcelProject.png"
            imgPDFfile = "\\prcfile\DeptFolders\FinAnalyst\Capital\Capital Tracking Database\Images\imgInstructionsProjectPDF.png"
            urlTableau = "https://prctableau.pechanga.com/#/views/CapitalBudgetTracking_16715568671250/BudgetTrackingDB?:iid=7"
End Select

'to test whether Outlook is open and to open it if it's not
Call TestOutlookIsOpen

With DIV_email
    .Display
End With

signature = DIV_email.HTMLbody

Select Case ProjectType
    Case "Budgeted Projects"
        With DIV_email
            .Subject = "FY " & FiscalYear & " " & Division & " Budgeted Capital Projects - Target Start Dates and Notes"
            .CC = EmailCcList(Division)
            .To = EmailToList(Division)
            .Attachments.Add imgEditExcel, 1, 0
            .Attachments.Add imgInstructions, 1, 0
            .HTMLbody = "<html><body>" _
                & "<p style=""font-family:calibri;"">Hello,<br><br>" _
                & "In an effort to bring more visibility to capital project statuses, this is your monthly email for requesting updated information concerning outstanding <strong>budgeted</strong> capital projects for <strong>" & Division & "</strong>." _
                & " The intent of this file is for you to identify target start dates for your budgeted capital projects and adding any relevant notes on the status of these projects.  By reviewing and updating this monthly, we will be able to better track the submissions of budgeted capital projects.<br><br>" _
                & " Please follow the instructions below on how to add target dates and notes for your division's projects.<br><br>" _
                & " Click on the link below to view " & Division & " capital projects that have not yet been submitted to the CSC for review." _
                & "<br><br><a href = " & urlSharePoint & ">" & Division & " - Budgeted Capital Projects </a>" _
                & "<br><br> After opening the link, proceed to edit the file using the <strong>Excel Web App</strong>." _
                & "<br><br><img src=""imgEditExcel.png"">" _
                & "<br><br> Using columns <strong>G and H</strong>, add the target start dates <strong>(date format only)</strong> as an approximation and any project-related notes." _
                & "<br><br><img src=""imgInstructions.png"">" _
                & "<br><br> Once completed, close out of the browser tab. Changes are automatically saved." _
                & "<br><br> After P&A ingests the data, it will be available to view in the Capital Budget Tracking dashboard (link below)." _
                & "<br><br><a href =" & urlTableau & ">Capital Budget Tracking </a>" _
                & "<br><br>If you have any questions or concerns, please reach out to the P&A department." _
                & "<br><br>Thank you!" _
                & "</p>" _
                & signature
            .Display
        End With
    Case "Approved Projects"
        With DIV_email
            .Subject = "FY " & FiscalYear & " " & Division & " Approved Capital Projects - Target Completion Dates and Notes"
            .CC = EmailCcList(Division)
            .To = EmailToList(Division)
            .Attachments.Add imgEditExcel, 1, 0
            .Attachments.Add imgInstructions, 1, 0
            .Attachments.Add imgPDFfile, 1, 0
            .HTMLbody = "<html><body>" _
                & "<p style=""font-family:calibri;"">Hello,<br><br>" _
                & "In an effort to bring more visibility to capital project statuses, this is your monthly email requesting updated information concerning outstanding <strong>approved</strong> capital projects for <strong>" & Division & "</strong>." _
                & " The intent of this file is for you to identify target completion dates for your open approved capital projects and adding any relevant notes on the status of the open projects.  By reviewing and updating this monthly, we will be able to better track progress and completion of open capital projects.<br><br>" _
                & " Please follow the instructions below on how to add target dates and notes for your division's projects.<br><br>" _
                & " Click on the link below to view " & Division & " capital projects that have been approved by the CSC/Board." _
                & "<br><br><a href = " & urlSharePoint & ">" & Division & " - Approved Capital Projects </a>" _
                & "<br><br> After opening the link, proceed to edit the file using the <strong>Excel Web App</strong>." _
                & "<br><br><img src=""imgEditExcelProject.png"">" _
                & "<br><br> Using columns <strong>J and K</strong>, add the target completion dates <strong>(date format only)</strong> as an approximation and any project-related notes." _
                & "<br><br><img src=""imgInstructionsProject.png"">" _
                & "<br><br> To view the approved CAR PDF files, please refer to the <strong>pdfURL</strong> column and click on the hyperlinks." _
                & "<br><br><img src=""imgInstructionsProjectPDF.png"">" _
                & "<br><br> Once completed, close out of the browser tab. Changes are automatically saved." _
                & "<br><br> If you have any questions or concerns, please reach out to the P&A department." _
                & "<br><br>Thank you!" _
                & "</p>" _
                & signature
            .Display
        End With
End Select
Set objOutlook = Nothing
Set DIV_email = Nothing

End Sub

Public Sub UpdateBudgetNotes(userName As String, NoteType As String)

DoCmd.SetWarnings False
Select Case NoteType
    Case "Budget Projects"
        'import data from SharePoint Excel file into SQL Server
        DoCmd.OpenQuery ("qryAppendBudgetNotes")
        
        'import data from query into staging table
        DoCmd.OpenQuery ("qryAppendStagingBudgetNotes")
        
        'update capital budget table from staging table
        DoCmd.OpenQuery ("qryUpdateBudgetNotes")
        
        'delete data from staging table)
        DoCmd.RunSQL ("DELETE * FROM PECHANGA\jgarcia_stagingCapitalBudgetNotes WHERE ImportedBy = '" & userName & "'")
    
    Case "Approved Projects"
        'import data from SharePoint Excel file into SQL Server
        DoCmd.OpenQuery ("qryAppendProjectNotes")
        
        'import data from query into staging table
        DoCmd.OpenQuery ("qryAppendStagingBudgetNotes")
        
        'update capital project table from staging table
        DoCmd.OpenQuery ("qryUpdateProjectNotes")
        
        'delete data from staging table)
        DoCmd.RunSQL ("DELETE * FROM PECHANGA\jgarcia_stagingCapitalBudgetNotes WHERE ImportedBy = '" & userName & "'")
End Select
DoCmd.SetWarnings True
End Sub

Public Sub ManualUpdateBudgetNotes(budgetRef As Variant _
                                 , TargetStartDate As Variant _
                                 , ChampionNotes As Variant _
                                 , ImportedBy As Variant)
DoCmd.SetWarnings False
If IsNull(budgetRef) = False And IsNull(TargetStartDate) = False And IsNull(ChampionNotes) = False Then
    If Left(budgetRef, 1) = "B" Then
        DoCmd.RunSQL ("INSERT INTO PECHANGA\jgarcia_factCapitalBudgetNotes " _
                    & " (budgetReference, TargetStartDate, ChampionNotes, [TimeStamp], ImportedBy)" _
                    & " VALUES ('" & budgetRef & "', #" & TargetStartDate & "#, '" & ChampionNotes & "'," _
                    & " #" & Now() & "#, '" & ImportedBy & "')")
    Else
        DoCmd.RunSQL ("INSERT INTO PECHANGA\jgarcia_factCapitalBudgetNotes " _
                    & " (budgetReference, TargetCompletionDate, ChampionNotes, [TimeStamp], ImportedBy)" _
                    & " VALUES ('" & budgetRef & "', #" & TargetStartDate & "#, '" & ChampionNotes & "'," _
                    & " #" & Now() & "#, '" & ImportedBy & "')")
        
    End If
End If
DoCmd.SetWarnings True
End Sub

Public Sub CreateReleasedFunds(CarNumber As String _
                             , CompletionDate As Date _
                             )
Dim AmountReleaseLookup As Double
Dim SQLInsert, SQLUpdate As String

If CurrentProject.AllForms("frmReportNavigation").IsLoaded = True Then
       
        AmountReleaseLookup = DLookup("[FundsToRelease]", "qryFundsToRelease", "CarNumber = '" & CarNumber & "'")
        
        SQLInsert = "INSERT INTO PECHANGA\jgarcia_factCapitalFundsReleased " _
                    & "(CarNumber, AmountReleased, ReleaseMonth, ReleasedBy, [timeStamp], FiscalYear) " _
                    & "VALUES ('" & CarNumber & "', " & AmountReleaseLookup & ", " _
                    & "#" & CompletionDate & "#, '" & userName() & "', #" & Now() & "#, " _
                    & FiscalYearCalc(CompletionDate) & ")"
        
        SQLUpdate = "UPDATE PECHANGA\jgarcia_factCapitalSpending " _
                  & "SET ReleasedDate = #" & CompletionDate & "#, " _
                  & "ReleasedBy = '" & userName() & "' " _
                  & "WHERE ReleasedDate IS NULL " _
                  & "AND [FinanceCodeBlock#Project] = '" & CarNumber & "'"
                    
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLInsert)
        DoCmd.RunSQL (SQLUpdate)
        'Debug.Print (SQLInsert)
        DoCmd.SetWarnings True

Else
        DoCmd.OpenForm "frmReportNavigation", acNormal, , , , acHidden
        
        AmountReleaseLookup = DLookup("[FundsToRelease]", "qryFundsToRelease", "CarNumber = '" & CarNumber & "'")
        
        SQLInsert = "INSERT INTO PECHANGA\jgarcia_factCapitalFundsReleased " _
                    & "(CarNumber, AmountReleased, ReleaseMonth, ReleasedBy, [timeStamp], FiscalYear) " _
                    & "VALUES ('" & CarNumber & "', " & AmountReleaseLookup & ", " _
                    & "#" & CompletionDate & "#, '" & userName() & "', #" & Now() & "#, " _
                    & FiscalYearCalc(CompletionDate) & ")"
        
        SQLUpdate = "UPDATE PECHANGA\jgarcia_factCapitalSpending " _
                  & "SET ReleasedDate = #" & CompletionDate & "#, " _
                  & "ReleasedBy = '" & userName() & "' " _
                  & "WHERE ReleasedDate IS NULL " _
                  & "AND [FinanceCodeBlock#Project] = '" & CarNumber & "'"
                    
        DoCmd.SetWarnings False
        DoCmd.RunSQL (SQLInsert)
        DoCmd.RunSQL (SQLUpdate)
        'Debug.Print (SQLInsert)
        DoCmd.Close acForm, "frmReportNavigation"
        DoCmd.SetWarnings True
End If

End Sub

Public Sub UpdateCARdates(Status As String, _
                            Optional BoardAgendaDate As Date = #1/1/2020#, _
                            Optional ApprovalDate As Date = #1/1/2020#, _
                            Optional CscAgendaDate As Date = #1/1/2020#, _
                            Optional CscAgendaDateNew As Date = #1/1/2020#, _
                            Optional BoardAgendaDateNew As Date = #1/1/2020# _
                            )

Dim dbsCapital As DAO.Database
Dim rstCARs As DAO.Recordset
Dim sqlSELECT, SQLUpdate, SQLDate, AgendaDateString, StatusString As String
Dim AgendaDate, NewAgendaDate, ApprovalDateVariant As Variant

Select Case Status
    Case "Submitted"
        SQLDate = "CscApprovalDate"
        AgendaDate = CscAgendaDate
        AgendaDateString = "CscAgendaDate"
        ApprovalDateVariant = ApprovalDate
        StatusString = Status
        
    Case "Pending"
        SQLDate = "BoardApprovalDate"
        AgendaDate = BoardAgendaDate
        AgendaDateString = "BoardAgendaDate"
        ApprovalDateVariant = ApprovalDate
        StatusString = Status
        
    Case "Update Board Agenda Dates"
        SQLDate = "BoardAgendaDate"
        AgendaDate = BoardAgendaDate
        NewAgendaDate = BoardAgendaDateNew
        AgendaDateString = "BoardAgendaDate"
    
    Case "Update CSC Agenda Dates"
        SQLDate = "CscAgendaDate"
        AgendaDate = CscAgendaDate
        NewAgendaDate = CscAgendaDateNew
        AgendaDateString = "CscAgendaDate"
        
End Select

Select Case Status
    Case "Submitted", "Pending"
        sqlSELECT = "SELECT CarNumber, CarAmount, CarStatus, CscApprovalDate, BoardApprovalDate, ProjectDescription FROM PECHANGA\jgarcia_factCapitalProjects" _
            & " WHERE " & AgendaDateString & " = #" & AgendaDate & "#" _
            & " AND CarStatus = '" & StatusString & "'"
    
    Case "Update Board Agenda Dates", "Update CSC Agenda Dates"
        sqlSELECT = "SELECT CarNumber, CarAmount, CarStatus, CscApprovalDate, BoardApprovalDate, ProjectDescription FROM PECHANGA\jgarcia_factCapitalProjects" _
            & " WHERE " & AgendaDateString & " = #" & AgendaDate & "#"
End Select

 Debug.Print (sqlSELECT)
Set dbsCapital = CurrentDb
Set rstCARs = dbsCapital.OpenRecordset(sqlSELECT, dbOpenDynaset)

DoCmd.SetWarnings False
Do While Not rstCARs.EOF
    Select Case Status
        Case "Submitted", "Pending"
            
            'CSC approved but not Board approved
            If rstCARs!CarAmount >= 25000 And rstCARs!CarStatus = "Submitted" Then
                SQLUpdate = "UPDATE PECHANGA\jgarcia_factCapitalProjects " _
                    & "SET CarStatus = 'Pending', " _
                    & SQLDate & " = #" & ApprovalDateVariant & "#, " _
                    & "BoardAgendaDate = #" & BoardAgendaDate & "# WHERE " _
                    & "CarNumber = '" & rstCARs!CarNumber & "'"
            Debug.Print (SQLUpdate)
                DoCmd.RunSQL (SQLUpdate)
                    
            'CSC approved
            ElseIf rstCARs!CarAmount < 25000 And rstCARs!CarStatus = "Submitted" Then
                SQLUpdate = "UPDATE PECHANGA\jgarcia_factCapitalProjects " _
                    & "SET CarStatus = 'Approved', " _
                    & "pdfURL = '" & UpdateCARpdfURLfunc(rstCARs!CarNumber, rstCARs!ProjectDescription) & "', " _
                    & SQLDate & " = #" & ApprovalDateVariant & "# WHERE" _
                    & " CarNumber = '" & rstCARs!CarNumber & "'"
                DoCmd.RunSQL (SQLUpdate)
            
            'Board approved
            ElseIf rstCARs!CarAmount >= 25000 And rstCARs!CarStatus = "Pending" Then
                SQLUpdate = "UPDATE PECHANGA\jgarcia_factCapitalProjects " _
                    & "SET CarStatus = 'Approved', " _
                    & "pdfURL = '" & UpdateCARpdfURLfunc(rstCARs!CarNumber, rstCARs!ProjectDescription) & "', " _
                    & SQLDate & " = #" & ApprovalDateVariant & "# WHERE" _
                    & " CarNumber = '" & rstCARs!CarNumber & "'"
                DoCmd.RunSQL (SQLUpdate)
            End If
        
        Case "Update Board Agenda Dates"
            SQLUpdate = "UPDATE PECHANGA\jgarcia_factCapitalProjects " _
                        & "SET " & SQLDate & " = #" & NewAgendaDate & "# " _
                        & "WHERE CarNumber = '" & rstCARs!CarNumber & "'"
                    DoCmd.RunSQL (SQLUpdate)
        
        
        Case "Update CSC Agenda Dates"
            SQLUpdate = "UPDATE PECHANGA\jgarcia_factCapitalProjects " _
                        & "SET " & SQLDate & " = #" & NewAgendaDate & "# " _
                        & "WHERE CarNumber = '" & rstCARs!CarNumber & "'"
                    DoCmd.RunSQL (SQLUpdate)
            
    End Select
    rstCARs.MoveNext
Loop
DoCmd.SetWarnings True
rstCARs.Close
Forms!frmCapitalProjects.Form.Requery
Beep
MsgBox "CAR dates updated!"

End Sub

Public Sub UpdateCARpdfURL(CarStatus, CarNumber, ProjectDescription As String, FiscalYear As Integer)

Dim SQL, pdfURL As String
DoCmd.SetWarnings False

pdfURL = "http://kiicha/sites/leadership/PA/Shared%20Documents/Approved%20Capital%20Projects/CAR%20" & ProjectDescription & ".pdf"
pdfURL = Replace(pdfURL, " ", "%20")

'SQL = "UPDATE PECHANGA\jgarcia_factCapitalProjects SET pdfURL = '" & pdfURL & "'" _
'& " WHERE CarNumber = '" & Me.CarNumber & "'"
'
'If Me.cboCARStatus.Value = "Approved" _
'    And FiscalYear >= 2023 _
'    And Int(Right(Me.CarNumber.Value, 3)) >= 21 Then
'        DoCmd.RunSQL (SQL)
'        DoCmd.RunCommand acCmdSaveRecord
'End If
'If Me.cboCARStatus.Value <> "Approved" _
'    And Me.FiscalYear.Value >= 2023 _
'    And Int(Right(Me.CarNumber.Value, 3)) >= 21 Then
'        DoCmd.RunSQL ("UPDATE PECHANGA\jgarcia_factCapitalProjects SET pdfURL = NULL" _
'    & " WHERE CarNumber = '" & Me.CarNumber & "'")
'        DoCmd.RunCommand acCmdSaveRecord
'End If
'DoCmd.SetWarnings True

End Sub

Public Function UpdateCARpdfURLfunc(CarNumber As String, ProjectDescription As String)

Dim pdfURL As String

pdfURL = "http://kiicha/sites/leadership/PA/Shared%20Documents/Approved%20Capital%20Projects/CAR%20" & CarNumber & "%20" & ProjectDescription & ".pdf"
pdfURL = Replace(pdfURL, " ", "%20")
UpdateCARpdfURLfunc = pdfURL

End Function

Public Sub ExportAllCode()
'Required References
'    Microsoft Visual Basic For Applications Extensibility

    Dim c As VBComponent
    Dim Sfx, exportPath As String
    
    exportPath = "\\prcfile\Users\jgarcia\My Documents\Code\VBA\Access VBA\Capital Management System"
    Debug.Print (CurrentProject.path)

    For Each c In Application.VBE.VBProjects(1).VBComponents
        Select Case c.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                Sfx = ".cls"
            Case vbext_ct_MSForm
                Sfx = ".frm"
            Case vbext_ct_StdModule
                Sfx = ".bas"
            Case Else
                Sfx = ""
        End Select

        If Sfx <> "" Then
            Debug.Print (c.Name)
            c.Export _
                FileName:=exportPath & "\" & _
                c.Name & Sfx
        End If
    Next c
Beep
MsgBox "VBA Code Exported!"
End Sub


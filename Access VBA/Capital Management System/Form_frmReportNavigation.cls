VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReportNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOpenActiveProjects_Click()

End Sub

Private Sub cboDivision_AfterUpdate()
'If IsNull(Me.cboDivision) = True Then
'    TempVars!ReportSQL = "Start"
'Else
'    TempVars!ReportSQL = "Division"
'End If
End Sub

Private Sub cboReports_AfterUpdate()
Me.cboDivision.Value = Null
End Sub

Private Sub cmdExcelExport_Click()
On Error Resume Next

Dim ReportExportName, ReportFileName As String

Select Case Me.cboReports.Value
    Case "Active Capital Projects"
        ReportExportName = "rptCapitalSpendingByProjectRollup"
    Case "Capital Funds Released for Completed Projects"
        ReportExportName = "rptCapitalFundsReleasedByProject"
End Select



If IsNull(Me.cboReports) = True Then
    Beep
    MsgBox "Please select a report!"
    Exit Sub
ElseIf IsNull(Me.cboFiscalYear) = True Then
    Beep
    MsgBox "Please select a fiscal year!"
    Exit Sub
ElseIf IsNull(Me.txtMonthEndDate) = True Then
    Beep
    MsgBox "Please select a month ending date!"
    Exit Sub
Else
    DoCmd.OutputTo acOutputQuery, ReportExportName, acFormatXLS, , True
End If
End Sub

Private Sub cmdOpenReport_Click()
'On Error Resume Next

Dim NewSQL, OldSQL, UpdateSQL As String

If IsNull(Me.cboReports) = True Then
    Beep
    MsgBox "Please select a report!"
    Exit Sub
ElseIf IsNull(Me.cboFiscalYear) = True Then
    Beep
    MsgBox "Please select a fiscal year!"
    Exit Sub
ElseIf IsNull(Me.txtMonthEndDate) = True Then
    Beep
    MsgBox "Please select a month ending date!"
    Exit Sub
Else
    Select Case Me.cboReports.Value
        
        Case "Capital Funds Released for Completed Projects"
            
            If Me.chkExcludeZeros = -1 Then
                CurrentDb.QueryDefs("rptCapitalFundsReleasedByProject").SQL = "SELECT rptCapitalFundsReleasedByProject_Stage1.*, qryPTReleaseNotesCombined.[Combined Notes], [rptCapitalFundsReleasedByProject_Stage1].[Remaining Balance]+Nz([qryProjectTradeInOffsets_ReleasedFunds].[SumOfAdjAmount],0) AS [Remaining Balance 2]" _
                        & " FROM (rptCapitalFundsReleasedByProject_Stage1 LEFT JOIN qryPTReleaseNotesCombined ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryPTReleaseNotesCombined.CarNumber) LEFT JOIN qryProjectTradeInOffsets_ReleasedFunds ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryProjectTradeInOffsets_ReleasedFunds.ProjectNumber" _
                        & " WHERE (((rptCapitalFundsReleasedByProject_Stage1.SumOfReleasedFunds1)<> 0));"
               Else
                CurrentDb.QueryDefs("rptCapitalFundsReleasedByProject").SQL = "SELECT rptCapitalFundsReleasedByProject_Stage1.*, qryPTReleaseNotesCombined.[Combined Notes], [rptCapitalFundsReleasedByProject_Stage1].[Remaining Balance]+Nz([qryProjectTradeInOffsets_ReleasedFunds].[SumOfAdjAmount],0) AS [Remaining Balance 2]" _
                        & " FROM (rptCapitalFundsReleasedByProject_Stage1 LEFT JOIN qryPTReleaseNotesCombined ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryPTReleaseNotesCombined.CarNumber) LEFT JOIN qryProjectTradeInOffsets_ReleasedFunds ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryProjectTradeInOffsets_ReleasedFunds.ProjectNumber"
                End If
        
        Case "Active Capital Projects By Division", "Active Capital Projects"
            
            NewSQL = "WHERE ((([PECHANGA\jgarcia_factCapitalProjects].FiscalYear)>=2023) AND (([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].CarStatus)=""Approved"") AND (([PECHANGA\jgarcia_factCapitalProjects].Division)=[Forms]![frmReportNavigation]![cboDivision])) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].CarNumber) Is Not Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].FiscalYearCarryOver)=[Forms]![frmReportNavigation]![cboFiscalYear]) AND (([PECHANGA\jgarcia_factCapitalProjects].Division)=[Forms]![frmReportNavigation]![cboDivision])) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].ActiveProjectInclusion)=True) AND (([PECHANGA\jgarcia_factCapitalProjects].Division)=[Forms]![frmReportNavigation]![cboDivision]))"
            OldSQL = "WHERE ((([PECHANGA\jgarcia_factCapitalProjects].FiscalYear)>=2023) AND (([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].CarStatus)=""Approved"")) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].CarNumber) Is Not Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].FiscalYearCarryOver)=Forms!frmReportNavigation!cboFiscalYear)) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].ActiveProjectInclusion)=True))"
            
            If IsNull(Me.cboDivision) = False Then
                'Debug.Print (CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL)
                UpdateSQL = Replace(CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL, OldSQL, NewSQL)
                Debug.Print (UpdateSQL)
                CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL = UpdateSQL
                TempVars!ReportSQL = "Division"
            Else
                OldSQL = "SELECT [PECHANGA\jgarcia_factCapitalProjects].CarNumber, [PECHANGA\jgarcia_factCapitalProjects].Division, [PECHANGA\jgarcia_factCapitalProjects].ProjectDescription, Sum([PECHANGA\jgarcia_factCapitalProjects].CarAmount) AS SumOfCarAmount, [PECHANGA\jgarcia_factCapitalProjects].TargetCompletionDate, Sum([PECHANGA\jgarcia_factCapitalProjects].CarAmount) AS [CARAmt+Adj], Nz([PECHANGA\jgarcia_factCapitalProjects].[BoardApprovalDate],[PECHANGA\jgarcia_factCapitalProjects].[CscApprovalDate]) AS ApprovedDate, DateDiff('d',Nz([PECHANGA\jgarcia_factCapitalProjects].[BoardApprovalDate],[PECHANGA\jgarcia_factCapitalProjects].[CscApprovalDate]),Now()) AS DaysOutstanding, [PECHANGA\jgarcia_factCapitalProjects].FiscalYear, Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfAP Invoice Additions]),0)/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% AP], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfAP Invoice Additions]) AS [SumOfSumOfAP Invoice Additions]," _
                    & " Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfRNI]),0)/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% RNI], Sum(qryCapitalSpendingByProjectRollupFilter.SumOfRNI) AS SumOfSumOfRNI, Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfTotal Additions]),0)/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% Total], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfTotal Additions]) AS [SumOfSumOfTotal Additions], Abs(Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfTransfers to FA]),0))/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% Transfer]," _
                    & " Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfTransfers to FA]) AS [SumOfSumOfTransfers to FA], Abs(Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfFA Adds]),0))/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% FA Adds], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfFA Adds]) AS [SumOfSumOfFA Adds], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfYear End Accrual]) AS [SumOfSumOfYear End Accrual]," _
                    & " Abs(Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfNet Additions]),0))/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]+Nz([qryProjectAdjustments].[SumOfAdjAmount],0)) AS [% Net Adds], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfNet Additions]) AS [SumOfSumOfNet Additions], [PECHANGA\jgarcia_factCapitalProjects].ChampionNotes, [CARAmt+Adj]-Nz(Sum([qryNetTotalSpendByProject].[NetTotalSpend]),0)-Nz([SumOfReleasedFunds1],0) AS [Remaining Balance]," _
                    & " ([CARAmt+Adj]-Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfNet Additions]),0))/[CARAmt+Adj] AS [% Bal], Sum(qryNetTotalSpendByProject.ReleasedFunds) AS SumOfReleasedFunds1, IIf(IsNull([PECHANGA\jgarcia_factCapitalCarryOver].[CarNumber]),""Current FY"",""Prior Year Carryover"") AS [Carry Over Class], [PECHANGA\jgarcia_factCapitalProjects].Capitalize, [PECHANGA\jgarcia_factCapitalProjects].Capitalize, [PECHANGA\jgarcia_factCapitalProjects].Notes, qryDivisionSort.SortOrder" _
                    & " FROM (qryNetTotalSpendByProject RIGHT JOIN ((([PECHANGA\jgarcia_factCapitalProjects] LEFT JOIN qryProjectAdjustments ON [PECHANGA\jgarcia_factCapitalProjects].CarNumber = qryProjectAdjustments.ProjectNumber) LEFT JOIN [PECHANGA\jgarcia_factCapitalCarryOver] ON [PECHANGA\jgarcia_factCapitalProjects].CarNumber = [PECHANGA\jgarcia_factCapitalCarryOver].CarNumber) LEFT JOIN qryCapitalSpendingByProjectRollupFilter ON [PECHANGA\jgarcia_factCapitalProjects].CarNumber = qryCapitalSpendingByProjectRollupFilter.[FinanceCodeBlock#Project]) ON qryNetTotalSpendByProject.[CarNumber] = [PECHANGA\jgarcia_factCapitalProjects].CarNumber) LEFT JOIN qryDivisionSort ON [PECHANGA\jgarcia_factCapitalProjects].Division = qryDivisionSort.LookupValue" _
                    & " WHERE ((([PECHANGA\jgarcia_factCapitalProjects].FiscalYear)>=2023) AND (([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].CarStatus)=""Approved"")) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].CarNumber) Is Not Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].FiscalYearCarryOver)=Forms!frmReportNavigation!cboFiscalYear)) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].ActiveProjectInclusion)=True))" _
                    & " GROUP BY [PECHANGA\jgarcia_factCapitalProjects].CarNumber, [PECHANGA\jgarcia_factCapitalProjects].Division, [PECHANGA\jgarcia_factCapitalProjects].ProjectDescription, [PECHANGA\jgarcia_factCapitalProjects].TargetCompletionDate, Nz([PECHANGA\jgarcia_factCapitalProjects].[BoardApprovalDate],[PECHANGA\jgarcia_factCapitalProjects].[CscApprovalDate]), DateDiff('d',Nz([PECHANGA\jgarcia_factCapitalProjects].[BoardApprovalDate],[PECHANGA\jgarcia_factCapitalProjects].[CscApprovalDate]),Now()), [PECHANGA\jgarcia_factCapitalProjects].FiscalYear, [PECHANGA\jgarcia_factCapitalProjects].ChampionNotes, IIf(IsNull([PECHANGA\jgarcia_factCapitalCarryOver].[CarNumber]),""Current FY"",""Prior Year Carryover""), [PECHANGA\jgarcia_factCapitalProjects].Capitalize, [PECHANGA\jgarcia_factCapitalProjects].Notes, qryDivisionSort.SortOrder, [PECHANGA\jgarcia_factCapitalProjects].Capitalize"
                 CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL = OldSQL
                 TempVars!ReportSQL = "Start"
            End If
          
          Case "Capital Funds Released", "Capital Funds Released Summary"
            If Me.chkExcludeZeros = -1 Then
                CurrentDb.QueryDefs("rptCapitalFundsReleasedByProject_Union").SQL = "SELECT rptCapitalFundsReleasedByProject_Stage1.Division AS DivisionName, rptCapitalFundsReleasedByProject_Stage1.CarNumber AS Reference, ""Completed and Released Projects"" AS Category, IIf([FiscalYear]<[Forms]![frmReportNavigation]![cboFiscalYear],""Prior Year Carryover"",""Current Year Projects"") AS Status, rptCapitalFundsReleasedByProject_Stage1.ProjectDescription, rptCapitalFundsReleasedByProject_Stage1.SumOfCarAmount AS ProjectAmount, rptCapitalFundsReleasedByProject_Stage1.[SumOfSumOfTotal Additions] AS CapEx, rptCapitalFundsReleasedByProject_Stage1.PriorReleasedFunds, rptCapitalFundsReleasedByProject_Stage1.SumOfReleasedFunds1 AS ReleasedAdjustment, qryPTReleaseNotesCombined.[Combined Notes] AS Notes, 2 AS SortOrder" _
                        & " FROM (rptCapitalFundsReleasedByProject_Stage1 LEFT JOIN qryPTReleaseNotesCombined ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryPTReleaseNotesCombined.CarNumber) LEFT JOIN qryProjectTradeInOffsets_ReleasedFunds ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryProjectTradeInOffsets_ReleasedFunds.ProjectNumber" _
                        & " WHERE (((rptCapitalFundsReleasedByProject_Stage1.SumOfReleasedFunds1)<>0));"
               Else
                CurrentDb.QueryDefs("rptCapitalFundsReleasedByProject_Union").SQL = "SELECT rptCapitalFundsReleasedByProject_Stage1.Division AS DivisionName, rptCapitalFundsReleasedByProject_Stage1.CarNumber AS Reference, ""Completed and Released Projects"" AS Category, IIf([FiscalYear]<[Forms]![frmReportNavigation]![cboFiscalYear],""Prior Year Carryover"",""Current Year Projects"") AS Status, rptCapitalFundsReleasedByProject_Stage1.ProjectDescription, rptCapitalFundsReleasedByProject_Stage1.SumOfCarAmount AS ProjectAmount, rptCapitalFundsReleasedByProject_Stage1.[SumOfSumOfTotal Additions] AS CapEx, rptCapitalFundsReleasedByProject_Stage1.PriorReleasedFunds, rptCapitalFundsReleasedByProject_Stage1.SumOfReleasedFunds1 AS ReleasedAdjustment, qryPTReleaseNotesCombined.[Combined Notes] AS Notes, 2 AS SortOrder" _
                        & " FROM (rptCapitalFundsReleasedByProject_Stage1 LEFT JOIN qryPTReleaseNotesCombined ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryPTReleaseNotesCombined.CarNumber) LEFT JOIN qryProjectTradeInOffsets_ReleasedFunds ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryProjectTradeInOffsets_ReleasedFunds.ProjectNumber"
                End If
            
            NewSQL = " WHERE [DivisionName] = Forms!frmReportNavigation!cboDivision "
            OldSQL = ";"

            If IsNull(Me.cboDivision) = False Then
                'Debug.Print (CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL)
                UpdateSQL = Replace(CurrentDb.QueryDefs("rptCapitalFundsReleased_Union").SQL, OldSQL, NewSQL)
                Debug.Print (UpdateSQL)
                CurrentDb.QueryDefs("rptCapitalFundsReleased_Union").SQL = UpdateSQL
                TempVars!ReportSQL = "Division"
            Else
                OldSQL = "SELECT *, ReleasedAdjustment*-1 AS [ReleasedAdjustment2]  FROM rptBudgetReference_Union;" _
                       & "UNION SELECT *,  ReleasedAdjustment*-1 AS [ReleasedAdjustment2]  FROM rptUnbudgetedCapital_Union;" _
                       & "UNION SELECT *,  ReleasedAdjustment*1 AS [ReleasedAdjustment2]  FROM rptCapitalFundsReleasedByProject_Union;"
                 CurrentDb.QueryDefs("rptCapitalFundsReleased_Union").SQL = OldSQL
                 TempVars!ReportSQL = "Start"
            End If
        
        Case "Capital Projects Tentative Schedules"
            If IsNull(Me.cboBoardAgendaDate) = True Then
                Beep
                MsgBox "Please select a board agenda date!"
                Exit Sub
            End If
        
        Case "Capital Budget by Division"
        
            If IsNull(Me.cboDivision) = False Then
                DoCmd.OpenReport "Capital Budget by Single Division", acViewReport
                Exit Sub
            End If
        End Select
    DoCmd.OpenReport Me.cboReports.Value, acViewReport
End If
End Sub

Private Sub cmdPDFexport_Click()
On Error Resume Next

Dim ReportsExportName As String

Select Case Me.cboReports.Value
        Case "Capital Funds Released for Completed Projects"
            If Me.chkExcludeZeros = -1 Then
                CurrentDb.QueryDefs("rptCapitalFundsReleasedByProject").SQL = "SELECT rptCapitalFundsReleasedByProject_Stage1.*, qryPTReleaseNotesCombined.[Combined Notes], [rptCapitalFundsReleasedByProject_Stage1].[Remaining Balance]+Nz([qryProjectTradeInOffsets_ReleasedFunds].[SumOfAdjAmount],0) AS [Remaining Balance 2]" _
                        & " FROM (rptCapitalFundsReleasedByProject_Stage1 LEFT JOIN qryPTReleaseNotesCombined ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryPTReleaseNotesCombined.CarNumber) LEFT JOIN qryProjectTradeInOffsets_ReleasedFunds ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryProjectTradeInOffsets_ReleasedFunds.ProjectNumber" _
                        & " WHERE (((rptCapitalFundsReleasedByProject_Stage1.SumOfReleasedFunds1)<> 0));"
            Else
                CurrentDb.QueryDefs("rptCapitalFundsReleasedByProject").SQL = "SELECT rptCapitalFundsReleasedByProject_Stage1.*, qryPTReleaseNotesCombined.[Combined Notes], [rptCapitalFundsReleasedByProject_Stage1].[Remaining Balance]+Nz([qryProjectTradeInOffsets_ReleasedFunds].[SumOfAdjAmount],0) AS [Remaining Balance 2]" _
                        & " FROM (rptCapitalFundsReleasedByProject_Stage1 LEFT JOIN qryPTReleaseNotesCombined ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryPTReleaseNotesCombined.CarNumber) LEFT JOIN qryProjectTradeInOffsets_ReleasedFunds ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryProjectTradeInOffsets_ReleasedFunds.ProjectNumber"
            End If
        
        Case "Active Capital Projects By Division", "Active Capital Projects"
            
            NewSQL = "WHERE ((([PECHANGA\jgarcia_factCapitalProjects].FiscalYear)>=2023) AND (([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].CarStatus)=""Approved"") AND (([PECHANGA\jgarcia_factCapitalProjects].Division)=[Forms]![frmReportNavigation]![cboDivision])) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].CarNumber) Is Not Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].FiscalYearCarryOver)=[Forms]![frmReportNavigation]![cboFiscalYear]) AND (([PECHANGA\jgarcia_factCapitalProjects].Division)=[Forms]![frmReportNavigation]![cboDivision])) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].ActiveProjectInclusion)=True) AND (([PECHANGA\jgarcia_factCapitalProjects].Division)=[Forms]![frmReportNavigation]![cboDivision]))"
            OldSQL = "WHERE ((([PECHANGA\jgarcia_factCapitalProjects].FiscalYear)>=2023) AND (([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].CarStatus)=""Approved"")) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].CarNumber) Is Not Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].FiscalYearCarryOver)=Forms!frmReportNavigation!cboFiscalYear)) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].ActiveProjectInclusion)=True))"
            
            If IsNull(Me.cboDivision) = False Then
                'Debug.Print (CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL)
                UpdateSQL = Replace(CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL, OldSQL, NewSQL)
                Debug.Print (UpdateSQL)
                CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL = UpdateSQL
                TempVars!ReportSQL = "Division"
            Else
                OldSQL = "SELECT [PECHANGA\jgarcia_factCapitalProjects].CarNumber, [PECHANGA\jgarcia_factCapitalProjects].Division, [PECHANGA\jgarcia_factCapitalProjects].ProjectDescription, Sum([PECHANGA\jgarcia_factCapitalProjects].CarAmount) AS SumOfCarAmount, [PECHANGA\jgarcia_factCapitalProjects].TargetCompletionDate, Sum([PECHANGA\jgarcia_factCapitalProjects].CarAmount) AS [CARAmt+Adj], Nz([PECHANGA\jgarcia_factCapitalProjects].[BoardApprovalDate],[PECHANGA\jgarcia_factCapitalProjects].[CscApprovalDate]) AS ApprovedDate, DateDiff('d',Nz([PECHANGA\jgarcia_factCapitalProjects].[BoardApprovalDate],[PECHANGA\jgarcia_factCapitalProjects].[CscApprovalDate]),Now()) AS DaysOutstanding, [PECHANGA\jgarcia_factCapitalProjects].FiscalYear, Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfAP Invoice Additions]),0)/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% AP], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfAP Invoice Additions]) AS [SumOfSumOfAP Invoice Additions]," _
                    & " Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfRNI]),0)/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% RNI], Sum(qryCapitalSpendingByProjectRollupFilter.SumOfRNI) AS SumOfSumOfRNI, Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfTotal Additions]),0)/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% Total], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfTotal Additions]) AS [SumOfSumOfTotal Additions], Abs(Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfTransfers to FA]),0))/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% Transfer]," _
                    & " Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfTransfers to FA]) AS [SumOfSumOfTransfers to FA], Abs(Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfFA Adds]),0))/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% FA Adds], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfFA Adds]) AS [SumOfSumOfFA Adds], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfYear End Accrual]) AS [SumOfSumOfYear End Accrual]," _
                    & " Abs(Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfNet Additions]),0))/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]+Nz([qryProjectAdjustments].[SumOfAdjAmount],0)) AS [% Net Adds], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfNet Additions]) AS [SumOfSumOfNet Additions], [PECHANGA\jgarcia_factCapitalProjects].ChampionNotes, [CARAmt+Adj]-Nz(Sum([qryNetTotalSpendByProject].[NetTotalSpend]),0)-Nz([SumOfReleasedFunds1],0) AS [Remaining Balance]," _
                    & " ([CARAmt+Adj]-Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfNet Additions]),0))/[CARAmt+Adj] AS [% Bal], Sum(qryNetTotalSpendByProject.ReleasedFunds) AS SumOfReleasedFunds1, IIf(IsNull([PECHANGA\jgarcia_factCapitalCarryOver].[CarNumber]),""Current FY"",""Prior Year Carryover"") AS [Carry Over Class], [PECHANGA\jgarcia_factCapitalProjects].Capitalize, [PECHANGA\jgarcia_factCapitalProjects].Capitalize, [PECHANGA\jgarcia_factCapitalProjects].Notes, qryDivisionSort.SortOrder" _
                    & " FROM (qryNetTotalSpendByProject RIGHT JOIN ((([PECHANGA\jgarcia_factCapitalProjects] LEFT JOIN qryProjectAdjustments ON [PECHANGA\jgarcia_factCapitalProjects].CarNumber = qryProjectAdjustments.ProjectNumber) LEFT JOIN [PECHANGA\jgarcia_factCapitalCarryOver] ON [PECHANGA\jgarcia_factCapitalProjects].CarNumber = [PECHANGA\jgarcia_factCapitalCarryOver].CarNumber) LEFT JOIN qryCapitalSpendingByProjectRollupFilter ON [PECHANGA\jgarcia_factCapitalProjects].CarNumber = qryCapitalSpendingByProjectRollupFilter.[FinanceCodeBlock#Project]) ON qryNetTotalSpendByProject.[CarNumber] = [PECHANGA\jgarcia_factCapitalProjects].CarNumber) LEFT JOIN qryDivisionSort ON [PECHANGA\jgarcia_factCapitalProjects].Division = qryDivisionSort.LookupValue" _
                    & " WHERE ((([PECHANGA\jgarcia_factCapitalProjects].FiscalYear)>=2023) AND (([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].CarStatus)=""Approved"")) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].CarNumber) Is Not Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].FiscalYearCarryOver)=Forms!frmReportNavigation!cboFiscalYear)) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].ActiveProjectInclusion)=True))" _
                    & " GROUP BY [PECHANGA\jgarcia_factCapitalProjects].CarNumber, [PECHANGA\jgarcia_factCapitalProjects].Division, [PECHANGA\jgarcia_factCapitalProjects].ProjectDescription, [PECHANGA\jgarcia_factCapitalProjects].TargetCompletionDate, Nz([PECHANGA\jgarcia_factCapitalProjects].[BoardApprovalDate],[PECHANGA\jgarcia_factCapitalProjects].[CscApprovalDate]), DateDiff('d',Nz([PECHANGA\jgarcia_factCapitalProjects].[BoardApprovalDate],[PECHANGA\jgarcia_factCapitalProjects].[CscApprovalDate]),Now()), [PECHANGA\jgarcia_factCapitalProjects].FiscalYear, [PECHANGA\jgarcia_factCapitalProjects].ChampionNotes, IIf(IsNull([PECHANGA\jgarcia_factCapitalCarryOver].[CarNumber]),""Current FY"",""Prior Year Carryover""), [PECHANGA\jgarcia_factCapitalProjects].Capitalize, [PECHANGA\jgarcia_factCapitalProjects].Notes, qryDivisionSort.SortOrder, [PECHANGA\jgarcia_factCapitalProjects].Capitalize"
                 CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL = OldSQL
                 TempVars!ReportSQL = "Start"
            End If
        
        Case "Capital Funds Released", "Capital Funds Released Summary"
            If Me.chkExcludeZeros = -1 Then
                CurrentDb.QueryDefs("rptCapitalFundsReleasedByProject_Union").SQL = "SELECT rptCapitalFundsReleasedByProject_Stage1.Division AS DivisionName, rptCapitalFundsReleasedByProject_Stage1.CarNumber AS Reference, ""Completed and Released Projects"" AS Category, IIf([FiscalYear]<[Forms]![frmReportNavigation]![cboFiscalYear],""Prior Year Carryover"",""Current Year Projects"") AS Status, rptCapitalFundsReleasedByProject_Stage1.ProjectDescription, rptCapitalFundsReleasedByProject_Stage1.SumOfCarAmount AS ProjectAmount, rptCapitalFundsReleasedByProject_Stage1.[SumOfSumOfTotal Additions] AS CapEx, rptCapitalFundsReleasedByProject_Stage1.PriorReleasedFunds, rptCapitalFundsReleasedByProject_Stage1.SumOfReleasedFunds1 AS ReleasedAdjustment, qryPTReleaseNotesCombined.[Combined Notes] AS Notes, 2 AS SortOrder" _
                        & " FROM (rptCapitalFundsReleasedByProject_Stage1 LEFT JOIN qryPTReleaseNotesCombined ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryPTReleaseNotesCombined.CarNumber) LEFT JOIN qryProjectTradeInOffsets_ReleasedFunds ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryProjectTradeInOffsets_ReleasedFunds.ProjectNumber" _
                        & " WHERE (((rptCapitalFundsReleasedByProject_Stage1.SumOfReleasedFunds1)<>0));"
            Else
                CurrentDb.QueryDefs("rptCapitalFundsReleasedByProject_Union").SQL = "SELECT rptCapitalFundsReleasedByProject_Stage1.Division AS DivisionName, rptCapitalFundsReleasedByProject_Stage1.CarNumber AS Reference, ""Completed and Released Projects"" AS Category, IIf([FiscalYear]<[Forms]![frmReportNavigation]![cboFiscalYear],""Prior Year Carryover"",""Current Year Projects"") AS Status, rptCapitalFundsReleasedByProject_Stage1.ProjectDescription, rptCapitalFundsReleasedByProject_Stage1.SumOfCarAmount AS ProjectAmount, rptCapitalFundsReleasedByProject_Stage1.[SumOfSumOfTotal Additions] AS CapEx, rptCapitalFundsReleasedByProject_Stage1.PriorReleasedFunds, rptCapitalFundsReleasedByProject_Stage1.SumOfReleasedFunds1 AS ReleasedAdjustment, qryPTReleaseNotesCombined.[Combined Notes] AS Notes, 2 AS SortOrder" _
                        & " FROM (rptCapitalFundsReleasedByProject_Stage1 LEFT JOIN qryPTReleaseNotesCombined ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryPTReleaseNotesCombined.CarNumber) LEFT JOIN qryProjectTradeInOffsets_ReleasedFunds ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryProjectTradeInOffsets_ReleasedFunds.ProjectNumber"
            End If
        
        NewSQL = " WHERE [DivisionName] = Forms!frmReportNavigation!cboDivision "
            OldSQL = ";"

            If IsNull(Me.cboDivision) = False Then
                'Debug.Print (CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL)
                UpdateSQL = Replace(CurrentDb.QueryDefs("rptCapitalFundsReleased_Union").SQL, OldSQL, NewSQL)
                Debug.Print (UpdateSQL)
                CurrentDb.QueryDefs("rptCapitalFundsReleased_Union").SQL = UpdateSQL
                TempVars!ReportSQL = "Division"
            Else
                OldSQL = "SELECT *, ReleasedAdjustment*-1 AS [ReleasedAdjustment2]  FROM rptBudgetReference_Union;" _
                       & "UNION SELECT *,  ReleasedAdjustment*-1 AS [ReleasedAdjustment2]  FROM rptUnbudgetedCapital_Union;" _
                       & "UNION SELECT *,  ReleasedAdjustment*1 AS [ReleasedAdjustment2]  FROM rptCapitalFundsReleasedByProject_Union;"
                 CurrentDb.QueryDefs("rptCapitalFundsReleased_Union").SQL = OldSQL
                 TempVars!ReportSQL = "Start"
            End If
        End Select

If IsNull(Me.cboReports) = True Then
    Beep
    MsgBox "Please select a report!"
    Exit Sub
ElseIf IsNull(Me.cboFiscalYear) = True Then
    Beep
    MsgBox "Please select a fiscal year!"
    Exit Sub
ElseIf IsNull(Me.txtMonthEndDate) = True Then
    Beep
    MsgBox "Please select a month ending date!"
    Exit Sub
Else
    Select Case Me.cboReports.Value
        Case "Capital Budget by Division"
            If IsNull(Me.cboDivision) = False Then
                DoCmd.OpenReport "Capital Budget by Single Division", acViewReport
                DoCmd.OutputTo acOutputReport, "Capital Budget by Single Division", acFormatPDF, , True
                DoCmd.Close acReport, "Capital Budget by Single Division"
                Exit Sub
            End If
    End Select

    DoCmd.OpenReport (Me.cboReports.Value), acViewReport
    DoCmd.OutputTo acOutputReport, Me.cboReports.Value, acFormatPDF, , True
    DoCmd.Close acReport, Me.cboReports.Value
End If

End Sub

Private Sub cmdRefreshReport_Click()
'On Error Resume Next

Dim rpt As Report
Dim NewSQL, OldSQL, UpdateSQL As String

For Each rpt In Reports
    'Debug.Print rpt.Name
    Select Case rpt.Name
        Case "Capital Funds Released for Completed Projects"
            If Me.chkExcludeZeros = -1 Then
                CurrentDb.QueryDefs("rptCapitalFundsReleasedByProject").SQL = "SELECT rptCapitalFundsReleasedByProject_Stage1.*, qryPTReleaseNotesCombined.[Combined Notes], [rptCapitalFundsReleasedByProject_Stage1].[Remaining Balance]+Nz([qryProjectTradeInOffsets_ReleasedFunds].[SumOfAdjAmount],0) AS [Remaining Balance 2]" _
                        & " FROM (rptCapitalFundsReleasedByProject_Stage1 LEFT JOIN qryPTReleaseNotesCombined ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryPTReleaseNotesCombined.CarNumber) LEFT JOIN qryProjectTradeInOffsets_ReleasedFunds ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryProjectTradeInOffsets_ReleasedFunds.ProjectNumber" _
                        & " WHERE (((rptCapitalFundsReleasedByProject_Stage1.SumOfReleasedFunds1)<> 0));"
               Else
                CurrentDb.QueryDefs("rptCapitalFundsReleasedByProject").SQL = "SELECT rptCapitalFundsReleasedByProject_Stage1.*, qryPTReleaseNotesCombined.[Combined Notes], [rptCapitalFundsReleasedByProject_Stage1].[Remaining Balance]+Nz([qryProjectTradeInOffsets_ReleasedFunds].[SumOfAdjAmount],0) AS [Remaining Balance 2]" _
                        & " FROM (rptCapitalFundsReleasedByProject_Stage1 LEFT JOIN qryPTReleaseNotesCombined ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryPTReleaseNotesCombined.CarNumber) LEFT JOIN qryProjectTradeInOffsets_ReleasedFunds ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryProjectTradeInOffsets_ReleasedFunds.ProjectNumber"
                End If
            Reports(rpt.Name).Requery
        
         Case "Active Capital Projects By Division", "Active Capital Projects"
            
            If IsNull(Me.cboDivision) = False And TempVars!ReportSQL = "Start" Then
            
                NewSQL = "WHERE ((([PECHANGA\jgarcia_factCapitalProjects].FiscalYear)>=2023) AND (([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].CarStatus)=""Approved"") AND (([PECHANGA\jgarcia_factCapitalProjects].Division)=[Forms]![frmReportNavigation]![cboDivision])) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].CarNumber) Is Not Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].FiscalYearCarryOver)=[Forms]![frmReportNavigation]![cboFiscalYear]) AND (([PECHANGA\jgarcia_factCapitalProjects].Division)=[Forms]![frmReportNavigation]![cboDivision])) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].ActiveProjectInclusion)=True) AND (([PECHANGA\jgarcia_factCapitalProjects].Division)=[Forms]![frmReportNavigation]![cboDivision]))"
                OldSQL = "WHERE ((([PECHANGA\jgarcia_factCapitalProjects].FiscalYear)>=2023) AND (([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].CarStatus)=""Approved"")) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].CarNumber) Is Not Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].FiscalYearCarryOver)=Forms!frmReportNavigation!cboFiscalYear)) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].ActiveProjectInclusion)=True))"
                UpdateSQL = Replace(CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL, OldSQL, NewSQL)
                'Debug.Print (UpdateSQL)
                CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL = UpdateSQL
                DoCmd.Close acReport, rpt.Name
                DoCmd.OpenReport "Active Capital Projects By Division", acViewReport
                'Reports(rpt.Name).Requery
                TempVars!ReportSQL = "Division"
                
            ElseIf (IsNull(Me.cboDivision) = True Or Me.cboDivision = "") And TempVars!ReportSQL = "Division" Then
            
                OldSQL = "SELECT [PECHANGA\jgarcia_factCapitalProjects].CarNumber, [PECHANGA\jgarcia_factCapitalProjects].Division, [PECHANGA\jgarcia_factCapitalProjects].ProjectDescription, Sum([PECHANGA\jgarcia_factCapitalProjects].CarAmount) AS SumOfCarAmount, [PECHANGA\jgarcia_factCapitalProjects].TargetCompletionDate, Sum([PECHANGA\jgarcia_factCapitalProjects].CarAmount) AS [CARAmt+Adj], Nz([PECHANGA\jgarcia_factCapitalProjects].[BoardApprovalDate],[PECHANGA\jgarcia_factCapitalProjects].[CscApprovalDate]) AS ApprovedDate, DateDiff('d',Nz([PECHANGA\jgarcia_factCapitalProjects].[BoardApprovalDate],[PECHANGA\jgarcia_factCapitalProjects].[CscApprovalDate]),Now()) AS DaysOutstanding, [PECHANGA\jgarcia_factCapitalProjects].FiscalYear, Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfAP Invoice Additions]),0)/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% AP], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfAP Invoice Additions]) AS [SumOfSumOfAP Invoice Additions]," _
                    & " Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfRNI]),0)/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% RNI], Sum(qryCapitalSpendingByProjectRollupFilter.SumOfRNI) AS SumOfSumOfRNI, Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfTotal Additions]),0)/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% Total], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfTotal Additions]) AS [SumOfSumOfTotal Additions], Abs(Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfTransfers to FA]),0))/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% Transfer]," _
                    & " Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfTransfers to FA]) AS [SumOfSumOfTransfers to FA], Abs(Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfFA Adds]),0))/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]) AS [% FA Adds], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfFA Adds]) AS [SumOfSumOfFA Adds], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfYear End Accrual]) AS [SumOfSumOfYear End Accrual]," _
                    & " Abs(Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfNet Additions]),0))/Sum([PECHANGA\jgarcia_factCapitalProjects].[CarAmount]+Nz([qryProjectAdjustments].[SumOfAdjAmount],0)) AS [% Net Adds], Sum(qryCapitalSpendingByProjectRollupFilter.[SumOfNet Additions]) AS [SumOfSumOfNet Additions], [PECHANGA\jgarcia_factCapitalProjects].ChampionNotes, [CARAmt+Adj]-Nz(Sum([qryNetTotalSpendByProject].[NetTotalSpend]),0)-Nz([SumOfReleasedFunds1],0) AS [Remaining Balance]," _
                    & " ([CARAmt+Adj]-Nz(Sum([qryCapitalSpendingByProjectRollupFilter].[SumOfNet Additions]),0))/[CARAmt+Adj] AS [% Bal], Sum(qryNetTotalSpendByProject.ReleasedFunds) AS SumOfReleasedFunds1, IIf(IsNull([PECHANGA\jgarcia_factCapitalCarryOver].[CarNumber]),""Current FY"",""Prior Year Carryover"") AS [Carry Over Class], [PECHANGA\jgarcia_factCapitalProjects].Capitalize, [PECHANGA\jgarcia_factCapitalProjects].Capitalize, [PECHANGA\jgarcia_factCapitalProjects].Notes, qryDivisionSort.SortOrder" _
                    & " FROM (qryNetTotalSpendByProject RIGHT JOIN ((([PECHANGA\jgarcia_factCapitalProjects] LEFT JOIN qryProjectAdjustments ON [PECHANGA\jgarcia_factCapitalProjects].CarNumber = qryProjectAdjustments.ProjectNumber) LEFT JOIN [PECHANGA\jgarcia_factCapitalCarryOver] ON [PECHANGA\jgarcia_factCapitalProjects].CarNumber = [PECHANGA\jgarcia_factCapitalCarryOver].CarNumber) LEFT JOIN qryCapitalSpendingByProjectRollupFilter ON [PECHANGA\jgarcia_factCapitalProjects].CarNumber = qryCapitalSpendingByProjectRollupFilter.[FinanceCodeBlock#Project]) ON qryNetTotalSpendByProject.[CarNumber] = [PECHANGA\jgarcia_factCapitalProjects].CarNumber) LEFT JOIN qryDivisionSort ON [PECHANGA\jgarcia_factCapitalProjects].Division = qryDivisionSort.LookupValue" _
                    & " WHERE ((([PECHANGA\jgarcia_factCapitalProjects].FiscalYear)>=2023) AND (([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].CarStatus)=""Approved"")) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].CarNumber) Is Not Null) AND (([PECHANGA\jgarcia_factCapitalCarryOver].FiscalYearCarryOver)=Forms!frmReportNavigation!cboFiscalYear)) OR ((([PECHANGA\jgarcia_factCapitalProjects].CompletionDate) Is Null) AND (([PECHANGA\jgarcia_factCapitalProjects].ActiveProjectInclusion)=True))" _
                    & " GROUP BY [PECHANGA\jgarcia_factCapitalProjects].CarNumber, [PECHANGA\jgarcia_factCapitalProjects].Division, [PECHANGA\jgarcia_factCapitalProjects].ProjectDescription, [PECHANGA\jgarcia_factCapitalProjects].TargetCompletionDate, Nz([PECHANGA\jgarcia_factCapitalProjects].[BoardApprovalDate],[PECHANGA\jgarcia_factCapitalProjects].[CscApprovalDate]), DateDiff('d',Nz([PECHANGA\jgarcia_factCapitalProjects].[BoardApprovalDate],[PECHANGA\jgarcia_factCapitalProjects].[CscApprovalDate]),Now()), [PECHANGA\jgarcia_factCapitalProjects].FiscalYear, [PECHANGA\jgarcia_factCapitalProjects].ChampionNotes, IIf(IsNull([PECHANGA\jgarcia_factCapitalCarryOver].[CarNumber]),""Current FY"",""Prior Year Carryover""), [PECHANGA\jgarcia_factCapitalProjects].Capitalize, [PECHANGA\jgarcia_factCapitalProjects].Notes, qryDivisionSort.SortOrder, [PECHANGA\jgarcia_factCapitalProjects].Capitalize"
                 CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL = OldSQL
                 DoCmd.Close acReport, rpt.Name
                 DoCmd.OpenReport "Active Capital Projects By Division", acViewReport
'                 Debug.Print (CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL)
                 'Reports(rpt.Name).Requery
                 TempVars!ReportSQL = "Start"
            
            Else
                Reports(rpt.Name).Requery
            End If
        
        Case "Capital Funds Released", "Capital Funds Released Summary"
        
            If Me.chkExcludeZeros = -1 Then
                CurrentDb.QueryDefs("rptCapitalFundsReleasedByProject_Union").SQL = "SELECT rptCapitalFundsReleasedByProject_Stage1.Division AS DivisionName, rptCapitalFundsReleasedByProject_Stage1.CarNumber AS Reference, ""Completed and Released Projects"" AS Category, IIf([FiscalYear]<[Forms]![frmReportNavigation]![cboFiscalYear],""Prior Year Carryover"",""Current Year Projects"") AS Status, rptCapitalFundsReleasedByProject_Stage1.ProjectDescription, rptCapitalFundsReleasedByProject_Stage1.SumOfCarAmount AS ProjectAmount, rptCapitalFundsReleasedByProject_Stage1.[SumOfSumOfTotal Additions] AS CapEx, rptCapitalFundsReleasedByProject_Stage1.PriorReleasedFunds, rptCapitalFundsReleasedByProject_Stage1.SumOfReleasedFunds1 AS ReleasedAdjustment, qryPTReleaseNotesCombined.[Combined Notes] AS Notes, 2 AS SortOrder" _
                        & " FROM (rptCapitalFundsReleasedByProject_Stage1 LEFT JOIN qryPTReleaseNotesCombined ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryPTReleaseNotesCombined.CarNumber) LEFT JOIN qryProjectTradeInOffsets_ReleasedFunds ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryProjectTradeInOffsets_ReleasedFunds.ProjectNumber" _
                        & " WHERE (((rptCapitalFundsReleasedByProject_Stage1.SumOfReleasedFunds1)<>0));"
               Else
                CurrentDb.QueryDefs("rptCapitalFundsReleasedByProject_Union").SQL = "SELECT rptCapitalFundsReleasedByProject_Stage1.Division AS DivisionName, rptCapitalFundsReleasedByProject_Stage1.CarNumber AS Reference, ""Completed and Released Projects"" AS Category, IIf([FiscalYear]<[Forms]![frmReportNavigation]![cboFiscalYear],""Prior Year Carryover"",""Current Year Projects"") AS Status, rptCapitalFundsReleasedByProject_Stage1.ProjectDescription, rptCapitalFundsReleasedByProject_Stage1.SumOfCarAmount AS ProjectAmount, rptCapitalFundsReleasedByProject_Stage1.[SumOfSumOfTotal Additions] AS CapEx, rptCapitalFundsReleasedByProject_Stage1.PriorReleasedFunds, rptCapitalFundsReleasedByProject_Stage1.SumOfReleasedFunds1 AS ReleasedAdjustment, qryPTReleaseNotesCombined.[Combined Notes] AS Notes, 2 AS SortOrder" _
                        & " FROM (rptCapitalFundsReleasedByProject_Stage1 LEFT JOIN qryPTReleaseNotesCombined ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryPTReleaseNotesCombined.CarNumber) LEFT JOIN qryProjectTradeInOffsets_ReleasedFunds ON rptCapitalFundsReleasedByProject_Stage1.CarNumber = qryProjectTradeInOffsets_ReleasedFunds.ProjectNumber"
                End If
             
             NewSQL = " WHERE [DivisionName] = Forms!frmReportNavigation!cboDivision "
             OldSQL = ";"

            If IsNull(Me.cboDivision) = False And TempVars!ReportSQL = "Start" Then
                'Debug.Print (CurrentDb.QueryDefs("rptCapitalSpendingByProjectRollup").SQL)
                UpdateSQL = Replace(CurrentDb.QueryDefs("rptCapitalFundsReleased_Union").SQL, OldSQL, NewSQL)
                Debug.Print (UpdateSQL)
                CurrentDb.QueryDefs("rptCapitalFundsReleased_Union").SQL = UpdateSQL
                TempVars!ReportSQL = "Division"
            Else
                OldSQL = "SELECT *, ReleasedAdjustment*-1 AS [ReleasedAdjustment2]  FROM rptBudgetReference_Union;" _
                       & "UNION SELECT *,  ReleasedAdjustment*-1 AS [ReleasedAdjustment2]  FROM rptUnbudgetedCapital_Union;" _
                       & "UNION SELECT *,  ReleasedAdjustment*1 AS [ReleasedAdjustment2]  FROM rptCapitalFundsReleasedByProject_Union;"
                 CurrentDb.QueryDefs("rptCapitalFundsReleased_Union").SQL = OldSQL
                 TempVars!ReportSQL = "Start"
            End If
            
            Reports(rpt.Name).Requery
            
            Case Else
                Reports(rpt.Name).Requery
            
    End Select
Next rpt

End Sub

Private Sub Form_Load()
'Me.txtMonthEndDate.Value = DMax("[MonthEndDate]", "[PECHANGA\jgarcia_factCapitalSpending]")
'Me.cboFiscalYear.Value = DMax("[FiscalYearCarryOver]", "[PECHANGA\jgarcia_factCapitalCarryOver]")
End Sub

Private Sub txtMonthEndDate_AfterUpdate()
On Error Resume Next
If Me.cboReports <> "CSC Agenda" Then
    Me.txtMonthEndDate.Value = GetNowLast(Me.txtMonthEndDate.Value)
End If
End Sub

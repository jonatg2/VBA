VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCapitalProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdExportTentativeSchedule_Click()
On Error Resume Next

Dim TentativeSchedule As String

TentativeSchedule = "Capital Projects Single Tentative Schedule"

DoCmd.OpenReport (TentativeSchedule), acViewReport
DoCmd.OutputTo acOutputReport, TentativeSchedule, acFormatPDF, , True
DoCmd.Close acReport, TentativeSchedule

End Sub

Private Sub cmdPDFexport_Click()
On Error Resume Next
    Dim rExport As String
    rExport = "Project Spending Master"
    DoCmd.OpenReport (rExport), acViewReport
    DoCmd.OutputTo acOutputReport, rExport, acFormatPDF, , True
    DoCmd.Close acReport, rExport
End Sub

Private Sub txtCapitalizationDate_AfterUpdate()
If IsNull(Me.txtCapitalizationDate) = False Then
    Me.txtCapitalizationDate.Value = GetNowLast(Me.txtCapitalizationDate)
End If
End Sub

Private Sub CarAmount_AfterUpdate()

Dim curRecord As Variant
curRecord = Me.CarNumber

If IsNull(Me.cboBudgetReference) = False Then

    'update Remaining Balance
    If IsNull(DLookup("[Remaining Balance]", "qryPTBudgetRemainingBalances", "[CarNumber] = '" & Me.CarNumber & "'")) = False Then
        Me.txtRemainingBal.Value = Format(DLookup("[Remaining Balance]", "qryPTBudgetRemainingBalances", "[CarNumber] = '" & Me.CarNumber & "'"), "Currency", "$0.00")
    Else
        Me.txtRemainingBal.Value = 0
    End If
    
    'update Current Balance
    If IsNull(DLookup("[Current Balance]", "qryPTBudgetRemainingBalances", "[CarNumber] = '" & Me.CarNumber & "'")) = False Then
        Me.txtCurrentBal.Value = Format(DLookup("[Current Balance]", "qryPTBudgetRemainingBalances", "[CarNumber] = '" & Me.CarNumber & "'"), "Currency", "$0.00")
    Else
        Me.txtCurrentBal.Value = 0
    End If
    
    DoCmd.Requery
    DoCmd.FindRecord curRecord, , True, , True

End If
End Sub

Private Sub CarNumber_AfterUpdate()

Me.FiscalYear.Value = Left(Year(Now()), 2) & Left(Me.CarNumber.Value, 2)
DoCmd.RunCommand acCmdSaveRecord

End Sub

Private Sub CarNumber_LostFocus()
Me.cboBudgetReference.Requery
End Sub

Private Sub cboBudgetReference_AfterUpdate()
Dim curRecord As Variant
curRecord = Me.CarNumber

'update Remaining Balance
If IsNull(DLookup("[Remaining Balance]", "qryPTBudgetRemainingBalances", "[CarNumber] = '" & Me.CarNumber & "'")) = False Then
    Me.txtRemainingBal.Value = Format(DLookup("[Remaining Balance]", "qryPTBudgetRemainingBalances", "[CarNumber] = '" & Me.CarNumber & "'"), "Currency", "$0.00")
Else
    Me.txtRemainingBal.Value = 0
End If

'update Current Balance
If IsNull(DLookup("[Current Balance]", "qryPTBudgetRemainingBalances", "[CarNumber] = '" & Me.CarNumber & "'")) = False Then
    Me.txtCurrentBal.Value = Format(DLookup("[Current Balance]", "qryPTBudgetRemainingBalances", "[CarNumber] = '" & Me.CarNumber & "'"), "Currency", "$0.00")
Else
    Me.txtCurrentBal.Value = 0
End If

DoCmd.Requery
DoCmd.FindRecord curRecord, , True, , True

End Sub

Private Sub cboBudgetReference_GotFocus()
Me.cboBudgetReference.Requery
End Sub

Private Sub cboCarStatus_AfterUpdate()
Dim SQL, pdfURL As String
DoCmd.SetWarnings False

pdfURL = "http://kiicha/sites/leadership/PA/Shared%20Documents/Approved%20Capital%20Projects/CAR%20" & Me.txtCARfilename & ".pdf"
pdfURL = Replace(pdfURL, " ", "%20")

SQL = "UPDATE PECHANGA\jgarcia_factCapitalProjects SET pdfURL = '" & pdfURL & "'" _
& " WHERE CarNumber = '" & Me.CarNumber & "'"

If Me.cboCARStatus.Value = "Approved" _
    And Me.FiscalYear.Value >= 2023 Then
        Me.txtCARfilename.Requery
        DoCmd.RunSQL (SQL)
        DoCmd.RunCommand acCmdSaveRecord
End If
If Me.cboCARStatus.Value <> "Approved" _
    And Me.FiscalYear.Value >= 2023 Then
        DoCmd.RunSQL ("UPDATE PECHANGA\jgarcia_factCapitalProjects SET pdfURL = NULL" _
    & " WHERE CarNumber = '" & Me.CarNumber & "'")
        DoCmd.RunCommand acCmdSaveRecord
End If
DoCmd.SetWarnings True
End Sub

Private Sub cboDepartment_GotFocus()
Me.cboDepartment.Requery
End Sub

Private Sub cboDivision_LostFocus()
Me.cboDepartment.Requery
Me.cboVendor.Requery
End Sub

Private Sub cboFilter1_AfterUpdate()
Select Case cboFilter1.Value
    Case "CSCAgendaDate"
        cboFilterInput1.RowSource = "SELECT DISTINCT [PECHANGA\jgarcia_factCapitalProjects].CscAgendaDate" _
        & " FROM [PECHANGA\jgarcia_factCapitalProjects]" _
        & " WHERE ((([PECHANGA\jgarcia_factCapitalProjects].CscAgendaDate) Is Not Null))" _
        & " ORDER BY [PECHANGA\jgarcia_factCapitalProjects].CscAgendaDate DESC;"
        Me.cboFilterInput1.Requery
    
    Case "BoardAgendaDate"
        cboFilterInput1.RowSource = "SELECT DISTINCT [PECHANGA\jgarcia_factCapitalProjects].BoardAgendaDate" _
        & " FROM [PECHANGA\jgarcia_factCapitalProjects]" _
        & " WHERE ((([PECHANGA\jgarcia_factCapitalProjects].BoardAgendaDate) Is Not Null))" _
        & " ORDER BY [PECHANGA\jgarcia_factCapitalProjects].BoardAgendaDate DESC;"
        Me.cboFilterInput1.Requery
    
    Case Else
        cboFilterInput1.RowSource = "SELECT [PECHANGA\jgarcia_dimCapExLookups].LookupValue" _
        & " FROM [PECHANGA\jgarcia_dimCapExLookups]" _
        & " WHERE ((([PECHANGA\jgarcia_dimCapExLookups].Section)=[Forms]![frmCapitalProjects]![cboFilter1]));"
        Me.cboFilterInput1.Requery

    End Select
End Sub
Private Sub cboFilter2_AfterUpdate()
Select Case cboFilter2.Value
    Case "CSCAgendaDate"
        cboFilterInput2.RowSource = "SELECT DISTINCT [PECHANGA\jgarcia_factCapitalProjects].CscAgendaDate" _
        & " FROM [PECHANGA\jgarcia_factCapitalProjects]" _
        & " WHERE ((([PECHANGA\jgarcia_factCapitalProjects].CscAgendaDate) Is Not Null))" _
        & " ORDER BY [PECHANGA\jgarcia_factCapitalProjects].CscAgendaDate DESC;"
        Me.cboFilterInput2.Requery

    Case "BoardAgendaDate"
        cboFilterInput2.RowSource = "SELECT DISTINCT [PECHANGA\jgarcia_factCapitalProjects].BoardAgendaDate" _
        & " FROM [PECHANGA\jgarcia_factCapitalProjects]" _
        & " WHERE ((([PECHANGA\jgarcia_factCapitalProjects].BoardAgendaDate) Is Not Null))" _
        & " ORDER BY [PECHANGA\jgarcia_factCapitalProjects].BoardAgendaDate DESC;"
        Me.cboFilterInput1.Requery
    
    Case Else
        cboFilterInput2.RowSource = "SELECT [PECHANGA\jgarcia_dimCapExLookups].LookupValue" _
        & " FROM [PECHANGA\jgarcia_dimCapExLookups]" _
        & " WHERE ((([PECHANGA\jgarcia_dimCapExLookups].Section)=[Forms]![frmCapitalProjects]![cboFilter2]));"
        Me.cboFilterInput2.Requery

    End Select
End Sub
Private Sub cboFilter3_AfterUpdate()
Select Case cboFilter3.Value
    Case "CSCAgendaDate"
        cboFilterInput3.RowSource = "SELECT DISTINCT [PECHANGA\jgarcia_factCapitalProjects].CscAgendaDate" _
        & " FROM [PECHANGA\jgarcia_factCapitalProjects]" _
        & " WHERE ((([PECHANGA\jgarcia_factCapitalProjects].CscAgendaDate) Is Not Null))" _
        & " ORDER BY [PECHANGA\jgarcia_factCapitalProjects].CscAgendaDate DESC;"
        Me.cboFilterInput3.Requery

    Case "BoardAgendaDate"
        cboFilterInput3.RowSource = "SELECT DISTINCT [PECHANGA\jgarcia_factCapitalProjects].BoardAgendaDate" _
        & " FROM [PECHANGA\jgarcia_factCapitalProjects]" _
        & " WHERE ((([PECHANGA\jgarcia_factCapitalProjects].BoardAgendaDate) Is Not Null))" _
        & " ORDER BY [PECHANGA\jgarcia_factCapitalProjects].BoardAgendaDate DESC;"
        Me.cboFilterInput1.Requery
    
    Case Else
        cboFilterInput3.RowSource = "SELECT [PECHANGA\jgarcia_dimCapExLookups].LookupValue" _
        & " FROM [PECHANGA\jgarcia_dimCapExLookups]" _
        & " WHERE ((([PECHANGA\jgarcia_dimCapExLookups].Section)=[Forms]![frmCapitalProjects]![cboFilter3]));"
        Me.cboFilterInput3.Requery

    End Select
End Sub

Private Sub cboFilter1_Change()
Me.cboFilterInput1 = Null
End Sub
Private Sub cboFilter2_Change()
Me.cboFilterInput2 = Null
End Sub
Private Sub cboFilter3_Change()
Me.cboFilterInput3 = Null
End Sub


Private Sub cboFilter1_LostFocus()
Me.cboFilterInput1.Requery
End Sub


Private Sub cboFilter2_LostFocus()
Me.cboFilterInput2.Requery
End Sub

Private Sub cboFilter3_LostFocus()
Me.cboFilterInput3.Requery
End Sub

Private Sub cboVendor_GotFocus()
Me.cboVendor.Requery
End Sub

Private Sub Check718_AfterUpdate()
'DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub cmdClearFilter_Click()
Me.Filter = ""
Me.FilterOn = False

'Clear filter options
Me.cboFilter1 = Null
Me.cboFilter2 = Null
Me.cboFilter3 = Null

'Clear filter values
Me.cboFilterInput1 = Null
Me.cboFilterInput2 = Null
Me.cboFilterInput3 = Null

End Sub

Private Sub cmdFilter_Click()
If IsNull(Me.cboFilter1) = False And IsNull(Me.cboFilterInput1) = False Then
    If IsNull(Me.cboFilter2) = False And IsNull(Me.cboFilterInput2) = False Then
        If IsNull(Me.cboFilter3) = False And IsNull(Me.cboFilterInput3) = False Then
            Me.Filter = FilterFormula(Me.cboFilter1, Me.cboFilterInput1.Value) _
                & " AND " & FilterFormula(Me.cboFilter2, Me.cboFilterInput2.Value) _
                & " AND " & FilterFormula(Me.cboFilter3, Me.cboFilterInput3.Value)
            Else
                Me.Filter = FilterFormula(Me.cboFilter1, Me.cboFilterInput1.Value) _
                & " AND " & FilterFormula(Me.cboFilter2, Me.cboFilterInput2.Value)
            End If
    Else
        Me.Filter = FilterFormula(Me.cboFilter1, Me.cboFilterInput1.Value)
    End If
Else
    Exit Sub
End If

Me.FilterOn = True
Me.cboFilter1.SetFocus

End Sub

Private Sub cmdNewRecord_Click()
DoCmd.GoToRecord , , acNewRec
Me.CarNumber.SetFocus
End Sub

Private Sub cmdSaveRecord_Click()
DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub cmdSpendDetail_Click()
DoCmd.OpenForm "frmSpendingDetailByProject", acFormDS, , , acFormEdit
End Sub

Private Sub cmdUpdateDates_Click()
DoCmd.OpenForm "frmUpdateApprovalDates", acNormal, , , acFormEdit
End Sub

Private Sub CompletionDate_AfterUpdate()
If IsNull(Me.CompletionDate) = False Then
    Call CreateReleasedFunds(Me.CarNumber.Value, Me.CompletionDate.Value)
End If
End Sub

Private Sub CompletionDate_BeforeUpdate(Cancel As Integer)

Dim ProjectCompleted As Integer

Beep
ProjectCompleted = MsgBox("Once a project is marked as completed, the date cannot be changed and funds will be released. Continue?", vbYesNo + vbDefaultButton2, "Mark as Completed?")
    If ProjectCompleted <> vbYes Then
        Cancel = True
        Me.Undo
        Exit Sub
    Else
        Cancel = False
    End If

If IsNull(Me.CompletionDate) = False And Me.CompletionDate.OldValue <> Me.CompletionDate.Value Then
    Beep
    MsgBox "Project has already been marked as completed! Unable to change.", vbCritical
    Me.Undo
    Cancel = True
    Exit Sub
End If
End Sub

Private Sub Form_Current()

'update CAR filename
Me.txtCARfilename.Value = [Forms]![frmCapitalProjects]![CarNumber] & " " & [Forms]![frmCapitalProjects]![ProjectDescription]
Me.txtCARfilename.Requery

If IsNull(Me.CarNumber) = False Then
    'update Remaining Balance
    If IsNull(DLookup("[Remaining Balance]", "qryPTBudgetRemainingBalances", "[CarNumber] = '" & Me.CarNumber & "'")) = False Then
        Me.txtRemainingBal.Value = Format(DLookup("[Remaining Balance]", "qryPTBudgetRemainingBalances", "[CarNumber] = '" & Me.CarNumber & "'"), "Currency", "$0.00")
    Else
        Me.txtRemainingBal.Value = 0
    End If
    
    'update Current Balance
    If IsNull(DLookup("[Current Balance]", "qryPTBudgetRemainingBalances", "[CarNumber] = '" & Me.CarNumber & "'")) = False Then
        Me.txtCurrentBal.Value = Format(DLookup("[Current Balance]", "qryPTBudgetRemainingBalances", "[CarNumber] = '" & Me.CarNumber & "'"), "Currency", "$0.00")
    Else
        Me.txtCurrentBal.Value = 0
    End If
End If

End Sub

Private Sub Form_Delete(Cancel As Integer)

Dim ConfirmDelete As Integer

Beep
ConfirmDelete = MsgBox("Are you sure you want to delete this record?", vbYesNoCancel + vbDefaultButton2, "Delete Record?")

If ConfirmDelete = vbYes Then
    'Me.OnCurrent = ""
    Cancel = False
Else
    Cancel = True
    Exit Sub

End If


End Sub

Private Sub Form_Dirty(Cancel As Integer)
Me.OnCurrent = "[Event Procedure]"
End Sub


Private Sub Form_Load()
On Error Resume Next
Me.Filter = FilterFormula(Me.cboFilter1, Me.cboFilterInput1.Value)
Me.FilterOn = True

Select Case User
    Case "jgarcia"
        Me.cmdUpdateDates.Visible = True
    Case Else
        Me.cmdUpdateDates.Visible = False
    End Select

End Sub

Private Sub Form_Open(Cancel As Integer)
'Me.WindowHeight = 4000
'Me.InsideHeight = 10000
'Me.InsideWidth = 15000
End Sub

Private Sub lblCARFile_Click()
Dim strFolderName As String
Dim strFolderExists As String

Dim wsh As Object
Set wsh = VBA.CreateObject("WScript.Shell")
Dim waitOnReturn As Boolean: waitOnReturn = True
Dim windowStyle As Integer: windowStyle = 1

    strFolderName = "\\prcfile\DeptFolders\FinAnalyst\Capital\" _
        & [Forms]![frmCapitalProjects]![FiscalYear] & "\4_CAR Supporting Documentation\" _
        & [Forms]![frmCapitalProjects]![CarNumber] & " " & [Forms]![frmCapitalProjects]![ProjectDescription]
    strFolderExists = Dir(strFolderName, vbDirectory)

    If strFolderExists = "" Then
        Beep
        MsgBox "The selected folder doesn't exist! Creating folder now."
        wsh.Run "cmd /c md """ & strFolderName & "", windowStyle, waitOnReturn
        'Shell "cmd /c md """ & strFolderName & ""
        'MkDir (strFolderName)
    End If
End Sub

Private Sub txtCARfilename_Click()

Me.txtCARfilename.Value = [Forms]![frmCapitalProjects]![CarNumber] & " " & [Forms]![frmCapitalProjects]![ProjectDescription]
Me.txtCARfilename.Requery
Me.txtCARfilename.SelStart = 0
Me.txtCARfilename.SelLength = Len(Me.txtCARfilename)

End Sub

Private Sub txtCARfilename_GotFocus()
Me.txtCARfilename.Value = [Forms]![frmCapitalProjects]![CarNumber] & " " & [Forms]![frmCapitalProjects]![ProjectDescription]
Me.txtCARfilename.Requery
Me.txtCARfilename.SelStart = 0
Me.txtCARfilename.SelLength = Len(Me.txtCARfilename)
End Sub


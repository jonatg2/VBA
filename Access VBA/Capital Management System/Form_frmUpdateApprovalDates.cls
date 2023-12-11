VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmUpdateApprovalDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdUpdateCARdates_Click()
    If IsNull(Me.cboStatus) = True Then
        Beep
        MsgBox "Please select a status!"
        Exit Sub
        
    ElseIf Me.cboStatus = "Submitted" And (IsNull(Me.txtCSCAgendaDate) = True Or IsNull(Me.txtBoardAgendaDate) = True Or _
                                IsNull(Me.txtApprovalDate) = True) Then
        Beep
        MsgBox "Please add all dates for Submitted CAR statuses!"
        Exit Sub
    
    ElseIf Me.cboStatus = "Pending" And (IsNull(Me.txtBoardAgendaDate) = True Or IsNull(Me.txtApprovalDate) = True) Then
        Beep
        MsgBox "Please add board agenda dates and approval dates for Pending CAR statuses!"
        Exit Sub
    
    ElseIf Me.cboStatus.Value = "Update Board Agenda Dates" And (IsNull(Me.txtBoardAgendaDate) = True Or IsNull(Me.txtBoardAgendaDateNew) = True) Then
        Beep
        MsgBox "Please add both current and new board agenda dates to update!"
        Exit Sub
    
    ElseIf Me.cboStatus.Value = "Update CSC Agenda Dates" And (IsNull(Me.txtCSCAgendaDate) = True Or IsNull(Me.txtCscAgendaDateNew) = True) Then
        Beep
        MsgBox "Please add both current and new CSC agenda dates to update!"
        Exit Sub
    
    ElseIf Me.cboStatus.Value = "Pending" Then
        Me.txtCSCAgendaDate.Value = Null
        Call UpdateCARdates(Me.cboStatus, Me.txtBoardAgendaDate, Me.txtApprovalDate)
    
    ElseIf Me.cboStatus.Value = "Update Board Agenda Dates" Then
        Call UpdateCARdates(Me.cboStatus, Me.txtBoardAgendaDate, , , , Me.txtBoardAgendaDateNew)
    
    ElseIf Me.cboStatus.Value = "Update CSC Agenda Dates" Then
        Call UpdateCARdates(Me.cboStatus, , , Me.txtCSCAgendaDate, Me.txtCscAgendaDateNew)
    
    Else
        Me.txtBoardAgendaDateNew.Value = Null
        Call UpdateCARdates(Me.cboStatus, Me.txtBoardAgendaDate, Me.txtApprovalDate, Me.txtCSCAgendaDate)
    End If
End Sub

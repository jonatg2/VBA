VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLookups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cboSection_AfterUpdate()
Me.cboDivision.Requery

End Sub

Private Sub cboSectionFilter_AfterUpdate()
If IsNull(Me.cboSectionFilter) = False Then
    Me.Filter = "[Section] = '" & Me.cboSectionFilter.Value & "'"
    Me.FilterOn = True
Else
    Me.Filter = ""
    Me.FilterOn = False
End If
End Sub

Private Sub cmdNewRecord_Click()
DoCmd.RunCommand acCmdRecordsGoToNew
End Sub

Private Sub cmdSaveRecord_Click()
DoCmd.Save
DoCmd.RunCommand acCmdSaveRecord
End Sub

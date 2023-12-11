VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmImportData_check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub txtImportDate_AfterUpdate()
If IsNull(Me.txtImportDate) = False Then
    Me.txtImportDate.Value = GetNowLast(Me.txtImportDate)
    Me.subfrmUnionRollups.Form.Requery
    Me.subfrmUnionTotals.Form.Requery
End If
End Sub


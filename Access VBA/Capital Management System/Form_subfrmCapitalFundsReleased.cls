VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subfrmCapitalFundsReleased"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ReleaseMonth_AfterUpdate()

Select Case Month(Me.ReleaseMonth.Value)
    Case 10, 11, 12
        Me.FiscalYear = Year(Me.ReleaseMonth.Value) + 1
    Case Else
        Me.FiscalYear = Year(Me.ReleaseMonth.Value)
End Select
    
End Sub

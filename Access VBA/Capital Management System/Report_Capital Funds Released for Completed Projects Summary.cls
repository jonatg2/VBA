VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Capital Funds Released for Completed Projects Summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkExcludeZeros_Click()
Me.Requery

End Sub

Private Sub Report_Load()

Dim MonthSpend As String
MonthSpend = Forms!frmReportNavigation!txtMonthEndDate
Me.lblSpendThrough.Caption = "Completed as of " & Format(MonthSpend, "MMM YYYY")

End Sub

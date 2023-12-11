VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Active Capital Projects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CarNumber_Click()
DoCmd.OpenReport "rptNotesLookup", acViewReport
End Sub

Private Sub Report_Load()

Dim MonthSpend As String
MonthSpend = Forms!frmReportNavigation!txtMonthEndDate
Me.lblSpendThrough.Caption = "Spend through " & Format(MonthSpend, "MMM YYYY")

End Sub

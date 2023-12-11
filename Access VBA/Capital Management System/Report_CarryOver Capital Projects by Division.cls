VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_CarryOver Capital Projects by Division"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Report_Load()

Dim FiscalYear As String
FiscalYear = Forms!frmReportNavigation!cboFiscalYear
Me.lblSpendThrough.Caption = FiscalYear & " Fiscal Year"

End Sub


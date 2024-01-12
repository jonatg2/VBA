VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Project Spending Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Report_Load()

Dim MaxEndDate As Date
MaxEndDate = DMax("MonthEndDate", "PECHANGA\jgarcia_factCapitalSpending")

Me.lblProjectName.Caption = [Forms]![frmCapitalProjects]![CarNumber] & " " & [Forms]![frmCapitalProjects]![ProjectDescription]
Me.lblSpendThrough.Caption = "Spend through " & Format(MaxEndDate, "MMM YYYY")

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Capital Projects Tentative Schedules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Report_Load()

Dim BoardAgendaDate As Date

BoardAgenda = Forms!frmReportNavigation!cboBoardAgendaDate
Me.lblSpendThrough.Caption = "Board Agenda: " & Format(BoardAgenda, "MMM DD, YYYY")
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_Active Capital Projects Summary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Report_Load()
Dim MonthSpend As String
MonthSpend = Forms!frmReportNavigation!txtMonthEndDate
Dim ReportWidthHide, ReportWidthUnhide As Integer
Dim ctl As Control

Me.lblSpendThroughRollup.Caption = "Spend through " & Format(MonthSpend, "MMM YYYY")

ReportWidthUnhide = 23690
ReportWidthHide = 16990

    If Forms!frmReportNavigation!chkHideColumns = -1 Then
        
        'to hide columns that don't need to be displayed and shrink page width
        For Each ctl In Me.Controls
            Select Case ctl.Tag
                Case "Hide_Group"
                    ctl.Width = 1
                    ctl.Visible = False
                Case "Hide_Percent"
                    ctl.Width = 1
                    ctl.Visible = False
                End Select
        Next
        Me.Report.Width = ReportWidthHide
        Me.txtCarryOverClass.Width = ReportWidthHide
    Else
        
        'to unhide columns
        For Each ctl In Me.Controls
            Select Case ctl.Tag
                Case "Hide_Group", "Report_Group"
                    ctl.Width = 1900
                    ctl.Visible = True
                Case "Hide_Percent", "Percent_Group"
                    ctl.Width = 900
                    ctl.Visible = True
                Case "Balance_Group"
                    ctl.Width = 1900
                End Select
            
        Next
'        Debug.Print (Me.lblDivSubtotal.Width)
        Me.Report.Width = ReportWidthUnhide
        Me.txtCarryOverClass.Width = ReportWidthUnhide
    
    End If
End Sub



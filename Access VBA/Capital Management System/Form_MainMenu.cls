VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdMinimize_Click()
If Me.Tag = 1 Then
    Me.Modal = False
    DoCmd.RunCommand acCmdAppMinimize
    Forms!MainMenu.SetFocus
    DoCmd.Restore
Else
    DoCmd.Close
    DoCmd.OpenForm "MainMenu", acNormal
'    DoCmd.RunCommand acCmdAppMinimize
'    Forms!MainMenu.SetFocus
'    DoCmd.Restore
'    Call ActivateAccessApp
End If

End Sub

Private Sub cmdOpenBudgetList_Click()
DoCmd.OpenForm "frmCapitalBudgets", acNormal, , , acFormEdit
End Sub

Private Sub cmdOpenImportMenu_Click()

Select Case User
    Case "jgarcia"
        DoCmd.OpenForm "frmImportData", acNormal, , , acFormEdit
        DoCmd.OpenForm "frmImportData_check", acNormal, , , acFormEdit
    Case Else
        DoCmd.OpenForm "frmImportData", acNormal, , , acFormEdit
    End Select



End Sub

Private Sub cmdOpenListForm_Click()
DoCmd.OpenForm "frmCapitalProjects", acNormal, , , acFormEdit
End Sub

Private Sub cmdOpenLookups_Click()
DoCmd.OpenForm "frmLookups", acNormal, , , acFormEdit
End Sub

Private Sub cmdOpenReports_Click()
DoCmd.OpenForm "frmReportNavigation", acNormal, , , acFormEdit
End Sub

Private Sub cmdPermissions_Click()
DoCmd.OpenForm "frmEmailContacts", acFormDS, , , acFormEdit
End Sub

Private Sub Form_Load()
If Me.Tag = 1 Then
    If Right(Application.CurrentProject.Name, 5) <> "accde" Then
        Beep
        MsgBox "Please use compiled version of program!", vbInformation, "Incorrect Version"
        DoCmd.Quit acQuitSaveAll
    End If
End If
DoCmd.RunCommand acCmdAppMinimize
Forms!MainMenu.SetFocus
DoCmd.Restore
Call ActivateAccessApp
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dim ExitProgram As Integer

If Me.Tag = 1 Then
        DoCmd.Quit acQuitSaveAll
        Exit Sub
    End If
'    Beep
'    ExitProgram = MsgBox("Do you wish to exit the program?", vbYesNo + vbDefaultButton2, "Exit Program?")
'        If ExitProgram = vbYes Then
'            DoCmd.Quit acQuitSaveAll
'        Else
'            Cancel = True
'            Exit Sub
'        End If
End Sub



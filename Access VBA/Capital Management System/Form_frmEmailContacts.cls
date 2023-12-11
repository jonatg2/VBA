VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEmailContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cboBadgeNumber_AfterUpdate()
Me.cboContactName = DLookup("[Employee Name]", "qryprcDimEmployeeMaster", "[Badge Number] =" & Me.cboBadgeNumber.Value)
Me.ContactEmail.Value = DLookup("[emailAddress]", "qryprcDimEmployeeMaster", "[Badge Number] =" & Me.cboBadgeNumber.Value)
Me.empTitle.Value = DLookup("[empTitle]", "qryprcDimEmployeeMaster", "[Badge Number] =" & Me.cboBadgeNumber.Value)
End Sub


Private Sub cboTableauDivision_AfterUpdate()
DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub cboTableauDivision_BeforeUpdate(Cancel As Integer)

Dim dbsCapital As DAO.Database
Dim rstUsers As DAO.Recordset
Dim listUsers, sqlSELECT, SQLUpdate As String

sqlSELECT = "SELECT ContactEmail,ContactName,TableauUserName FROM PECHANGA\jgarcia_dimCapExEmailContacts" _
    & " WHERE TableauDivision = '" & Nz(Me.cboTableauDivision.OldValue, "NULL") & "'" _
    & " AND ID <> " & Nz(Me.ID.Value, "NULL")
        
Set dbsCapital = CurrentDb
Set rstUsers = dbsCapital.OpenRecordset(sqlSELECT, dbOpenDynaset)
listUsers = ""

If rstUsers.EOF Then
'Debug.Print ("No records!")
DoCmd.RunSQL ("UPDATE PECHANGA\jgarcia_dimCapExLookups" _
        & " SET tableauGroupPermissions = NULL WHERE" _
        & " Section = 'Division'" _
        & " AND LookupValue = '" & Me.cboTableauDivision.OldValue & "'")
    Exit Sub
End If

Do While Not rstUsers.EOF
    listUsers = listUsers & Switch(IsNull(rstUsers!TableauUserName) = False, rstUsers!TableauUserName & ", ", IsNull(rstUsers!ContactEmail) = False, Replace(rstUsers!ContactEmail, "@pechanga.com", "") & ", ")
    rstUsers.MoveNext
Loop
'Debug.Print (listUsers)
rstUsers.Close

'update value in lookup table
DoCmd.RunSQL ("UPDATE PECHANGA\jgarcia_dimCapExLookups" _
        & " SET tableauGroupPermissions = '" & listUsers & "' WHERE" _
        & " Section = 'Division'" _
        & " AND LookupValue = '" & Me.cboTableauDivision.OldValue & "'")
    

End Sub

Private Sub cboTableauDivision_Exit(Cancel As Integer)

Dim dbsCapital As DAO.Database
Dim rstUsers As DAO.Recordset
Dim listUsers, sqlSELECT, SQLUpdate As String
    
    
'update record so it can be read in the DAO recordset
DoCmd.RunSQL ("UPDATE PECHANGA\jgarcia_dimCapExEmailContacts" _
            & " SET TableauDivision = '" & Me.cboTableauDivision.Value & "' WHERE" _
            & " ID = " & Nz(Me.ID.Value, "NULL"))

Set dbsCapital = CurrentDb

If IsNull(Me.cboTableauDivision) = False Then
            sqlSELECT = "SELECT ContactEmail,ContactName,TableauUserName FROM PECHANGA\jgarcia_dimCapExEmailContacts" _
                & " WHERE TableauDivision = '" & Me.cboTableauDivision.Value & "'"
        Set rstUsers = dbsCapital.OpenRecordset(sqlSELECT, dbOpenDynaset)
        listUsers = ""
    
        If rstUsers.EOF Then
            Exit Sub
End If
    
    Do While Not rstUsers.EOF
        listUsers = listUsers & Switch(IsNull(rstUsers!TableauUserName) = False, rstUsers!TableauUserName & ", ", IsNull(rstUsers!ContactEmail) = False, Replace(rstUsers!ContactEmail, "@pechanga.com", "") & ", ")
        rstUsers.MoveNext
    Loop
    'Debug.Print (listUsers)
    rstUsers.Close
    
    'update value in lookup table
    DoCmd.RunSQL ("UPDATE PECHANGA\jgarcia_dimCapExLookups" _
            & " SET tableauGroupPermissions = '" & listUsers & "' WHERE" _
            & " Section = 'Division'" _
            & " AND LookupValue = '" & Me.cboTableauDivision.Value & "'")
    
End If
End Sub

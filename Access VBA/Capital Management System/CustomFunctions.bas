Attribute VB_Name = "CustomFunctions"
Option Compare Database
Option Explicit
Public Function userName()

userName = TempVars("User") & "_" & TempVars("Workstation")

End Function
Public Function User()

User = TempVars("User")

End Function

Public Function FilterFormula(Filter, FValue As String) As String


Select Case Filter
    Case "FiscalYear"
        FilterFormula = "[FiscalYear]=" & FValue & ""
    
    Case "CAR_Status"
        FilterFormula = "[CarStatus]='" & FValue & "'"
    
    Case "Division"
        FilterFormula = "[Division]='" & FValue & "'"
    
    Case "Department"
        FilterFormula = "[Department]='" & FValue & "'"
    
    Case "SubGroup_1"
        FilterFormula = "[SubGroup_1]='" & FValue & "'"
    
    Case "ProjectType"
        FilterFormula = "[ProjectType]='" & FValue & "'"
    
    Case "ProjectDescription"
        FilterFormula = "[ProjectDescription] like '*" & FValue & "*'"
    Case "CARNumber"
        FilterFormula = "[CarNumber] like '*" & FValue & "*'"
    Case "CSCAgendaDate"
        FilterFormula = "[CscAgendaDate] = #" & FValue & "#"
    Case "BoardAgendaDate"
        FilterFormula = "[BoardAgendaDate] = #" & FValue & "#"
    Case "BudgetReference"
        FilterFormula = "[budgetReference] = '" & FValue & "'"
        
        
End Select
End Function


Public Function GetNowLast(InputDate As Date) As Date
Dim Dyear, dMonth As Integer
Dim getDate As Date

    Dyear = Year(InputDate)
    dMonth = Month(InputDate)

    getDate = DateSerial(Dyear, dMonth + 1, 0)

    GetNowLast = getDate

End Function
Public Function OpenExcelFile(strFilePath As String) As Boolean
    'Required: Tools > Refences: Add reference to Microsoft Excel Object Library
    
    Dim appExcel As Excel.Application
    Dim myWorkbook As Excel.Workbook

    Set appExcel = CreateObject("Excel.Application")
    Set myWorkbook = appExcel.Workbooks.Open(strFilePath)
    appExcel.Visible = True
    
    'Do Something or Just Leave Open
    Set appExcel = Nothing
    Set myWorkbook = Nothing
End Function

Public Function EmailToList(Division As String) As String
    Dim dbsCapital As DAO.Database
    Dim rstEmailTo As DAO.Recordset
    Dim emailTo, sqlSELECT As String
    
    sqlSELECT = "SELECT DISTINCT BadgeNumber, ContactEmail, Division, emailTo" _
                & " FROM PECHANGA\jgarcia_dimCapExEmailContacts" _
                & " WHERE (Division = '" & Division & "' OR Department = '1540 - PLANNING AND ANALYSIS')" _
                & " AND emailTo = TRUE"
    
    Set dbsCapital = CurrentDb
    Set rstEmailTo = dbsCapital.OpenRecordset(sqlSELECT, dbOpenDynaset)
    emailTo = ""
    
    If rstEmailTo.EOF Then
        Exit Function
    End If
    
    Do While Not rstEmailTo.EOF
        'Debug.Print (rstEmailTo!ContactEmail & ", " & rstEmailTo!EmailTo)
        emailTo = emailTo & rstEmailTo!ContactEmail & "; "
        rstEmailTo.MoveNext
    Loop
    rstEmailTo.Close
    
    'Debug.Print (EmailTo)
    EmailToList = emailTo
End Function

Public Function EmailCcList(Division As String) As String
    Dim dbsCapital As DAO.Database
    Dim rstEmailCc As DAO.Recordset
    Dim emailCc, sqlSELECT As String
    
    sqlSELECT = "SELECT DISTINCT BadgeNumber, ContactEmail, Division, EmailCc" _
                & " FROM PECHANGA\jgarcia_dimCapExEmailContacts" _
                & " WHERE (Division = '" & Division & "' OR Department = '1540 - PLANNING AND ANALYSIS')" _
                & " AND EmailCc = TRUE"
    
    Set dbsCapital = CurrentDb
    Set rstEmailCc = dbsCapital.OpenRecordset(sqlSELECT, dbOpenDynaset)
    emailCc = ""
    
    If rstEmailCc.EOF Then
        Exit Function
    End If
    
    Do While Not rstEmailCc.EOF
        'Debug.Print (rstEmailCc!ContactEmail & ", " & rstEmailCc!EmailCc)
        emailCc = emailCc & rstEmailCc!ContactEmail & "; "
        rstEmailCc.MoveNext
    Loop
    rstEmailCc.Close
    
    'Debug.Print (EmailCc)
    EmailCcList = emailCc

End Function

Public Function FiscalYearCalc(InputDate As Date) As Integer
Dim MonthNum As Integer

MonthNum = Month(InputDate)

Select Case MonthNum
    Case 10, 11, 12
        FiscalYearCalc = Year(InputDate) + 1
    Case Else
        FiscalYearCalc = Year(InputDate)
End Select
'Debug.Print FiscalYearCalc
End Function

Public Function LookupReleasedFunds(CarNumber As String) As Long
LookupReleasedFunds = DLookup("[FundsToRelease]", "qryFundsToRelease", "CarNumber = '" & CarNumber & "'")
Debug.Print (LookupReleasedFunds)
End Function

Public Function FiscalMonthSort(InputDate As Date) As Integer

Select Case Month(InputDate)
    Case 10, 11, 12
        FiscalMonthSort = Month(InputDate) - 9
    Case Else
        FiscalMonthSort = Month(InputDate) + 3
    End Select
End Function



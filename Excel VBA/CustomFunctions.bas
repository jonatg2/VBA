Attribute VB_Name = "CustomFunctions"
Option Explicit
 Function TransactionClass(aString As String, PostingDate As Date) As String
    Dim myExpression As Variant
    Dim regEx As RegExp
    Dim re(20, 1) As String
    Dim matches, submatches As Object
    Dim i, j As Integer
    
    Set regEx = New RegExp
    
    'regEx expressions column'
    re(0, 0) = "Qtr(.*)Divers(.*)(Fu(.*)|Fd)" 'Pechanga Diversification Fund
    re(1, 0) = "((.*)?(PAYROLL|Payroll|PYRL|PAYR)(.*)?(Tax)?(P/E)?(.*)?)|(PAYROLL|Payroll)(.*)CLOSE TO GENERAL|Retire(ment)?(.*)?Match" 'Longevity/Bonuses Paid
    re(2, 0) = "Accrued(.*)Payroll Taxes SUI" 'EDD SUI Payments
    re(3, 0) = "Liab(.*)State Revenue Share" 'CA Revenue Sharing Payments
    re(4, 0) = "((LCCF|PCF) QE (.*))|(PCF(.*)Quarters?)" 'Pechanga Community Fund
    re(5, 0) = "(.*)?(Pay(ment|down)?|Pmt|Refi)" 'Loan Payments
    re(6, 0) = "(.*)?Draw(.*)?" 'Loan Draws
    re(7, 0) = "^(County of Riverside(.*)?IGA)|((.*)?City of (Temecula|Enforceme?nt)(.*)?)" 'IGA - City of Temecula'
    re(8, 0) = "QE(.*)|([0-9]{1,2}\/[0-9]{1,2}\/[0-9]{1,2})(.*)?" 'Compact Fees'
    re(9, 0) = "(.*)?Ded(icated)?(.*)(Svc|Serv(ice)?)?(.*)?|(.*)?Mgmn?t Inc(en)?(tive)?(.*)?" 'Longevity/Bonuses Accrued
    re(10, 0) = "(.*)?(Libor)?(.*)?Loan(.*)?(Interest)?(.*)?" 'Interest Payments
    re(11, 0) = ""
    re(12, 0) = ""
    re(13, 0) = ""
    re(14, 0) = ""
    re(15, 0) = ""
    re(16, 0) = ""
    re(17, 0) = ""
    re(18, 0) = ""
    re(19, 0) = ""
    re(20, 0) = ""
    
    'Transactions classification column'
    re(0, 1) = "Pechanga Diversification Fund"
    re(1, 1) = "Longevity/Bonuses Paid"
    re(2, 1) = "EDD SUI Payments"
    re(3, 1) = "CA Revenue Sharing Payments"
    re(4, 1) = "Pechanga Community Fund"
    re(5, 1) = "Loan Payments"
    re(6, 1) = "Loan Draws"
    re(7, 1) = "IGA - City of Temecula"
    re(8, 1) = "Compact Fees"
    re(9, 1) = "Longevity/Bonuses Accrued"
    re(10, 1) = "Interest Payments"
    re(11, 1) = ""
    re(12, 1) = ""
    re(13, 1) = ""
    re(14, 1) = ""
    re(15, 1) = ""
    re(16, 1) = ""
    re(17, 1) = ""
    re(18, 1) = ""
    re(19, 1) = ""
    re(20, 1) = ""
    
    regEx.IgnoreCase = True
    regEx.Global = True
       
    'to loop through regex formulas and find corresponding classification
    For i = LBound(re, 1) To UBound(re, 1)
        regEx.Pattern = re(i, 0)
        If regEx.Test(aString) = True Then
            TransactionClass = re(i, 1)
            Exit For
        Else
            TransactionClass = ""
        End If
        
    Next

End Function

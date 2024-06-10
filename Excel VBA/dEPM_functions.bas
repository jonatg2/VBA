Attribute VB_Name = "dEPM_functions"
Option Explicit

Function Scenario(ScenarioType as String) As String
    Select Case ScenarioType
        Case "Actuals"
            Scenario = "[GLTOT_SCENARIO].[PRC].[PRC/1].[PRC/2]"
        Case "OpExBudget"
            Scenario = "[GLTOT_SCENARIO].[PRC].[PRC/5]"
        Case "WorkingBudget"
            Scenario = "[GLTOT_SCENARIO].[PRC].[PRC/23]"
        Case "CapExBudget"
            Scenario = "[GLTOT_SCENARIO].[PRC].[PRC/7]"
    End Select
End Function

Function CalendarPeriod(Cube As String, Period As String, InputDate As Date) As String
        Dim PeriodClass(11, 2) As String
        Dim RootFormula, RootFormulaLTD, RootFormulaYTD, RootYear, RootYearLTD,RootYearYTD, sMonth, sYear, sQuarter, sMonth2 As String
        Dim i As Integer
        
    'Root Formulas
        sMonth = Month(InputDate)
        
        If Month(InputDate) = 10 Or Month(InputDate) = 11 Or Month(InputDate) = 12 Then
            sYear = Year(InputDate) + 1
        Else
            sYear = Year(InputDate)
        End If
        
        RootFormula = "[" & Cube & "_CALENDARPERIOD].[PRC].[PRC/2_TOP_NODE]."
        RootFormulaLTD = "[" & Cube & "_CALENDARPERIOD].[PRC].[PRC/2_TOP_NODE_LTD]."
        RootFormulaYTD = "[" & Cube & "_CALENDARPERIOD].[PRC].[PRC/2_TOP_NODE_YTD]."
        RootYear = RootFormula & "[PRC/2_" & sYear & "]"
        RootYearLTD = RootFormulaLTD & "[PRC/2_" & sYear & "_LTD]"
        RootYearYTD = RootFormulaYTD & "[PRC/2_" & sYear & "_YTD]"

    'Calendar Month Dimension'
        PeriodClass(0, 0) = "10"
        PeriodClass(1, 0) = "11"
        PeriodClass(2, 0) = "12"
        PeriodClass(3, 0) = "1"
        PeriodClass(4, 0) = "2"
        PeriodClass(5, 0) = "3"
        PeriodClass(6, 0) = "4"
        PeriodClass(7, 0) = "5"
        PeriodClass(8, 0) = "6"
        PeriodClass(9, 0) = "7"
        PeriodClass(10, 0) = "8"
        PeriodClass(11, 0) = "9"
        
    'Fiscal Quarter Dimension'
        PeriodClass(0, 1) = "Q1"
        PeriodClass(1, 1) = "Q1"
        PeriodClass(2, 1) = "Q1"
        PeriodClass(3, 1) = "Q2"
        PeriodClass(4, 1) = "Q2"
        PeriodClass(5, 1) = "Q2"
        PeriodClass(6, 1) = "Q3"
        PeriodClass(7, 1) = "Q3"
        PeriodClass(8, 1) = "Q3"
        PeriodClass(9, 1) = "Q4"
        PeriodClass(10, 1) = "Q4"
        PeriodClass(11, 1) = "Q4"
        
    'Fiscal Month Dimension'
        PeriodClass(0, 2) = "M01"
        PeriodClass(1, 2) = "M02"
        PeriodClass(2, 2) = "M03"
        PeriodClass(3, 2) = "M04"
        PeriodClass(4, 2) = "M05"
        PeriodClass(5, 2) = "M06"
        PeriodClass(6, 2) = "M07"
        PeriodClass(7, 2) = "M08"
        PeriodClass(8, 2) = "M09"
        PeriodClass(9, 2) = "M10"
        PeriodClass(10, 2) = "M11"
        PeriodClass(11, 2) = "M12"
        
        'to loop through array and find corresponding classification
    Select Case Period

        'Yearly
        Case Is = "Y"
            CalendarPeriod = RootYear

        'Quarterly
        Case Is = "Q"
        For i = LBound(PeriodClass, 1) To UBound(PeriodClass, 1)
                If PeriodClass(i, 0) = sMonth Then
                    sQuarter = PeriodClass(i, 1)
                    CalendarPeriod = RootYear & ".[PRC/2_" & sYear & sQuarter & "]"
                    Exit For
                End If
        Next
        
        'Monthly
        Case Is = "M"
            For i = LBound(PeriodClass, 1) To UBound(PeriodClass, 1)
                    If PeriodClass(i, 0) = sMonth Then
                        sQuarter = PeriodClass(i, 1)
                        sMonth2 = PeriodClass(i, 2)
                        CalendarPeriod = RootYear & ".[PRC/2_" & sYear & sQuarter & "]" _
                        & ".[PRC/2_" & sYear & sMonth2 & "]"
                        Exit For
                    End If
            Next
        
        'Month to Date
        Case Is = "MTD"
            For i = LBound(PeriodClass, 1) To UBound(PeriodClass, 1)
                    If PeriodClass(i, 0) = sMonth Then
                        sQuarter = PeriodClass(i, 1)
                        sMonth2 = PeriodClass(i, 2)
                        CalendarPeriod = RootYear & ".[PRC/2_" & sYear & sQuarter & "_MTD]" _
                        & ".[PRC/2_" & sYear & sMonth2 & "_MTD]"
                        Exit For
                    End If
            Next

        'Life to Date
        Case Is = "LTD"
            For i = LBound(PeriodClass, 1) To UBound(PeriodClass, 1)
                    If PeriodClass(i, 0) = sMonth Then
                        sQuarter = PeriodClass(i, 1)
                        sMonth2 = PeriodClass(i, 2)
                        CalendarPeriod = RootYearLTD & ".[PRC/2_" & sYear & sQuarter & "_LTD]" _
                        & ".[PRC/2_" & sYear & sMonth2 & "_LTD]"
                        Exit For
                    End If
            Next
        
        'Year to Date
        Case Is = "YTD"
            For i = LBound(PeriodClass, 1) To UBound(PeriodClass, 1)
                    If PeriodClass(i, 0) = sMonth Then
                        sQuarter = PeriodClass(i, 1)
                        sMonth2 = PeriodClass(i, 2)
                        CalendarPeriod = RootYearYTD & ".[PRC/2_" & sYear & sQuarter & "_YTD]" _
                        & ".[PRC/2_" & sYear & sMonth2 & "_YTD]"
                        Exit For
                    End If
            Next
    End Select

End Function

Function glMeasures(Account As String) As String
On Error GoTo ErrHandler

Select Case Int(Left(Account, 6))
    Case Is >= 990000
        glMeasures = "[GLTOT_MEASURES].[UNITSAMOUNT]"
    Case Is < 990000
        glMeasures = "[GLTOT_MEASURES].[FUNCTIONALAMOUNT]"
    End Select

ErrHandler:
    Select Case Err.Number
        Case Is = 13
            glMeasures = "[GLTOT_MEASURES].[FUNCTIONALAMOUNT]"
    End Select

End Function

Function ChartAccount(Account As String) As String
    On Error Resume Next
    Dim RootFormula As String
    Dim AccountInt As Long
        
    'Formulas
    AccountInt = Int(Account)



    'Case statement to determine syntax

    Select Case AccountInt

        'Cash on Hand
        Case 100000 To 100999
            If AccountInt = 100000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_100000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_100000].[PRC/" & Account & "]"
            End If
        
        'Cash in Bank
        Case 101000 To 101999
            If AccountInt = 101000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_101000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_101000].[PRC/" & Account & "]"
            End If
            
        'Cash Clearing
        Case 102000 To 102999
            If AccountInt = 102000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_102000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_102000].[PRC/" & Account & "]"
            End If
            
        'Accounts Receivable
        Case 103000 To 103999
            If AccountInt = 103000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_103000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_103000].[PRC/" & Account & "]"
            End If
            
        'Prepaid Expenses
        Case 104000 To 104999
            If AccountInt = 104000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_104000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_104000].[PRC/" & Account & "]"
            End If
        
        'Inventory
        Case 105000 To 105999
            If AccountInt = 105000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_105000]"
            ElseIf AccountInt = 105100 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_105000].[PRC/2_105100]"
            
            'Inventory - Gaming
            ElseIf AccountInt >= 105110 And AccountInt <= 105150 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_105000].[PRC/2_105100].[PRC/" & Account & "]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_105000].[PRC/" & Account & "]"
            End If
        
        'Fixed assets
        Case 106000 To 106999
            If AccountInt = 106000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_106000]"
            
            'Buildings
            ElseIf AccountInt = 106220 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_106000].[PRC/2_106220]"
            ElseIf AccountInt >= 106240 And AccountInt <= 106360 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_106000].[PRC/2_106220].[PRC/" & Account & "]"

            'Furniture, Fixtures & Equipment
            ElseIf AccountInt = 106480 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_106000].[PRC/2_106480]"
            ElseIf AccountInt >= 106500 And AccountInt <= 106760 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_106000].[PRC/2_106480].[PRC/" & Account & "]"
            
            'Land & Improvements
            ElseIf AccountInt = 106800 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_106000].[PRC/2_106800]"
            ElseIf AccountInt >= 106820 And AccountInt <= 106880 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_106000].[PRC/2_106800].[PRC/" & Account & "]"
            
            'Utility Improvements
            ElseIf AccountInt = 106900 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_106000].[PRC/2_106900]"
            ElseIf AccountInt >= 106420 And AccountInt <= 106460 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_106000].[PRC/2_106900].[PRC/" & Account & "]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_106000].[PRC/" & Account & "]"
            End If
                
        'Accumulated Depreciation
        Case 107000 To 107999
            If AccountInt = 107000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000]"
            
            'Construction in Progress
            ElseIf AccountInt = 107100 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/2_107100]"
            ElseIf AccountInt >= 107110 And AccountInt <= 107140 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/2_107100].[PRC/" & Account & "]"
            
            'Buildings
            ElseIf AccountInt = 107220 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/2_107220]"
            ElseIf AccountInt >= 107240 And AccountInt <= 107360 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/2_107220].[PRC/" & Account & "]"
            
            'Furniture, Fixtures & Equipment
            ElseIf AccountInt = 107480 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/2_107480]"
            ElseIf AccountInt >= 107500 And AccountInt <= 107770 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/2_107480].[PRC/" & Account & "]"
            
            'Land & Improvements
            ElseIf AccountInt = 107800 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/2_107800]"
            ElseIf AccountInt >= 107820 And AccountInt <= 107880 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/2_107800].[PRC/" & Account & "]"
            
            'Utility Improvements
            ElseIf AccountInt = 107900 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/2_107900]"
            'ElseIf AccountInt >= 107420 And AccountInt <= 107470 Then
                'RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/2_107900].[PRC/" & Account & "]"
            
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/" & Account & "]"
            End If
        
        'Other
        Case 109000 To 109999
            If AccountInt = 109000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_109000]"
            
            'FA - Loan Inception Costs/Uniforms
            ElseIf AccountInt = 109120 Or AccountInt = 109170 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_106000].[PRC/" & Account & "]"
            
            'A/D - Loan Inception Costs/Uniforms
            ElseIf AccountInt = 109130 Or AccountInt = 109180 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/" & Account & "]"
            
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_109000].[PRC/" & Account & "]"
            End If
            
        'Equity
        Case 300500 To 304000
            RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_EQUITY].[PRC/" & Account & "]"
        
        'Accrued Payroll Liabilities
        Case 202110 To 202980
            
            'Gratuity Liability
            If AccountInt >= 202780 And AccountInt <= 202980 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_ACCRUED PAYROLL LIABILIT].[PRC/2_202770].[PRC/" & Account & "]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_ACCRUED PAYROLL LIABILIT].[PRC/" & Account & "]"
            End If
        
        'Current Liabilities
        Case 201000 To 201990
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_CURRENT LIABILITIES].[PRC/" & Account & "]"
        
        'Gaming Liabilities
        Case 203000 To 203460
            If AccountInt = 203000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_LIABILITIES].[PRC/2_203000]"
            
            'Liab - Slots
            ElseIf AccountInt >= 203050 And AccountInt <= 203180 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_LIABILITIES].[PRC/2_203000].[PRC/2_203050].[PRC/" & Account & "]"
            
            'Liab - Table Games
            ElseIf AccountInt = 203200 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_LIABILITIES].[PRC/2_203000].[PRC/2_203200]"
            ElseIf AccountInt >= 203220 And AccountInt <= 203400 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_LIABILITIES].[PRC/2_203000].[PRC/2_203200].[PRC/" & Account & "]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_LIABILITIES].[PRC/2_203000].[PRC/" & Account & "]"
            End If
            
        'A/P Intercompany
        Case 203650 To 203790
            If AccountInt = 203650 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_LIABILITIES].[PRC/2_203650]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_LIABILITIES].[PRC/2_203650].[PRC/" & Account & "]"
            End If
            
        'Other Liabilities
        Case 204000 To 204999
            If AccountInt = 204000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_LIABILITIES].[PRC/2_204000]"
            ElseIf AccountInt >= 204610 And AccountInt <= 204650 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_LIABILITIES].[PRC/2_204000].[PRC/2_204600].[PRC/" & Account & "]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_LIABILITIES].[PRC/2_204000].[PRC/" & Account & "]"
            End If
        
        'Long Term Liabilities
        Case 205100 To 205500
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_LIABILITIES AND NET ASSE].[PRC/2_TOTAL LIABILITIES].[PRC/2_LONG-TERM LIABILITIES].[PRC/" & Account & "]"
        
        'Revenues - Video Lottery
        Case 400000 To 400190
            If AccountInt = 400000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_400000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_400000].[PRC/" & Account & "]"
            End If
        'Revenues - Slot
        Case 409670 To 410970
            If AccountInt = 409670 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_409670]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_409670].[PRC/" & Account & "]"
            End If
            
        'Revenues - Table Games
        Case 411000 To 411920
            If AccountInt = 411000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_411000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_411000].[PRC/" & Account & "]"
            End If
                
        'Revenues - Poker
        Case 412000 To 412500
            If AccountInt = 412000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_412000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_412000].[PRC/" & Account & "]"
            End If
            
        'Revenues - Class II
        Case 413000 To 416310
            If AccountInt = 413000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_413000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_413000].[PRC/" & Account & "]"
            End If
        
        'Revenues - Non-Gaming
        Case 420000 To 459530
            If AccountInt = 420000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_420000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_420000].[PRC/" & Account & "]"
            End If
        
        'Revenues - Showroom
        Case 460000 To 479100
            If AccountInt = 460000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_460000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_460000].[PRC/" & Account & "]"
            End If
        
        'Revenues - Other
        Case 480000 To 489000
            If AccountInt = 480000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_480000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_480000].[PRC/" & Account & "]"
            End If
        
        'Revenues - Non-Operating
        Case 490000 To 499990
            If AccountInt = 490000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_490000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES].[PRC/2_490000].[PRC/" & Account & "]"
            End If
            
        'Expenses - COGS
        Case 500000 To 570000
            If AccountInt = 500000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_500000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_500000].[PRC/" & Account & "]"
            End If
        
        'Expenses - Comps/Coupons/Discounts
        Case 600000 To 604220
            
            'Expenses - Comps
            If AccountInt = 600000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_600000]"
            
            ElseIf AccountInt = 601000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_600000].[PRC/2_601000]"
            
            ElseIf AccountInt >= 601020 And AccountInt <= 601300 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_600000].[PRC/2_601000].[PRC/" & Account & "]"

            'Expenses - Coupons
            ElseIf AccountInt = 602000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_600000].[PRC/2_602000]"
            
            ElseIf AccountInt >= 602020 And AccountInt <= 602100 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_600000].[PRC/2_602000].[PRC/" & Account & "]"
            
            'Expenses - Manager Adjustments
            ElseIf AccountInt = 603000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_600000].[PRC/2_603000]"

            ElseIf AccountInt >= 603020 And AccountInt <= 603380 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_600000].[PRC/2_603000].[PRC/" & Account & "]"
            
            'Expenses - Discounts
            ElseIf AccountInt = 604000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_600000].[PRC/2_604000]"        
            ElseIf AccountInt >= 604020 And AccountInt <= 604220 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_600000].[PRC/2_604000].[PRC/" & Account & "]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_600000].[PRC/" & Account & "]"
            End If
        
        'Expenses - Salaries and Wages
        Case 710000 To 719900
            If AccountInt = 710000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_710000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_710000].[PRC/" & Account & "]"
            End If
        
        'Expenses - PTO
        Case 720000 To 739150
            If AccountInt = 720000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_720000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_720000].[PRC/" & Account & "]"
            End If
        
        'Expenses - Payroll Taxes
        Case 760000 To 770000
            If AccountInt = 760000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_760000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_760000].[PRC/" & Account & "]"
            End If
        
        'Expenses - Taxes
        Case 800000 To 809000
            If AccountInt = 800000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_800000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_800000].[PRC/" & Account & "]"
            End If
                
        'Expenses - Supplies
        Case 810000 To 819000
            If AccountInt = 810000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_810000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_810000].[PRC/" & Account & "]"
            End If
            
        'Expenses - Repairs & Maintenance/Service Contracts
        Case 830000 To 835100
            If AccountInt = 830000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_830000]"
            ElseIf AccountInt >= 835000 And AccountInt <= 835100 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_830000].[PRC/2_835000].[PRC/" & Account & "]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_830000].[PRC/" & Account & "]"
            End If
            
        
        'Expenses - Rental/Lease
        Case 835500 To 836000
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_835500].[PRC/" & Account & "]"
        
        'Expenses - Professional Services
        Case 840000 To 849500
            If AccountInt = 840000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_840000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_840000].[PRC/" & Account & "]"
            End If
        
        'Expenses - Marketing
        Case 850000 To 859100
            If AccountInt = 850000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_850000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_850000].[PRC/" & Account & "]"
            End If
        
        'Expenses - Communications
        Case 860000 To 866000
            If AccountInt = 860000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_860000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_860000].[PRC/" & Account & "]"
            End If
            
        'Expenses - Utilities
        Case 870000 To 875000
            If AccountInt = 870000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_870000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_870000].[PRC/" & Account & "]"
            End If
        
        'Expenses - Insurance
        Case 880000 To 885000
            If AccountInt = 880000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_880000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_880000].[PRC/" & Account & "]"
            End If
        'Expenses - Depreciation
        Case 890000 To 895000
            If AccountInt = 890000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_890000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_890000].[PRC/" & Account & "]"
            End If
            
        'Expenses - Other
        Case 950000 To 959900
            If AccountInt = 950000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_950000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_950000].[PRC/" & Account & "]"
            End If
            
        'Expenses - Capital Management
        Case 960000 To 961000
            If AccountInt = 960000 Then
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_960000]"
            Else
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_960000].[PRC/" & Account & "]"
            End If
        'Stat accounts
        Case 991101 To 999025
                RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_STAT].[PRC/" & Account & "]"
                
    End Select


    'this case statement is for header accounts that do not follow the same pattern as the others
    Select Case AccountInt
        
        'Slots liability header account.
        Case 20305
            RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_TOTAL ASSETS].[PRC/2_107000].[PRC/2_203050]"
        
        'Service Contracts header account
        Case 83500
            RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_830000].[PRC/2_835000]"
        
        'Rental/Lease header account
        Case 83550
            RootFormula = "[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_EXPENSES].[PRC/2_835500]"

    End Select
    'this case statement is for the top level accounts (REVENUES, COGS, COMPS, etc.)
    Select Case Account
        Case "Revenues"
            RootFormula ="[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME].[PRC/2_REVENUES]"
        
        Case "Net Income"
            RootFormula ="[GLTOT_CHARTACCOUNT].[PRC].[PRC/2_TOP_NODE].[PRC/2_NET INCOME]"

    End Select

    ChartAccount = RootFormula

End Function


Function ChartAccounts(Accounts As String) As String

    Dim splitAccounts() As String
    Dim AccountGrouping As String
    Dim i As Long

    AccountGrouping = ""

    If Instr(Accounts,",") = 0 Then
    ChartAccounts = ChartAccount(Accounts)
    Else
        splitAccounts = Split(Accounts, ",")
        For i = LBound(splitAccounts, 1) To UBound(splitAccounts, 1)
            AccountGrouping = AccountGrouping & "," & ChartAccount(splitAccounts(i))
        Next i
        AccountGrouping = "{" & Replace(AccountGrouping, ",", "", , 1) & "}"
        ChartAccounts = AccountGrouping
    End If

    'Debug.Print (ChartAccounts)

End Function


Function Department(deptNumber As String) As String
    On Error Resume Next
    Dim deptNumberInt As Long
    Dim rootFormula, deptNumberString As String

    deptNumberString = Format(deptNumber, "0000")
    deptNumberInt = Int(deptNumber)
    rootFormula = "[GLTOT_FINANCEDIMENSION1].[PRC].[PRC/2_TOP_NODE]"


    'to group by department
    Select Case deptNumberInt

        'Gaming
        Case 100 To 190
            Department = rootFormula & ".[PRC/2_GAMING].[PRC/" & deptNumberString & "]"
    
        'Golf
        Case 250 To 290
            Department = rootFormula & ".[PRC/2_GOLF].[PRC/" & deptNumberString & "]"
        
        'Hotel
        Case 200 To 240
            Department = rootFormula & ".[PRC/2_HOTEL].[PRC/" & deptNumberString & "]"
    
        'Food 
        Case 300 To 390
            
            'Player Food
            If deptNumber = 340 Or deptNumber = 373 Or deptNumber = 375 Then
                Department = rootFormula & ".[PRC/2_FOOD AND BEV].[PRC/2_FOOD].[PRC/2_PLAYER FOOD].[PRC/" & deptNumberString & "]"
            
            'Team Member Dining
            ElseIf deptNumber = 380 Then
                Department = rootFormula & ".[PRC/2_FOOD AND BEV].[PRC/2_FOOD].[PRC/2_TEAM MEMBER].[PRC/" & deptNumberString & "]"
            
            'Guest Food
            Else
            Department = rootFormula & ".[PRC/2_FOOD AND BEV].[PRC/2_FOOD].[PRC/2_GUEST FOOD].[PRC/" & deptNumberString & "]"
        End If

        'Beverage
        Case 400 To 490
            Department = rootFormula & ".[PRC/2_FOOD AND BEV].[PRC/2_BEVERAGE].[PRC/" & deptNumberString & "]"
    
        'Retail
        Case 500 To 501
            Department = rootFormula & ".[PRC/2_RETAIL].[PRC/" & deptNumberString & "]"
        
        'Entertainment
        Case 600 To 601
            Department = rootFormula & ".[PRC/2_ENTERTAINMENT].[PRC/" & deptNumberString & "]"
    
        'Other
        Case 900 To 991
            Department = rootFormula & ".[PRC/2_OTHER].[PRC/" & deptNumberString & "]"

        'Preopening
        Case 995 To 999
            Department = rootFormula & ".[PRC/2_PRE OPENING].[PRC/" & deptNumberString & "]"

        'Marketing
        Case 1100 To 1170
            Department = rootFormula & ".[PRC/2_MARKETING].[PRC/" & deptNumberString & "]"

        'Facilities
        Case 1220 To 1290
            Department = rootFormula & ".[PRC/2_FACILITIES].[PRC/" & deptNumberString & "]"
        
        'DPS (Security)
        Case 1300
            Department = rootFormula & ".[PRC/2_SECURITY].[PRC/" & deptNumberString & "]"

        'General & Administrative
        Case 1500, 1530, 1531, 1535, 1540
            Department = rootFormula & ".[PRC/2_GENERAL AND A].[PRC/" & deptNumberString & "]"
        
        'Human Resources
        Case 1510, 1511, 1512, 1513, 1514, 1515
            Department = rootFormula & ".[PRC/2_HUMAN RESOURC].[PRC/" & deptNumberString & "]"
        
        'Information Technology
        Case 1520, 1521, 1550
            Department = rootFormula & ".[PRC/2_IT].[PRC/" & deptNumberString & "]"
    
        'Capital Mangement
        Case 1910 To 1959
            Department = rootFormula & ".[PRC/2_CAPITAL MANAG].[PRC/" & deptNumberString & "]"
        Case Else
            'to group by division and sub-divisions
            Select Case deptNumber
                Case "All"
                    Department = rootFormula
                Case "Gaming"
                    Department = rootFormula & ".[PRC/2_GAMING]"
                Case "Hotel"
                    Department = rootFormula & ".[PRC/2_GOLF]," _ 
                                & rootFormula & ".[PRC/2_HOTEL]"
                Case "F&B"
                    Department = rootFormula & ".[PRC/2_FOOD AND BEV]"   
                Case "Food"
                    Department = rootFormula & ".[PRC/2_FOOD AND BEV].[PRC/2_FOOD]"
                Case "Guest Food"
                    Department = rootFormula & ".[PRC/2_FOOD AND BEV].[PRC/2_FOOD].[PRC/2_GUEST FOOD]"
                Case "Player Food"
                    Department = rootFormula & ".[PRC/2_FOOD AND BEV].[PRC/2_FOOD].[PRC/2_PLAYER FOOD]"

                Case "Beverage"
                    Department = rootFormula & ".[PRC/2_FOOD AND BEV].[PRC/2_BEVERAGE]"
                Case "Retail"
                    Department = rootFormula & ".[PRC/2_RETAIL]"
                Case "Entertainment"
                    Department = rootFormula & ".[PRC/2_ENTERTAINMENT]"
                Case "Other"
                    Department = rootFormula & ".[PRC/2_OTHER]"
                Case "Facilities"
                    Department = rootFormula & ".[PRC/2_FACILITIES]"

                Case "Marketing"
                    Department = rootFormula & ".[PRC/2_MARKETING]"
                Case "DPS"
                    Department = rootFormula & ".[PRC/2_SECURITY]"
                Case "G&A"
                    Department = rootFormula & ".[PRC/2_GENERAL AND A]," _ 
                                & rootFormula & ".[PRC/2_HUMAN RESOURC]," _ 
                                & rootFormula & ".[PRC/2_IT]"
                Case "HR"
                    Department =  rootFormula & ".[PRC/2_HUMAN RESOURC]" 
                Case "IT"
                    Department = rootFormula & ".[PRC/2_IT]"
                Case "Capital Management"
                    Department = rootFormula & ".[PRC/2_CAPITAL MANAG]"
                Case Else
                    Department = "Department doesn't exist!"
            End Select
    End Select

End Function


Function Departments(deptNumbers As String) As String

    'Use this function when needing to include multiple departments into a calculation
    'Deliminate each department with a comma (e.g 0100, 0110)

    Dim splitDepartments() As String
    Dim DepartmentGrouping, DepartmentString As String
    Dim i As Long

    DepartmentGrouping = ""
    DepartmentString = Department(deptNumbers)

    If InStr(deptNumbers, ",") <> 0 Or InStr(DepartmentString, ",") <> 0 Then
        splitDepartments = Split(deptNumbers, ",")
        For i = LBound(splitDepartments, 1) To UBound(splitDepartments, 1)
            DepartmentGrouping = DepartmentGrouping & "," & Department(splitDepartments(i))
        Next i
        DepartmentGrouping = "{" & Replace(DepartmentGrouping, ",", "", , 1) & "}"
        Departments = DepartmentGrouping
        
    Else
        Departments = Department(deptNumbers)
    End If

    'Debug.Print (Departments)

End Function
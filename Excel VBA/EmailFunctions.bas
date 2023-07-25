Attribute VB_Name = "EmailFunctions"
Option Explicit

Function TaskKill(sTaskName)
    TaskKill = CreateObject("WScript.Shell").Run("taskkill /f /im " & sTaskName, 0, True)
End Function


Function UnFvEbita(GamingTotals, NonGamingTotals, MTDPayrollTotals, MTDEbitaTotals, MTDCompVariance As Variant) As String

If MTDEbitaTotals > 0 And GamingTotals > 0 And NonGamingTotals > 0 Then
    UnFvEbita = "Favorable MTD EBITDA was due to Gaming (" & RoundTo(GamingTotals) & ") and Non-Gaming (" & RoundTo(NonGamingTotals) & ")"

        ElseIf MTDEbitaTotals > 0 And GamingTotals > 0 And NonGamingTotals <= 0 Then
                UnFvEbita = "Favorable MTD EBITDA was due to Gaming (" & RoundTo(GamingTotals) & ")"
            
        ElseIf MTDEbitaTotals > 0 And GamingTotals <= 0 And NonGamingTotals > 0 Then
                UnFvEbita = "Favorable MTD EBITDA was due to Non-Gaming (" & RoundTo(NonGamingTotals) & ")"
                  
ElseIf MTDEbitaTotals < 0 And GamingTotals < 0 And NonGamingTotals < 0 Then
            UnFvEbita = "Unfavorable MTD EBITDA was due to Gaming (" & RoundTo(GamingTotals) & ") and Non-Gaming (" & RoundTo(NonGamingTotals) & ")"
        
        ElseIf MTDEbitaTotals < 0 And GamingTotals < 0 And NonGamingTotals > 0 Then
            UnFvEbita = "Unfavorable MTD EBITDA was due to Gaming (" & RoundTo(GamingTotals) & ")"
        
        ElseIf MTDEbitaTotals < 0 And GamingTotals > 0 And NonGamingTotals < 0 Then
            UnFvEbita = "Unfavorable MTD EBITDA was due to Non-Gaming (" & RoundTo(NonGamingTotals) & ") and total payroll expenses (" & RoundTo(MTDPayrollTotals) & " vs budget)"
        
        ElseIf MTDEbitaTotals < 0 And GamingTotals > 0 And NonGamingTotals > 0 Then
            UnFvEbita = "Unfavorable MTD EBITDA was due to total comp expense (" & RoundTo(MTDCompVariance) & " vs budget) and total payroll expenses (" & RoundTo(MTDPayrollTotals) & " vs budget)"
Else
    UnFvEbita = "Check DOR or DOR Variance Analysis for reasons behind Unfavorable/Favorable EBITDA"
    
End If

End Function


Function RoundTo(Value As Variant) As String

Select Case Value
    Case Is <= -95000
        RoundTo = "-$" & Round(Value / 1000000, 1) * -1 & "m"
    Case Is <= -949
        RoundTo = "-$" & Round(Value / 1000, 0) * -1 & "k"
    Case Is <= 0
        RoundTo = "Flat with budget" '"-$" & Round(Value, 0) * -1
    Case Is <= 949
        RoundTo = "$" & Round(Value, 0)
    Case Is >= 95500
        RoundTo = "+$" & Round(Value / 1000000, 1) & "m"
    Case Is >= 950
        RoundTo = "+$" & Round(Value / 1000, 0) & "k"
    
    End Select
    
End Function

Function RoundToNetSlots(Value As Variant) As String

Select Case Value
    Case Is <= -95000
        RoundToNetSlots = "$" & Round(Value / 1000000, 0) * -1 & "m"
    Case Is <= -949
        RoundToNetSlots = "$" & Round(Value / 1000, 0) * -1 & "k"
    Case Is <= 0
        RoundToNetSlots = "Flat with budget" '"$" & Round(Value / 1000, 0) * -1 & "k"
    Case Is <= 949
        RoundToNetSlots = "$" & Round(Value, 0)
    Case Is >= 95500
        RoundToNetSlots = "$" & Round(Value / 1000000, 0) & "m"
    Case Is >= 950
        RoundToNetSlots = "$" & Round(Value / 1000, 0) & "k"
    
    End Select
    
End Function
Function NetSlotsEmail(MTDSlotsActualRevenue, MTDSlotsBudgetRevenue, MTDActualCoinIn, _
                        MTDBudgetCoinIn, NetSlots, ChangeinCoinIn, ChangeinSlotHold, _
                        MTDBudgetSlotHold As Variant) As String

'calculations for when net slots is positive
If NetSlots > 0 And (MTDActualCoinIn / MTDBudgetCoinIn) - 1 > 0 And (MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn) > 0 Then
    NetSlotsEmail = "Net Slots (" & RoundTo(NetSlots) & ", due to Coin in (" & RoundTo(ChangeinCoinIn) & ") and Hold (" & RoundTo(ChangeinSlotHold) & ")); " & RoundToNetSlots(MTDActualCoinIn) & " Coin in is +" & Format((MTDActualCoinIn / MTDBudgetCoinIn) - 1, "Percent") & " vs budget of " & RoundToNetSlots(MTDBudgetCoinIn) & ";" _
                    & " Hold percent of " & Format((MTDSlotsActualRevenue / MTDActualCoinIn), "Percent") & " is + " & Round((((MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn)) * 100), 2) _
                        & " ppt vs budget of " & Format(MTDBudgetSlotHold, "Percent")
    ElseIf NetSlots > 0 And (MTDActualCoinIn / MTDBudgetCoinIn) - 1 > 0 And (MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn) < 0 Then
    NetSlotsEmail = "Net Slots (" & RoundTo(NetSlots) & ", due to Coin in); " & RoundToNetSlots(MTDActualCoinIn) & " Coin in is +" & Format((MTDActualCoinIn / MTDBudgetCoinIn) - 1, "Percent") & " vs budget of " & RoundToNetSlots(MTDBudgetCoinIn)
    
    ElseIf NetSlots > 0 And (MTDActualCoinIn / MTDBudgetCoinIn) - 1 < 0 And (MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn) > 0 Then
    NetSlotsEmail = "Net Slots (" & RoundTo(NetSlots) & ", due to Hold); Hold percent of " & Format((MTDSlotsActualRevenue / MTDActualCoinIn), "Percent") & " is + " & Round((((MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn)) * 100), 2) _
                        & " ppt vs budget of " & Format(MTDBudgetSlotHold, "Percent")
'calculations for when net slots is negative
ElseIf NetSlots < 0 And (MTDActualCoinIn / MTDBudgetCoinIn) - 1 < 0 And (MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn) < 0 Then
    NetSlotsEmail = "Net Slots (" & RoundTo(NetSlots) & ", due to Coin in (" & RoundTo(ChangeinCoinIn) & ") and Hold (" & RoundTo(ChangeinSlotHold) & ")); " & RoundToNetSlots(MTDActualCoinIn) & " Coin in is " & Format((MTDActualCoinIn / MTDBudgetCoinIn) - 1, "Percent") & " vs budget of " & RoundToNetSlots(MTDBudgetCoinIn) & ";" _
                    & " Hold percent of " & Format((MTDSlotsActualRevenue / MTDActualCoinIn), "Percent") & " is " & Format((MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn), "Percent") _
                    & " ppt vs budget of " & Format(MTDBudgetSlotHold, "Percent")

    ElseIf NetSlots < 0 And (MTDActualCoinIn / MTDBudgetCoinIn) - 1 > 0 And (MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn) < 0 Then
    NetSlotsEmail = "Net Slots (" & RoundTo(NetSlots) & ", due to Hold); Hold percent of " & Format((MTDSlotsActualRevenue / MTDActualCoinIn), "Percent") & " is " & Round((((MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn)) * 100), 2) _
                        & " ppt vs budget of " & Format(MTDBudgetSlotHold, "Percent")
    ElseIf NetSlots < 0 And (MTDActualCoinIn / MTDBudgetCoinIn) - 1 < 0 And (MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn) > 0 Then
    NetSlotsEmail = "Net Slots (" & RoundTo(NetSlots) & ", due to Coin in); " & RoundToNetSlots(MTDActualCoinIn) & " Coin in is " & Format((MTDActualCoinIn / MTDBudgetCoinIn) - 1, "Percent") & " vs budget of " & RoundToNetSlots(MTDBudgetCoinIn)
    

'calculations for when net slots is positive due to EZ play
ElseIf NetSlots > 0 And (MTDActualCoinIn / MTDBudgetCoinIn) - 1 < 0 And (MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn) < 0 Then
    NetSlotsEmail = "Net Slots (" & RoundTo(NetSlots) & "); " & RoundToNetSlots(MTDActualCoinIn) & " Coin in is " & Format((MTDActualCoinIn / MTDBudgetCoinIn) - 1, "Percent") & " vs budget of " & RoundToNetSlots(MTDBudgetCoinIn) & ";" _
                    & " Hold percent of " & Format((MTDSlotsActualRevenue / MTDActualCoinIn), "Percent") & " is " & Format((MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn), "Percent") _
                    & " ppt vs budget of " & Format(MTDBudgetSlotHold, "Percent") & ")"
                    
    ElseIf NetSlots > 0 And (MTDActualCoinIn / MTDBudgetCoinIn) - 1 > 0 And (MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn) < 0 Then
    NetSlotsEmail = "Net Slots (" & RoundTo(NetSlots) & "); " & RoundToNetSlots(MTDActualCoinIn) & " Coin in is " & Format((MTDActualCoinIn / MTDBudgetCoinIn) - 1, "Percent") & " vs budget of " & RoundToNetSlots(MTDBudgetCoinIn) & ";" _
                    & " Hold percent of " & Format((MTDSlotsActualRevenue / MTDActualCoinIn), "Percent") & " is + " & Format((MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn), "Percent") _
                    & " ppt vs budget of " & Format(MTDBudgetSlotHold, "Percent")
    ElseIf NetSlots > 0 And (MTDActualCoinIn / MTDBudgetCoinIn) - 1 < 0 And (MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn) > 0 Then
    NetSlotsEmail = "Net Slots (" & RoundTo(NetSlots) & "); " & RoundToNetSlots(MTDActualCoinIn) & " Coin in is +" & Format((MTDActualCoinIn / MTDBudgetCoinIn) - 1, "Percent") & " vs budget of " & RoundToNetSlots(MTDBudgetCoinIn) & ";" _
                    & " Hold percent of " & Format((MTDSlotsActualRevenue / MTDActualCoinIn), "Percent") & " is " & Format((MTDSlotsActualRevenue / MTDActualCoinIn) - (MTDSlotsBudgetRevenue / MTDBudgetCoinIn), "Percent") _
                    & " ppt vs budget of " & Format(MTDBudgetSlotHold, "Percent")
Else
    NetSlotsEmail = "Check DOR or DOR Variance Analysis to determine net slots variance"
    
    End If
End Function

Function NetTableEmail(MTDTableActualRevenue, MTDTableBudgetRevenue, MTDActualDrop, MTDBudgetDrop, NetTable, _
                       ChangeinDrop, ChangeinTableHold, MTDBudgetTableHold As Variant) As String

If NetTable > 0 And (MTDActualDrop / MTDBudgetDrop) - 1 > 0 And (MTDTableActualRevenue / MTDActualDrop) - (MTDTableBudgetRevenue / MTDBudgetDrop) > 0 Then
    NetTableEmail = "Net Table (" & RoundTo(NetTable) & ", due to Drop (" & RoundTo(ChangeinDrop) & ") and Hold (" & RoundTo(ChangeinTableHold) & ")); " & RoundToNetSlots(MTDActualDrop) & " Drop is +" & Format((MTDActualDrop / MTDBudgetDrop) - 1, "Percent") & " vs budget of " & RoundToNetSlots(MTDBudgetDrop) & ";" _
                    & " Hold percent of " & Format((MTDTableActualRevenue / MTDActualDrop), "Percent") & " is + " & Round((((MTDTableActualRevenue / MTDActualDrop) - (MTDTableBudgetRevenue / MTDBudgetDrop)) * 100), 2) _
                    & " ppt vs budget of " & Format(MTDBudgetTableHold, "Percent")
    ElseIf NetTable > 0 And (MTDActualDrop / MTDBudgetDrop) - 1 > 0 And (MTDTableActualRevenue / MTDActualDrop) - (MTDTableBudgetRevenue / MTDBudgetDrop) < 0 Then
    NetTableEmail = "Net Table (" & RoundTo(NetTable) & ", due to Drop); " & RoundToNetSlots(MTDActualDrop) & " Drop is +" & Format((MTDActualDrop / MTDBudgetDrop) - 1, "Percent") & " vs budget"
    
    ElseIf NetTable > 0 And (MTDActualDrop / MTDBudgetDrop) - 1 < 0 And (MTDTableActualRevenue / MTDActualDrop) - (MTDTableBudgetRevenue / MTDBudgetDrop) > 0 Then
    NetTableEmail = "Net Table (" & RoundTo(NetTable) & ", due to Hold); Hold percent of " & Format((MTDTableActualRevenue / MTDActualDrop), "Percent") & " is + " & Round((((MTDTableActualRevenue / MTDActualDrop) - (MTDTableBudgetRevenue / MTDBudgetDrop)) * 100), 2) _
                    & " ppt vs budget of " & Format(MTDBudgetTableHold, "Percent")
                    
ElseIf NetTable < 0 And (MTDActualDrop / MTDBudgetDrop) - 1 < 0 And (MTDTableActualRevenue / MTDActualDrop) - (MTDTableBudgetRevenue / MTDBudgetDrop) < 0 Then
    NetTableEmail = "Net Table (" & RoundTo(NetTable) & ", due to Drop (" & RoundTo(ChangeinDrop) & ") and Hold (" & RoundTo(ChangeinTableHold) & "); " & RoundToNetSlots(MTDActualDrop) & " Drop is " & Format((MTDActualDrop / MTDBudgetDrop) - 1, "Percent") & " vs budget of " & RoundToNetSlots(MTDBudgetDrop) & ";" _
                    & " Hold percent of " & Format((MTDTableActualRevenue / MTDActualDrop), "Percent") & " is " & Round((((MTDTableActualRevenue / MTDActualDrop) - (MTDTableBudgetRevenue / MTDBudgetDrop)) * 100), 2) _
                    & " ppt vs budget of " & Format(MTDBudgetTableHold, "Percent")

    ElseIf NetTable < 0 And (MTDActualDrop / MTDBudgetDrop) - 1 > 0 And (MTDTableActualRevenue / MTDActualDrop) - (MTDTableBudgetRevenue / MTDBudgetDrop) < 0 Then
        NetTableEmail = "Net Table (" & RoundTo(NetTable) & ", due to Hold); Hold percent of " & Format((MTDTableActualRevenue / MTDActualDrop), "Percent") & " is " & Round((((MTDTableActualRevenue / MTDActualDrop) - (MTDTableBudgetRevenue / MTDBudgetDrop)) * 100), 2) _
                        & " ppt vs budget of " & Format(MTDBudgetTableHold, "Percent")
    ElseIf NetTable < 0 And (MTDActualDrop / MTDBudgetDrop) - 1 < 0 And (MTDTableActualRevenue / MTDActualDrop) - (MTDTableBudgetRevenue / MTDBudgetDrop) > 0 Then
    NetTableEmail = "Net Table (" & RoundTo(NetTable) & ", due to Drop); " & RoundToNetSlots(MTDActualDrop) & " Drop is " & Format((MTDActualDrop / MTDBudgetDrop) - 1, "Percent") & " vs budget"
    

Else
    NetTableEmail = "Check DOR or DOR Variance Analysis to determine net tables variance"
    
    End If
End Function


Function HotelFoodRetail_Email(MTD_Hotel, MTD_Food, MTD_Retail As Variant) As String

HotelFoodRetail_Email = "Hotel (" & RoundTo(MTD_Hotel) & "), Food (" & RoundTo(MTD_Food) & "), Retail (" & RoundTo(MTD_Retail) & ")"

End Function
Function RoundToEBITA(Value As Variant) As String

Select Case Value
    Case Is <= -95000
        RoundToEBITA = "$" & Round(Value / 1000000, 1) * -1 & "m"
    Case Is <= -949
        RoundToEBITA = "$" & Round(Value / 1000, 0) * -1 & "k"
    Case Is <= 0
        RoundToEBITA = "Flat with budget" '"$" & Round(Value / 1000, 0) * -1 & "k"
    Case Is <= 949
        RoundToEBITA = "$" & Round(Value, 0)
    Case Is >= 95500
        RoundToEBITA = "$" & Round(Value / 1000000, 1) & "m"
    Case Is >= 950
        RoundToEBITA = "$" & Round(Value / 1000, 0) & "k"
    
    End Select
    
End Function




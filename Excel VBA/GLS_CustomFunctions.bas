Attribute VB_Name = "GLS_CustomFunctions"
Option Explicit

Function VarString(Variance As Variant) As String

Dim vString As String
Dim absVariance As Double

absVariance = Abs(Variance)
Select Case True
    Case absVariance > 0 And absVariance < 1
        vString = Str(absVariance * 1000) & "k"
    Case absVariance > 1
        vString = Str(Round(absVariance, 1)) & "m"
End Select

If Variance < 0 Then
    vString = "-" & vString
End If
 
VarString = Trim(vString)

End Function

Function EventName(EventString As String, ParseType As String) As String

Dim EventArr() As String

EventArr = Split(EventString, "-")
Select Case ParseType
    Case "Year"
        EventName = EventArr(0)
    Case "Name"
        EventName = EventArr(4)
End Select


End Function

Function HTMLmarkup(FunctionType As String, _
                    InputString As Variant, _
                    Optional VarianceNumber As Variant = 0, _
                    Optional ReverseVariance As Boolean = False) As String

Dim ExceptionString As String

Select Case FunctionType
    Case "Variance"
        If VarianceNumber < 0 Then
            HTMLmarkup = "<span style=""color: red;"" >" & InputString & "</span>"
        ElseIf VarianceNumber < 0 And ReverseVariance = True Then
            HTMLmarkup = "<span style=""color: green;"" >" & InputString & "</span>"
        ElseIf VarianceNumber > 0 And ReverseVariance = True Then
            HTMLmarkup = "<span style=""color: red;"" >" & InputString & "</span>"
        Else
            HTMLmarkup = "<span style=""color: green;"" >" & InputString & "</span>"
        End If
    
    Case "Event Target"
        If InputString = "Target" Then
            HTMLmarkup = "This event met the estimated guest count"
        ElseIf InputString = "Exceeded" Then
            HTMLmarkup = "This event exceeded the estimated guest count"
        Else
            HTMLmarkup = "This event missed the estimated guest count"
        End If
    
    Case "Exceptions"
        If Int(InputString) > 0 Then
            HTMLmarkup = "With some exceptions, we saw lift across all ADT groups for Coin In/Drop."
        Else
            HTMLmarkup = "We saw lift across all ADT groups for Coin In/Drop."
        End If
    
    Case "ADT Exceptions"
        If VarianceNumber > 0 Then
            HTMLmarkup = "<ul><li>Increased " & InputString & " for all ADT groups with the exceptions listed below</li><ul>"
        Else
            HTMLmarkup = "<ul><li>Increased " & InputString & " for all ADT groups</li></ul>"
            
        End If
        
End Select
    

End Function
Function ExceptionsString(ExceptionList As Range, ADTgroup As Range) As String

Dim exc, Adt As Range
Dim excString As String
Dim i, count As Integer

i = 0
count = 0
excString = ""

For Each exc In ExceptionList
        count = count + 1
        If Not IsEmpty(exc.Value) Then
            i = i + 1
            Set Adt = ADTgroup.Cells(count, 1)
            excString = excString + "<li>" & Adt.Value & "</li>"
        End If
Next exc

If i <> 0 Then
    ExceptionsString = excString + "</ul></ul>"
Else
     ExceptionsString = excString
End If

End Function

Function ExceptionsCount(ExceptionList As Range) As Integer

Dim exc As Range
Dim i As Integer

i = 0

For Each exc In ExceptionList
    If Not IsEmpty(exc.Value) Then
        i = i + 1
    End If
Next exc

ExceptionsCount = i

End Function

Sub GLS_emailCreation()


Dim objOutlook, GLS_email As Object
Dim eList, eAddress, TestEmail, ExceptionsTable As Range
Dim CoinInExceptions As Range
Dim DropExceptions As Range
Dim nTwinExceptions As Range
Dim nAwinExceptions As Range
Dim ADTgroup As Range
Dim emailTo, PDFSavePath, eName, eTargetStatus, eMinResponse _
    , eMaxResponse, rTotal, rTotalnonEvent, ActiveTotalnonEvent, RP, RNP, NRNP, NRBP, RP_P, RNP_P, NRNP_P, NRBP_P, GR, EZR, nTwinPercent, nAwinPercent, gCoinIn_event, gCoinIn_nonEvent, gDrop_event, gDrop_nonEvent, gnTwin_event _
    , gnTwin_nonEvent, gnAwin_event, gnAwin_nonEvent, gCoinInVarPercent, gDropVarPercent, rgCoinIn_event, rgCoinIn_nonEvent, rgDrop_event, rgDrop_nonEvent, rgTotalCOGS_event _
    , rgTotalCOGS_nonEvent, rgnAwin_event, rgnAwin_nonEvent, rgCoinInVarPercent, rgDropVarPercent, rgTotalCOGSVarPercent, rgnAwinVarPercent, offeredTotal, activeTotal, ActivePercentOfOffered, ActiveLiftPercent, RedemptionLiftPercent As String
Dim gCoinIn_Var, gDrop_Var, gnTwin_Var, gnAwin_Var, rgCoinIn_Var, rgDrop_Var, rgTotalCOGS_Var, rgnAwin_Var, ActiveLift, RedemptionLift As Variant

'OutLook and Range variables
Set objOutlook = CreateObject("Outlook.Application")
Set eList = Worksheets("EmailRecipients").Range("EmailAddress")
Set GLS_email = objOutlook.CreateItem(0)
Set ExceptionsTable = Worksheets("Data").Range("TotalExceptions")
Set CoinInExceptions = Worksheets("Data").Range("CoinInExceptions")
Set DropExceptions = Worksheets("Data").Range("DropExceptions")
Set nTwinExceptions = Worksheets("Data").Range("nTwinExceptions")
Set nAwinExceptions = Worksheets("Data").Range("nAwinExceptions")
Set ADTgroup = Worksheets("Data").Range("ADTgroup")

'String variables
PDFSavePath = Worksheets("Data").Range("PDF_FileSavePath")
eName = Worksheets("Data").Range("eName")
eTargetStatus = Worksheets("Data").Range("eTargetStatus")
eMinResponse = Worksheets("Data").Range("eMinResponse")
eMaxResponse = Worksheets("Data").Range("eMaxResponse")
rTotal = Worksheets("Data").Range("rTotal")
rTotalnonEvent = Worksheets("Data").Range("rTotalnonEvent")
activeTotal = Worksheets("Data").Range("ActiveTotal")
ActiveTotalnonEvent = Worksheets("Data").Range("ActiveTotalnonEvent")
RP = Worksheets("Data").Range("RP")
RNP = Worksheets("Data").Range("RNP")
NRBP = Worksheets("Data").Range("NRBP")
NRNP = Worksheets("Data").Range("NRNP")
RP_P = Worksheets("Data").Range("RP_P")
RNP_P = Worksheets("Data").Range("RNP_P")
NRBP_P = Worksheets("Data").Range("NRBP_P")
NRNP_P = Worksheets("Data").Range("NRNP_P")
GR = Worksheets("Data").Range("GR")
EZR = Worksheets("Data").Range("EZR")
gCoinIn_event = Worksheets("Data").Range("gCoinIn_event")
gCoinIn_nonEvent = Worksheets("Data").Range("gCoinIn_nonEvent")
gDrop_event = Worksheets("Data").Range("gDrop_event")
gDrop_nonEvent = Worksheets("Data").Range("gDrop_nonEvent")
gnTwin_event = Worksheets("Data").Range("gnTwin_event")
gnTwin_nonEvent = Worksheets("Data").Range("gnTwin_nonEvent")
gnAwin_event = Worksheets("Data").Range("gnAwin_event")
gnAwin_nonEvent = Worksheets("Data").Range("gnAwin_nonEvent")
gCoinInVarPercent = Worksheets("Data").Range("gCoinInVarPercent")
gDropVarPercent = Worksheets("Data").Range("gDropVarPercent")
rgCoinInVarPercent = Worksheets("Data").Range("rgCoinInVarPercent")
rgDropVarPercent = Worksheets("Data").Range("rgDropVarPercent")
rgTotalCOGSVarPercent = Worksheets("Data").Range("rgTotalCOGSVarPercent")
rgnAwinVarPercent = Worksheets("Data").Range("rgnAwinVarPercent")
nTwinPercent = Worksheets("Data").Range("nTwinPercent")
nAwinPercent = Worksheets("Data").Range("nAwinPercent")
offeredTotal = Worksheets("Data").Range("OfferedTotal")
ActivePercentOfOffered = Worksheets("Data").Range("ActivePercentOfOffered")
ActiveLiftPercent = Worksheets("Data").Range("ActiveLiftPercent")
RedemptionLiftPercent = Worksheets("Data").Range("RedemptionLiftPercent")

'Variant variables
gCoinIn_Var = Worksheets("Data").Range("gCoinIn_Var")
gDrop_Var = Worksheets("Data").Range("gDrop_Var")
gnTwin_Var = Worksheets("Data").Range("gnTwin_Var")
gnAwin_Var = Worksheets("Data").Range("gnAwin_Var")
rgCoinIn_Var = Worksheets("Data").Range("rgCoinIn_Var")
rgDrop_Var = Worksheets("Data").Range("rgDrop_Var")
rgTotalCOGS_Var = Worksheets("Data").Range("rgTotalCOGS_Var")
rgnAwin_Var = Worksheets("Data").Range("rgnAwin_Var")
ActiveLift = Worksheets("Data").Range("ActiveLift")
RedemptionLift = Worksheets("Data").Range("RedemptionLift")

'to check if Outlook is open
Call TestOutlookIsOpen

'to loop through the email list on the "Email recipients" tab and add them to the recipient list
    For Each eAddress In eList
        emailTo = emailTo & ";" & eAddress.Value
    Next
    
'to create the base email
With GLS_email
            .Subject = "Gifting Lift Summary - " & eName
            .To = emailTo
            .Cc = "PRC_PlanningandAnalysis@pechanga.com"
            .Attachments.Add (PDFSavePath)
            .Display
            .HTMLBody = "<body style=font-size:11pt;font-family:Calibri Light><p>Hello,</p>" _
                    & "<p>Please see attached for the Gifting Lift Summary for the <b>" & eName & "</b> event." _
                    & "<br><br> " & HTMLmarkup("Event Target", eTargetStatus) & " (" & Format(eMinResponse, "#,###") & "-" & Format(eMaxResponse, "#,###") & " est. vs <b>" & Format(rTotal, "#,###") & "</b> actuals)" _
                    & " of which <b>" & Format(GR, "#,###") & "</b> players redeemed the gift and <b>" & Format(EZR, "#,###") & "</b> players redeemed EZ play." _
                    & " " & HTMLmarkup("Exceptions", ExceptionsCount(ExceptionsTable)) & " Total nTwin was up <b>" & Format(nTwinPercent, "0%") & "</b> with total nAwin being up <b>" & Format(nAwinPercent, "0%") & "</b>." & "<br><br>" _
                    & "Below are the event statistics: <br>" & "<ul> <li> Total Gaming </li>" _
                        & "<ul> <li> Coin In - $" & VarString(gCoinIn_event) & " vs $" & VarString(gCoinIn_nonEvent) & " Non-Event Date, " & HTMLmarkup("Variance", "$" & VarString(gCoinIn_Var), gCoinIn_Var) & " <b>" & HTMLmarkup("Variance", "(" & Format(gCoinInVarPercent, "0%") & ")", gCoinIn_Var) & "</b> lift" _
                        & "<li>Table Drop - $" & VarString(gDrop_event) & " vs $" & VarString(gDrop_nonEvent) & " Non-Event Date, " & HTMLmarkup("Variance", "$" & VarString(gDrop_Var), gDrop_Var) & " <b>" & HTMLmarkup("Variance", "(" & Format(gDropVarPercent, "0%") & ")", gDrop_Var) & "</b> lift" _
                        & "<li> nTwin - $" & VarString(gnTwin_event) & " vs $" & VarString(gnTwin_nonEvent) & " Non-Event Date, " & HTMLmarkup("Variance", "$" & VarString(gnTwin_Var), gnTwin_Var) & " <b>" & HTMLmarkup("Variance", "(" & Format(nTwinPercent, "0%") & ")", gnTwin_Var) & "</b> lift" _
                        & "<li> nAwin - $" & VarString(gnAwin_event) & " vs $" & VarString(gnAwin_nonEvent) & " Non-Event Date, " & HTMLmarkup("Variance", "$" & VarString(gnAwin_Var), gnAwin_Var) & " <b>" & HTMLmarkup("Variance", "(" & Format(nAwinPercent, "0%") & ")", gnAwin_Var) & "</b> lift" & "</ul></ul>" _
                    & "<ul> <li> Event Gaming </li>" _
                        & "<ul> <li> Redemption Stats </li>" _
                            & "<ul> <li> Offered - " & Format(offeredTotal, "#,###") & "</li>" & "<li> Active*  - " & Format(activeTotal, "#,###") & " (<b>" & Format(ActivePercentOfOffered, "0.0%") & "</b> of offered group), " & HTMLmarkup("Variance", Format(ActiveLift, "#,###"), ActiveLift) & "<b> " & HTMLmarkup("Variance", "(" & Format(ActiveLiftPercent, "0%") & ")", ActiveLift) & "</b> lift vs offered group Non-Event Date</li>" _
                                & "<ul> <li> Redemption with Play - " & Format(RP, "#,###") & "<b> (" & Format(RP_P, "0%") & ")</b> </li>" & "<li> Redemption No Play - " & Format(RNP, "#,###") & "<b> (" & Format(RNP_P, "0%") & ")</b> </li>" _
                                & "<li> Play with No Redemption - " & Format(NRBP, "#,###") & "<b> (" & Format(NRBP_P, "0%") & ")</b> </li>" & "<li> No Play No Redemption - " & Format(NRNP, "#,###") & "<b> (" & Format(NRNP_P, "0%") & ")</b> </li>" & "</ul></ul></ul>" _
                        & "<ul> <li> Gaming Stats" _
                            & "<ul> <li> Coin In - " & HTMLmarkup("Variance", "$" & VarString(rgCoinIn_Var), rgCoinIn_Var) & "<b> " & HTMLmarkup("Variance", "(" & Format(rgCoinInVarPercent, "0%") & ")", rgCoinIn_Var) & "</b> lift vs Non-Event Date for offered group </li>" & "<li> Table Drop - " & HTMLmarkup("Variance", "$" & VarString(rgDrop_Var), rgDrop_Var) & "<b> " & HTMLmarkup("Variance", "(" & Format(rgDropVarPercent, "0%") & ")", rgDrop_Var) & "</b> lift</b> vs Non-Event Date for offered group </li>" _
                            & "<li> nAwin - " & HTMLmarkup("Variance", "$" & VarString(rgnAwin_Var), rgnAwin_Var) & "<b> " & HTMLmarkup("Variance", "(" & Format(rgnAwinVarPercent, "0%") & ")", rgnAwin_Var) & "</b> lift</b> vs Non-Event Date for offered group </li>" & "<li> Total COGS* - " & HTMLmarkup("Variance", "$" & VarString(rgTotalCOGS_Var), rgTotalCOGS_Var, True) & "<b> " & HTMLmarkup("Variance", "(" & Format(rgTotalCOGSVarPercent, "0%") & ")", rgTotalCOGS_Var, True) & "</b> lift</b> vs Non-Event Date for offered group </li></ul></ul>" _
                        & "<ul> <li> ADT Stats" _
                            & "<ul> <li> Coin In" & HTMLmarkup("ADT Exceptions", "Coin In", ExceptionsCount(CoinInExceptions)) & ExceptionsString(CoinInExceptions, ADTgroup) & "</ul><ul><li> Drop" & HTMLmarkup("ADT Exceptions", "Drop", ExceptionsCount(DropExceptions)) & ExceptionsString(DropExceptions, ADTgroup) & "</ul><ul><li> Total nTheo" & HTMLmarkup("ADT Exceptions", "nTwin", ExceptionsCount(nTwinExceptions)) & ExceptionsString(nTwinExceptions, ADTgroup) & "</ul><ul><li> Total nAwin" & HTMLmarkup("ADT Exceptions", "nAwin", ExceptionsCount(nAwinExceptions)) & ExceptionsString(nAwinExceptions, ADTgroup) & "</ul></ul></ul>" _
                        & "<ul> <li> Redemption Group <ul> <li>" & Format(rTotal, "#,###") & " Active vs " & Format(rTotalnonEvent, "#,###") & " on Non-Event Dates increase of " & HTMLmarkup("Variance", Format(RedemptionLift, "#,###"), RedemptionLift) & "<b> " & HTMLmarkup("Variance", "(" & Format(RedemptionLiftPercent, "0%") & ")", RedemptionLift) & "</b> of group </ul></ul></ul> <br><br>" _
                        & "*Active - Invited for earned day of, and was on property day of event <br> *Total COGS - All marketing COGS used by the player day of event" _
                    & "</body>"

          End With
End Sub

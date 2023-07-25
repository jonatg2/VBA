Attribute VB_Name = "CreateDOREmails"
Option Explicit

Sub DailyDOREmail()

Dim DORDailyEmail As Object
Dim objOutlook As Object
Dim DORDate As Variant
Dim DORFilePath As Variant
Dim lookupList, lookupValue As Range
Dim eList, eAddress, TestEmail As Range
Dim emailTo, PDFSavePath As String

'current month DOR values
Dim YTDEbitaEmail, MTDEbitaEmail, YTDEbitaVsBudgetEmail, MTDEbitaVsBudgetEmail As Variant
Dim EbitaEmail, NetSlotsEmail, NetTableEmail, HotelFoodRetailEmail As Variant
Dim HotelMTDEmail, HotelMTDEmail_Total, HotelMTDEmail_Available, FoodMTDEmail As Variant

Set lookupList = Worksheets("Lookups").Range("Lookups")

'to test whether there are any broken DOR lookups before creating the email

'to test whether Outlook is open and to open it if it's not
Call TestOutlookIsOpen

Set DORFilePath = Worksheets("Setup").Range("DORHyperlink")

Set objOutlook = CreateObject("Outlook.Application")
Set eList = Worksheets("EmailRecipients").Range("EmailAddress")
Set DORDate = Worksheets("DOR Central").Range("DOR_Date")
Set DORDailyEmail = objOutlook.CreateItem(0)
Set TestEmail = Worksheets("Setup").Range("TestEmails")
PDFSavePath = Worksheets("Setup").Range("PDF_FileSavePath")

' current month DOR values
Set YTDEbitaEmail = Worksheets("Lookups").Range("YTDEbita_Email")
Set MTDEbitaEmail = Worksheets("Lookups").Range("MTDEbita_Email")
Set YTDEbitaVsBudgetEmail = Worksheets("Lookups").Range("YTDEbitaVsBudget_Email")
Set MTDEbitaVsBudgetEmail = Worksheets("Lookups").Range("MTDEbitaVsBudget_Email")

Set EbitaEmail = Worksheets("Lookups").Range("Ebita_Email")
Set NetSlotsEmail = Worksheets("Lookups").Range("NetSlots_Email")
Set NetTableEmail = Worksheets("Lookups").Range("NetTable_Email")
Set HotelFoodRetailEmail = Worksheets("Lookups").Range("HotelFoodRetailEmail")

Set HotelMTDEmail = Worksheets("Lookups").Range("HotelMTD_Email")
Set HotelMTDEmail_Total = Worksheets("Lookups").Range("HotelMTD_Email_Total")
Set HotelMTDEmail_Available = Worksheets("Lookups").Range("HotelMTD_Email_Available")
Set FoodMTDEmail = Worksheets("Lookups").Range("FoodMTD_Email")

'to loop through the email list on the "Email recipients" tab and add them to the recipient list
    For Each eAddress In eList
        emailTo = emailTo & ";" & eAddress.Value
    Next

        With DORDailyEmail
            .Subject = "Daily Operating Report - " & DORDate
                'Toggle between test email recipients and production email recipients
                If Worksheets("DOR Central").TestEmail.Value = True Then
                    .To = TestEmail
                ElseIf Worksheets("DOR Central").ProdEmail.Value = True Then
                    .To = emailTo
                End If
            .Attachments.Add (PDFSavePath)
            .Display
            .HTMLBody = "<body style=font-size:11pt;font-family:Calibri Light><p>Good Morning,</p>" _
                    & "<p>Please <A href=" & DORFilePath & ">click here</A> for the <b>" & Format(DORDate, "DDDD") & ", " _
                    & Format(DORDate, "MM") & "/" & Format(DORDate, "D") & "</b> DOR" _
                    & "<br> Attached is a PDF version of the DOR for viewing on mobile devices and some highlights below vs. budget<br><br>" _
                    & "As of <b>" & Format(DORDate, "M") & "/" & Format(DORDate, "DD") & " </b><i>est.</i>" _
                    & "<ul> <li>" & MTDEbitaEmail & "</li>" _
                    & "<li>" & YTDEbitaEmail & "</li>" _
                    & "<li>" & EbitaEmail & "</li>" _
                    & "<ul> <li>" & NetSlotsEmail & "</li>" _
                            & "<li>" & NetTableEmail & "</li>" _
                            & "<li>" & HotelFoodRetailEmail & "</li>" _
                            & "<li>Offset by all others</li>" _
                    & "</ul></ul>" _
                    & "<ul><li>" & HotelMTDEmail & "</li>" _
                            & "<ul><li>" & HotelMTDEmail_Available & "</li>" _
                            & "<li>" & HotelMTDEmail_Total & "</li>" _
                    & "</ul></ul>" _
                    & "<ul><li>" & FoodMTDEmail & "</li></ul></body>" _
                    & .HTMLBody
                    
          End With
   
End Sub


Sub MondayDOREmail()

Dim DORMondayEmail As Object
Dim objOutlook As Object
Dim DORThursdayDate, DORFridayDate, DORDate, DORDatePrevious As Variant
Dim DORFilePath, ThurDORFilePath, FriDORFilePath As Variant
Dim PDFSavePath, ThurPDFSavePath, FriPDFSavePath As Variant
Dim eList, eAddress, TestEmail As Range
Dim emailTo, previousEmailBody, currentEmailBody As String
Dim dorMonthNum, dorThursdayNum, dorFridayNum As Integer


'current month DOR values
Dim YTDEbitaEmail, MTDEbitaEmail, YTDEbitaVsBudgetEmail, MTDEbitaVsBudgetEmail As Variant
Dim EbitaEmail, NetSlotsEmail, NetTableEmail, HotelFoodRetailEmail As Variant
Dim HotelMTDEmail, HotelMTDEmail_Total, HotelMTDEmail_Available, FoodMTDEmail As Variant

'previous month DOR values
Dim YTDEbitaEmail_PREVIOUS, MTDEbitaEmail_PREVIOUS, YTDEbitaVsBudgetEmail_PREVIOUS, MTDEbitaVsBudgetEmail_PREVIOUS As Variant
Dim EbitaEmail_PREVIOUS, NetSlotsEmail_PREVIOUS, NetTableEmail_PREVIOUS, HotelFoodRetailEmail_PREVIOUS As Variant
Dim HotelMTDEmail_PREVIOUS, FoodMTDEmail_PREVIOUS As Variant

'to test whether Outlook is open and to open it if it's not
Call TestOutlookIsOpen


Set DORFilePath = Worksheets("Setup").Range("DORHyperlink")
Set ThurDORFilePath = Worksheets("Setup").Range("DORHyperlink_Thursday")
Set FriDORFilePath = Worksheets("Setup").Range("DORHyperlink_Friday")

Set objOutlook = CreateObject("Outlook.Application")
Set eList = Worksheets("EmailRecipients").Range("EmailAddress")
Set DORDate = Worksheets("DOR Central").Range("DOR_Date")
Set DORMondayEmail = objOutlook.CreateItem(0)
Set TestEmail = Worksheets("Setup").Range("TestEmails")

PDFSavePath = Worksheets("Setup").Range("PDF_FileSavePath")
ThurPDFSavePath = Worksheets("Setup").Range("PDF_FileSavePath_Thur")
FriPDFSavePath = Worksheets("Setup").Range("PDF_FileSavePath_Fri")

' current month DOR values
Set YTDEbitaEmail = Worksheets("Lookups").Range("YTDEbita_Email")
Set MTDEbitaEmail = Worksheets("Lookups").Range("MTDEbita_Email")
Set YTDEbitaVsBudgetEmail = Worksheets("Lookups").Range("YTDEbitaVsBudget_Email")
Set MTDEbitaVsBudgetEmail = Worksheets("Lookups").Range("MTDEbitaVsBudget_Email")

Set EbitaEmail = Worksheets("Lookups").Range("Ebita_Email")
Set NetSlotsEmail = Worksheets("Lookups").Range("NetSlots_Email")
Set NetTableEmail = Worksheets("Lookups").Range("NetTable_Email")
Set HotelFoodRetailEmail = Worksheets("Lookups").Range("HotelFoodRetailEmail")

Set HotelMTDEmail = Worksheets("Lookups").Range("HotelMTD_Email")
Set HotelMTDEmail_Total = Worksheets("Lookups").Range("HotelMTD_Email_Total")
Set HotelMTDEmail_Available = Worksheets("Lookups").Range("HotelMTD_Email_Available")
Set FoodMTDEmail = Worksheets("Lookups").Range("FoodMTD_Email")

' previous month DOR values
Set YTDEbitaEmail_PREVIOUS = Worksheets("Lookups").Range("YTDEbita_Email_PREVIOUS")
Set MTDEbitaEmail_PREVIOUS = Worksheets("Lookups").Range("MTDEbita_Email_PREVIOUS")
Set YTDEbitaVsBudgetEmail_PREVIOUS = Worksheets("Lookups").Range("YTDEbitaVsBudget_Email_PREVIOUS")
Set MTDEbitaVsBudgetEmail_PREVIOUS = Worksheets("Lookups").Range("MTDEbitaVsBudget_Email_PREVIOUS")

Set EbitaEmail_PREVIOUS = Worksheets("Lookups").Range("Ebita_Email_PREVIOUS")
Set NetSlotsEmail_PREVIOUS = Worksheets("Lookups").Range("NetSlots_Email_PREVIOUS")
Set NetTableEmail_PREVIOUS = Worksheets("Lookups").Range("NetTable_Email_PREVIOUS")
Set HotelFoodRetailEmail_PREVIOUS = Worksheets("Lookups").Range("HotelFoodRetailEmail_PREVIOUS")

Set HotelMTDEmail_PREVIOUS = Worksheets("Lookups").Range("HotelMTD_Email_PREVIOUS")
Set FoodMTDEmail_PREVIOUS = Worksheets("Lookups").Range("FoodMTD_Email_PREVIOUS")

' dates for Thursday and Friday DOR's
DORThursdayDate = DateAdd("D", -2, DORDate)
DORFridayDate = DateAdd("D", -1, DORDate)

dorMonthNum = Month(DORDate)
dorThursdayNum = Month(DORThursdayDate)
dorFridayNum = Month(DORFridayDate)

' to determine which day the end of the previous month fell on
If dorThursdayNum = dorFridayNum Then
    DORDatePrevious = DORFridayDate
Else
    DORDatePrevious = DORThursdayDate
End If


previousEmailBody = "As of <b>" & Format(DORDatePrevious, "M") & "/" & Format(DORDatePrevious, "DD") & " </b><i>est.</i>" _
                    & "<ul> <li>" & YTDEbitaEmail_PREVIOUS & "</li>" _
                    & "<li>" & MTDEbitaEmail_PREVIOUS & "</li>" _
                    & "<li>" & YTDEbitaVsBudgetEmail_PREVIOUS & "</li>" _
                    & "<li>" & MTDEbitaVsBudgetEmail_PREVIOUS & "</li>" _
                    & "<li>" & EbitaEmail_PREVIOUS & "</li>" _
                    & "<ul> <li>" & NetSlotsEmail_PREVIOUS & "</li>" _
                            & "<li>" & NetTableEmail_PREVIOUS & "</li>" _
                            & "<li>" & HotelFoodRetailEmail_PREVIOUS & "</li>" _
                            & "<li>Offset by all others</li>" _
                    & "</ul></ul>" _
                    & "<ul><li>" & HotelMTDEmail_PREVIOUS & "</li>" _
                    & "<li>" & FoodMTDEmail_PREVIOUS & "</li></ul></body>"

currentEmailBody = "As of <b>" & Format(DORDate, "M") & "/" & Format(DORDate, "DD") & " </b><i>est.</i>" _
                    & "<ul> <li>" & MTDEbitaEmail & "</li>" _
                    & "<li>" & YTDEbitaEmail & "</li>" _
                    & "<li>" & EbitaEmail & "</li>" _
                    & "<ul> <li>" & NetSlotsEmail & "</li>" _
                            & "<li>" & NetTableEmail & "</li>" _
                            & "<li>" & HotelFoodRetailEmail & "</li>" _
                            & "<li>Offset by all others</li>" _
                    & "</ul></ul>" _
                    & "<ul><li>" & HotelMTDEmail & "</li>" _
                    & "<li>" & FoodMTDEmail & "</li></ul></body>"

'to loop through the email list on the "Email recipients" tab and add them to the recipient list
    For Each eAddress In eList
        emailTo = emailTo & ";" & eAddress.Value
    Next

        With DORMondayEmail
            .Subject = "Daily Operating Report - " & DORThursdayDate & "-" & DORDate
                'Toggle between test email recipients and production email recipients
                If Worksheets("DOR Central").TestEmail.Value = True Then
                    .To = TestEmail
                ElseIf Worksheets("DOR Central").ProdEmail.Value = True Then
                    .To = emailTo
                End If
            .Display
                With .Attachments 'add all three PDF's from Thursday thru Saturday
                        .Add (ThurPDFSavePath)
                        .Add (FriPDFSavePath)
                        .Add (PDFSavePath)
                End With
            If dorThursdayNum <> dorMonthNum _
                Or dorFridayNum <> dorMonthNum Then
            'email body for Monday DOR's that fall between two different months

            .HTMLBody = "<body style=font-size:11pt;font-family:Calibri Light><p>Good Morning,</p>" _
                    & "<p>Please <A href=" & ThurDORFilePath & ">click here</A> for the <b>" & Format(DORThursdayDate, "DDDD") & ", " _
                    & Format(DORThursdayDate, "MM") & "/" & Format(DORThursdayDate, "D") & "</b> DOR" _
                    & "<p>Please <A href=" & FriDORFilePath & ">click here</A> for the <b>" & Format(DORFridayDate, "DDDD") & ", " _
                    & Format(DORFridayDate, "MM") & "/" & Format(DORFridayDate, "D") & "</b> DOR" _
                    & "<br>" _
                    & "<p>Please <A href=" & DORFilePath & ">click here</A> for the <b>" & Format(DORDate, "DDDD") & ", " _
                    & Format(DORDate, "MM") & "/" & Format(DORDate, "D") & "</b> DOR" _
                    & "<br>" _
                    & "<br> Attached are PDF versions of the DOR for viewing on mobile devices and some highlights below vs. budget<br><br>" _
                    & previousEmailBody _
                    & "<br>" _
                    & currentEmailBody _
                    & .HTMLBody

            Else
            'email body for Monday DOR's that fall in the same month
            
            .HTMLBody = "<body style=font-size:11pt;font-family:Calibri Light><p>Good Morning,</p>" _
                    & "<p>Please <A href=" & ThurDORFilePath & ">click here</A> for the <b>" & Format(DORThursdayDate, "DDDD") & ", " _
                    & Format(DORThursdayDate, "MM") & "/" & Format(DORThursdayDate, "D") & "</b> DOR" _
                    & "<p>Please <A href=" & FriDORFilePath & ">click here</A> for the <b>" & Format(DORFridayDate, "DDDD") & ", " _
                    & Format(DORFridayDate, "MM") & "/" & Format(DORFridayDate, "D") & "</b> DOR" _
                    & "<br>" _
                    & "<p>Please <A href=" & DORFilePath & ">click here</A> for the <b>" & Format(DORDate, "DDDD") & ", " _
                    & Format(DORDate, "MM") & "/" & Format(DORDate, "D") & "</b> DOR" _
                    & "<br>" _
                    & "<br> Attached are PDF versions of the DOR for viewing on mobile devices and some highlights below vs. budget<br><br>" _
                    & currentEmailBody _
                    & .HTMLBody

                    
            End If
          End With

End Sub


Sub WeeklyDOREmail()

Dim DORWeeklyEmail As Object
Dim objOutlook As Object
Dim WeeklyDORDate As Variant
Dim DORFilePath, PDFSavePath As Variant
Dim eList, eAddress, TestEmail As Range
Dim emailTo As String

'to test whether Outlook is open and to open it if it's not
Call TestOutlookIsOpen

Set DORFilePath = Worksheets("Setup").Range("DORHyperlink_Weekly")

Set objOutlook = CreateObject("Outlook.Application")
Set eList = Worksheets("EmailRecipients").Range("EmailAddress")
Set WeeklyDORDate = Worksheets("DOR Central").Range("DOR_Date_Weekly")
Set DORWeeklyEmail = objOutlook.CreateItem(0)
Set TestEmail = Worksheets("Setup").Range("TestEmails")

PDFSavePath = Worksheets("Setup").Range("PDF_FileSavePath_Weekly")



'to loop through the email list on the "Email recipients" tab
    For Each eAddress In eList
        emailTo = emailTo & ";" & eAddress.Value
    Next

        With DORWeeklyEmail
            .Subject = "Weekly Operating Report - " & WeeklyDORDate
                'Toggle between test email recipients and production email recipients
                If Worksheets("DOR Central").TestEmail.Value = True Then
                    .To = TestEmail
                ElseIf Worksheets("DOR Central").ProdEmail.Value = True Then
                    .To = emailTo
                End If
            .Attachments.Add PDFSavePath
            .Display
            .HTMLBody = "<body style=font-size:11pt;font-family:Calibri Light><p>Good Morning,</p>" _
                    & "<p> The Weekly Operating Report has been updated; please <A href=" & DORFilePath & ">click here</A> for the Weekly Operating Report for the period ending <b>" & Format(WeeklyDORDate, "MM/DD/YY") & ".</b>" _
                    & "<br><br> Attached is a PDF version of the Weekly Operating report for viewing on mobile devices.<br><br>" _
                    & .HTMLBody
          End With

End Sub
 Sub EmailSelector()

 
 If Worksheets("DOR Central").TestEmail.Value = True Then
        If Worksheets("DOR Central").DailyDOR.Value = True Then
            Call DailyDOREmail
        ElseIf Worksheets("DOR Central").MondayDOR.Value = True Then
            Call MondayDOREmail
        ElseIf Worksheets("DOR Central").WeeklyDOR.Value = True Then
            Call WeeklyDOREmail
    End If
    
ElseIf Worksheets("DOR Central").ProdEmail.Value = True Then
                 If Worksheets("DOR Central").DailyDOR.Value = True Then
                    Call DailyDOREmail
                ElseIf Worksheets("DOR Central").MondayDOR.Value = True Then
                    Call MondayDOREmail
                ElseIf Worksheets("DOR Central").WeeklyDOR.Value = True Then
                    Call WeeklyDOREmail
            Else
                Exit Sub
        End If
End If

End Sub

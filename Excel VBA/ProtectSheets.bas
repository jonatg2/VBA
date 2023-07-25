Attribute VB_Name = "ProtectSheets"
Sub LockAllSheets()
Dim wsheet As Worksheet
    For Each wsheet In ThisWorkbook.Worksheets
        wsheet.Protect Password:=""
Next wsheet

End Sub

Sub UnlockAllSheets()

Dim wsheet As Worksheet
    For Each wsheet In ThisWorkbook.Worksheets
        wsheet.Unprotect Password:=""
Next wsheet

End Sub

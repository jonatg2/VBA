Attribute VB_Name = "DeleteFile"
Option Explicit

Function FileExists(ByVal FileToTest As String) As Boolean
   FileExists = (Dir(FileToTest) <> "")
End Function




Sub DeleteDORFile(ByVal FileToDelete As String)
On Error GoTo ErrHandler

           If FileExists(FileToDelete) Then 'See above
              ' First remove readonly attribute, if set
              SetAttr FileToDelete, vbNormal
              ' Then delete the file
              Kill FileToDelete
        End If
ErrHandler:
            If Err.Number = 70 Then
                Beep
                MsgBox "Permission denied for the following file: " _
                & Chr(10) & Chr(10) & FileToDelete & Chr(10) & Chr(10) _
                & " This is most likely due to: " _
        & Chr(10) & Chr(9) & Chr(149) & "The file is currently opened by you" _
        & Chr(10) & Chr(9) & Chr(149) & "The file is opened by another user" _
        & Chr(10) & Chr(10) & "Please check to see if the file is still opened by you or another user, close the file, then retry.", vbCritical, "Permission Denied"

    End
End If

End Sub

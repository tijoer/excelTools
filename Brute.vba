Sub BruteForce()
' Breaks worksheet password protection.
' Displays the password of a worksheet and displays its value.
' Relies on hash collision

Dim i As Long, j As Long, k As Long
Dim l As Long, m As Long, n As Long
Dim i1 As Long, i2 As Long, i3 As Long
Dim i4 As Long, i5 As Long, i6 As Long

Dim clipboard As MSForms.DataObject
Dim password  As String

On Error Resume Next
For i = 65 To 66:
    For j = 65 To 66:
        For k = 65 To 66
            For l = 65 To 66:
                For m = 65 To 66:
                    For i1 = 65 To 66
                        For i2 = 65 To 66:
                            For i3 = 65 To 66:
                                For i4 = 65 To 66
                                    For i5 = 65 To 66:
                                        For i6 = 65 To 66:
                                            For n = 32 To 126
                                                password = Chr(i) & Chr(j) & _
                                                    Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
                                                    Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
                                            
                                                ActiveSheet.Unprotect password
                                                    
                                                If ActiveSheet.ProtectContents = False Then
                                                    MsgBox "One usable password is:" & vbNewLine & _
                                                    password & vbNewLine & _
                                                    "Copied password to clipbard."
                                                    
                                                    Set clipboard = New MSForms.DataObject
                                                    clipboard.SetText password
                                                    clipboard.PutInClipboard
                                                    
                                                    Exit Sub
                                                End If
                                            Next:
                                        Next:
                                    Next:
                                Next:
                            Next:
                        Next:
                    Next:
                Next:
            Next:
        Next:
    Next:
Next:
End Sub

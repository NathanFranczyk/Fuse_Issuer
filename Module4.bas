Attribute VB_Name = "Module4"
Function getsmd(fuse As String)
'returns a string in the form of smd1a
'function disects string and searches for a sequence that starts with S and ends on a space or end of the string
'assume smd1a if info not available
    Dim smd As String
    Dim counter As Integer
    For counter = 1 To Len(fuse)
        If smd <> "" And Mid(fuse, counter, 1) = " " Then
            getsmd = smd
            Exit Function
        End If
        
        If Mid(fuse, counter, 1) = "S" Or Mid(fuse, counter, 1) = "s" Then
            counter = counter + 1
            smd = "S"
            While Mid(fuse, counter, 1) <> " " And counter <= Len(fuse)
                If Mid(fuse, counter, 1) <> "-" Then
                    smd = smd & Mid(fuse, counter, 1)
                End If
            counter = counter + 1
            Wend
                '^go to next part of the string
        End If
        
    Next counter
    
    If smd = "" Then
        getsmd = "SMD1A"
        Exit Function
        
    End If
        getsmd = smd
End Function


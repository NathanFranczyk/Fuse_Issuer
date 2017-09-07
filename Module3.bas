Attribute VB_Name = "Module3"
Function getsize(fuse As String)
'returns a string in the form of a number with E at the end.
'function disects the string to find a sequence that starts with a number and ends with an E before a space or end of the string
'^must be in that format to work

    Dim E As String
    Dim counter As Integer
    For counter = 1 To Len(fuse)
        If E <> "" And Mid(fuse, counter, 1) = "E" Then
            E = E & "E"
            getsize = E
            Exit Function
        End If
        
        If IsNumeric(Mid(fuse, counter, 1)) Then
            E = E & Mid(fuse, counter, 1)
        Else
            While Mid(fuse, counter, 1) <> " " And counter < Len(fuse)
                E = ""
                counter = counter + 1
            Wend
            '^go to next part of the string
        End If
    Next counter
    
    getsize = "X"
        
End Function


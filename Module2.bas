Attribute VB_Name = "Module2"
Function getspeed(fuse As String)
'Function that returns the speed of 153, 119, etc.
'Input argument is a string that will be disected. This must be in a form that starts and ends with a number
'This will only read the speed if only the speed starts and ends with a number before a space or end of the string

    Dim speed As String
    Dim counter As Integer
    For counter = 1 To Len(fuse)
        If speed <> "" And Mid(fuse, counter, 1) = " " Then
            getspeed = Mid(speed, 1, 3)
            'speed should only be numbers
            Exit Function
        End If
        
        If IsNumeric(Mid(fuse, counter, 1)) Or Mid(fuse, counter, 1) = "-" Then
            speed = speed & Mid(fuse, counter, 1)
            'If counter = len(fuse) and
        Else
            While Mid(fuse, counter, 1) <> " " And counter <= Len(fuse)
                speed = ""
                counter = counter + 1
            Wend
            '^go to next part of the string
        End If
    Next counter
    
    If speed = "" Then
        getspeed = "153"
        Exit Function
    End If
    
    getspeed = Mid(speed, 1, 3)
    'speed should not reach this point. if so, an X is issued
End Function

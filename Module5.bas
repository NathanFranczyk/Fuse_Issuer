Attribute VB_Name = "Module5"
Function closestkv(numb As Double)
'This function allows the input number to be rounded to the closest voltage of either
'34, 69, or 138 kv
    If numb > 0 And numb < 52 Then
        closestkv = 34
        Exit Function
    End If
    If numb > 51 And numb < 104 Then
        closestkv = 69
        Exit Function
    End If

    closestkv = 138
    
    
    
End Function


Attribute VB_Name = "Module7"
Sub formatter()
    Dim wks As Worksheet
    Dim lastrow As Integer
    Dim rowrange As Range
    
    Set wks = ThisWorkbook.Sheets(1)
    
    lastrow = wks.Cells(wks.Rows.Count, "A").End(xlUp).row
    Set rowrange = wks.Range("A1:A" & lastrow)
    
    For i = 1 To lastrow
        wks.Cells(i, 3) = Replace(wks.Cells(i, 3), "-", " ")
        wks.Cells(i, 4) = Replace(wks.Cells(i, 4), "-", " ")
        wks.Cells(i, 3) = Replace(wks.Cells(i, 3), "'", "")
        wks.Cells(i, 4) = Replace(wks.Cells(i, 4), "'", "")
    Next i
    
End Sub

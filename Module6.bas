Attribute VB_Name = "Module6"
Function findbestfuse(highkv As Integer, infinitecurrent As Double, kconst As Double, dividingcurrent As Double)
'function that will return a fuse string that will protect the transformer by checking mechanical damage curve
'takes in infinitecurrent and the kconstant, as well as the high side kv
'This finds the best curve, which I defined as the curve that has the closest average distance to the clearance
    Dim speedarray As Variant
    Dim smdarray As Variant
    Dim datacol As Integer
    Dim badfuse As Integer
    Dim rowcount As Integer
    Dim tempint As Integer
    Dim tempint2 As Integer
    
    Dim smdi As Integer
    Dim speedi As Integer
    Dim exists As Integer
    Dim curravgtimesum As Double
    Dim returnstring As String
    Dim row2count As Integer
    Dim mechtime As Double
    Dim curravgtimedif As Double
    Dim bestavgtimedif As Double
    Dim curriterations As Double
    Dim avgtime As Double
    Dim avgcurrent As Double
    speedarray = Array(153, 119, 176)
    smdarray = Array("smd1a", "smd2b", "smd2c", "smd3", "smd50", "sm4", "sm5", "smu20")
    bestavgtimedif = 9999999
    speedi = 0
    smdi = 0
    For Each speed In speedarray
        smdi = 0
        For Each smd In smdarray
            i = 1
            exists = 0
            
            For i = 1 To ThisWorkbook.Worksheets.Count
                If ThisWorkbook.Sheets(i).Name = speedarray(speedi) & smdarray(smdi) & highkv & "kvclear" Then
                    exists = 1
                End If
            Next i
            
            If exists = 1 Then
                'Debug.Print ("CHECKING: " & speedarray(speedi) & smdarray(smdi) & highkv)
                Set fusesheet = ThisWorkbook.Sheets(speedarray(speedi) & smdarray(smdi) & highkv & "kvclear")
                '^opens the workbook that contains the fuse data
                
                Set fuserange = fusesheet.Range("A1:A" & fusesheet.Cells(fusesheet.Rows.Count, "A").End(xlUp).row)
                datacol = 2
                
                For Each Value In fuserange
                    If datacol Mod 2 = 0 Then
                        curravgtimesum = 0
                        curravgtimedif = 0
                        curriterations = 0
                        row2count = 1
                        badfuse = 0
                        
                        For Each row2 In fuserange
                            If IsNumeric(fusesheet.Cells(row2count, datacol)) Then
                                If CDbl(fusesheet.Cells(row2count, datacol)) > (dividingcurrent) And CDbl(fusesheet.Cells(row2count, datacol)) < infinitecurrent Then
                                    'within range to compare
                                    tempint = datacol + 1
                                    badfuse = 2
                                    mechtime = kconst / (CDbl((fusesheet.Cells(row2count, datacol))) * CDbl(fusesheet.Cells(row2count, datacol)))
                                    'Debug.Print (fusesheet.Cells(row2count, tempint))
                                    If CDbl(fusesheet.Cells(row2count, tempint)) > mechtime Then
                                        'this fuse does not protect the xfmr. highlight
                                        'wks.Cells(rowcount, oneline).Interior.ColorIndex = 3
                                        'highlight fuseinservice cell
                                        badfuse = 1
                                        Exit For
                                    End If
                                    curravgtimesum = curravgtimesum + (mechtime - CDbl(fusesheet.Cells(row2count, tempint)))
                                    curriterations = curriterations + 1
                                End If
                            End If
                        row2count = row2count + 1
                        Next row2
                    If curriterations <> 0 Then
                         curravgtimedif = curravgtimesum / CDbl(curriterations)
                    End If
                    
                    If curravgtimedif < bestavgtimedif And curravgtimedif <> 0 And badfuse = 2 Then
                        bestavgtimedif = curravgtimedif
                        returnstring = ""
                        returnstring = CStr(speedarray(speedi))
                        returnstring = returnstring & " " & CStr(smdarray(smdi)) & " " & fusesheet.Cells(6, datacol)
                        'Debug.Print (returnstring & " is currently the best fuse")
                    End If
                End If
                        datacol = datacol + 1
                Next Value
                'Call fusedata.Close
            End If
        smdi = smdi + 1
        Next smd
    speedi = speedi + 1
    Next speed
    findbestfuse = returnstring
End Function

Attribute VB_Name = "Module1"
Sub Makesettings()
'this subroutine is the main part of the program that will do all calculations and eventually issue the settings
Application.DisplayAlerts = False
    
    
    Dim settings As Integer
    Dim fuseinservice As Integer
    Dim oneline As Integer
    Dim needswork As Integer
    Dim loc As Integer
    Dim numberofsheets As Integer
    
    settings = 7
    fuseinservice = 4
    oneline = 5
    needswork = 8
    loc = 2
    numberofsheets = 4
    '^^^UPDATE THESE FOR EACH DIVISION
    
    Dim lastrow As Long
    Dim lastxfmrrow As Long
    Dim sandc As Integer
    Dim i As Integer
    Dim wkscount As Integer
    Dim xfmrcount As Integer
    Dim j As Integer
    Dim basecurrent As Double
    Dim mva As Double
    Dim percentimpedance As Double
    Dim pucurrent As Double
    Dim infinitebuscurrent As Double
    Dim kconst As Double
    Dim avgcurrent As Double
    Dim avgtime As Double
    Dim curravgcurrent As Double
    Dim halfinfinitetime As Double
    Dim dividingcurrent As Double
    Dim datacol As Integer
    Dim datacount As Integer
    Dim tempint As Integer
    Dim badfuseservice As Integer
    Dim badfuseone As Integer
    Dim rowcount As Integer
    Dim xfmrrowcount As Integer
    Dim row2count As Integer
    Dim tempint2 As Integer
    
    Dim deltawye As Integer
    '^1 for deltawye, 0 else
    
    Dim rowrange As Range
    Dim xfmrrange As Range
    Dim fuserange As Range
    
    Dim settingfusestr As String
    Dim fuseservicestr As String
    Dim fuseonelinestr As String
    Dim fuseservicespeed As String
    Dim fuseservicesize As String
    Dim onespeed As String
    Dim onesize As String
    Dim xfmrloc As String
    Dim onesmd As String
    Dim fuseservicesmd As String
    Dim highkv As String
    Dim locstr As String
    Dim settingsmd As String
    Dim settingspeed As String
    Dim settingsize As String
    Dim exists As Integer
    
    Dim belowclassIII As Boolean
    
    Dim fusecoordbook As Workbook
    Dim xfmrdata As Workbook
    Dim settingbook As Workbook
    Dim settingsheet As Worksheet
    Dim wks As Worksheet
    Dim xfmrsheet As Worksheet
    Dim fusesheet As Worksheet
    
    i = 1
    wkscount = ActiveWorkbook.Worksheets.Count
    Set fusecoordbook = ActiveWorkbook

    For i = 1 To numberofsheets
        Set wks = fusecoordbook.Sheets(i)
        
        
        lastrow = wks.Cells(wks.Rows.Count, "A").End(xlUp).row
        Set rowrange = wks.Range("A1:A" & lastrow)
        '^Set up last row and range
        
        Set xfmrdata = Workbooks.Open("Z:\Relay Decatur\Xfmr Fuse_files\_2017_HS_FUSE_RECORD_UPDATE\Division 1 Fuse Calculations minmelt\AllXfmrData.xls")
        '^^Open workbook containing xfmr data
        
        Set xfmrsheet = xfmrdata.Sheets(1)
        wks.Activate
        
        lastxfmrrow = xfmrsheet.Cells(xfmrsheet.Rows.Count, "A").End(xlUp).row
        Set xfmrrange = xfmrsheet.Range("A1:A" & lastxfmrrow)
        rowcount = 1
        
        For Each row In rowrange
        '^^iterate through each row in the sheet
            
            If wks.Cells(rowcount, needswork).Value = "Y" Then
                fuseservicestr = wks.Cells(rowcount, fuseinservice).Value
                fuseonelinestr = wks.Cells(rowcount, oneline).Value
                onespeed = getspeed(fuseonelinestr)
                onesize = getsize(fuseonelinestr)
                onesmd = getsmd(fuseonelinestr)
                fuseservicesmd = getsmd(fuseservicestr)
                fuseservicesize = getsize(fuseservicestr)
                fuseservicespeed = getspeed(fuseservicestr)
                xfmrloc = wks.Cells(rowcount, loc).Value
                'Debug.Print (xfmrloc)
                xfmrrowcount = 1
                For Each xfmrrow In xfmrrange
                    If xfmrloc = xfmrsheet.Cells(xfmrrowcount, 3).Value Then GoTo Cont
                    xfmrrowcount = xfmrrowcount + 1
                Next xfmrrow
                '^finds the row in the xfmr data file
                
Cont:
                'On Error GoTo nextrowerror
                highkv = closestkv(CDbl(xfmrsheet.Cells(xfmrrowcount, 7)))
                
                mva = CDbl(xfmrsheet.Cells(xfmrrowcount, 10))
                
                If Not IsNumeric(mva) Then
                    mva = CDbl(xfmrsheet.Cells(xfmrrowcount, 11))
                End If

                percentimpedance = CDbl(xfmrsheet.Cells(xfmrrowcount, 9))
                
                If highkv = 0 Or mva = 0 Or percentimpedance = 0 Then
                    Debug.Print ("CHECK XFMR Data at" & xfmrloc)
                    GoTo NEXTROW
                End If
                '^Something doesn't exist in the xfmrdata. will cause an error
                
                If xfmrsheet.Cells(xfmrrowcount, 8) = "DELTA/WYE" Then
                    deltawye = 1
                Else
                    deltawye = 0
                End If
                
                basecurrent = (mva * CDbl(1000000)) / (CDbl(1.73) * highkv * CDbl(1000))
                pucurrent = CDbl(1) / (percentimpedance / CDbl(100))
                '^power calculation to get the per unit current
                infinitebuscurrent = basecurrent * pucurrent
                
                If deltawye = 1 Then
                    infinitebuscurrent = infinitebuscurrent / CDbl(1.73)
                End If
                '^this makes the endpoint for the line that needs to be checked.
                'This point always occurs at 2 seconds after fault
                'So the resulting point is (infinitebuscurrent, 2)
                
                kconst = infinitebuscurrent * infinitebuscurrent * 2
                'K = I^2 *t
                'Above is the formula for the infinite bus potential curve that will check for
                'mechanical damage to the transformer. This is what we really coordinate with
                'NOTE: the other endpoint is one half of the infinite bus current
                
                halfinfinitetime = kconst / (CDbl(0.25) * (infinitebuscurrent * infinitebuscurrent))
                'other endpoint = (.5 infinitebuscurrent, halfinfinitebustime)
                
                If mva < CDbl(5) Then
                    belowclassIII = True
                End If
                
                
                'Below is a condition that chooses which column to look at. fuse in service is priority, then one line

                datacount = 1
                row2count = 1
                datacol = 2
                strarr = Array(fuseservicestr, fuseonelinestr)
                smdarr = Array(fuseservicesmd, onesmd)
                speedarr = Array(fuseservicespeed, onespeed)
                sizearr = Array(fuseservicesize, onesize)
                colarr = Array(fuseinservice, oneline)
                tempint = 0
                badfusearr = Array(0, 0)
                smdmeltarr = Array("smd1a", "smd1a")
                If belowclassIII = True Then
                        dividingcurrent = infinitebuscurrent * CDbl(0.7)
                    Else
                        dividingcurrent = infinitebuscurrent * CDbl(0.5)
                End If
                
                For j = 0 To 1
                
                    If strarr(j) <> "X" And strarr(j) <> "N/A" And strarr(j) <> "NA" And sizearr(j) <> "X" Then
                        If LCase(smdarr(j)) = "smd1a" Or LCase(smdarr(j)) = "smd2c" Or LCase(smdarr(j)) = "smd2b" _
                        Or LCase(smdarr(j)) = "smd3" Or LCase(smdarr(j)) = "smd50" Then
                            smdmeltarr(j) = "smd"
                        End If
                        
                        If LCase(smdarr(j)) = "sm4" Or LCase(smdarr(j)) = "sm5" Then
                            smdmeltarr(j) = "sm"
                        End If
                        
                        If LCase(smdarr(j)) = "smu20" Or LCase(smdarr(j)) = "smu40" Then
                            smdmeltarr(j) = "smu"
                        End If
                        
                        exists = 0
                        For X = 1 To ThisWorkbook.Worksheets.Count
                            If ThisWorkbook.Sheets(X).Name = speedarr(j) & LCase(smdmeltarr(j)) & "allkvminmelt" Then
                                exists = 1
                            End If
                        Next X
                        If exists = 1 Then
                            Set fusesheet = ThisWorkbook.Sheets(speedarr(j) & LCase(smdmeltarr(j)) & "allkvminmelt")
                            '^opens the workbook that contains the fuse data
                            
                            Set fuserange = fusesheet.Range("A1:A" & fusesheet.Cells(fusesheet.Rows.Count, "A").End(xlUp).row)
                            datacol = 2
                            
                            For Each Value In fuserange
                                    If datacol Mod 2 = 0 Then
                                        If fusesheet.Cells(6, datacol) = sizearr(j) Then
                                            row2count = 1
                                            prevI = 0
                                            currI = 0
                                            For Each row2 In fuserange
                                                    If IsNumeric(fusesheet.Cells(row2count, datacol)) Then
                                                        
                                                        If CDbl(fusesheet.Cells(row2count, datacol)) > (dividingcurrent) And CDbl(fusesheet.Cells(row2count, datacol)) < infinitebuscurrent Then
                                                            'within range to compare
                                                            tempint = datacol + 1
                                                            badfusearr(j) = 2
                                                            mechtime = kconst / (CDbl((fusesheet.Cells(row2count, datacol))) * CDbl(fusesheet.Cells(row2count, datacol)))
                                                            If CDbl(fusesheet.Cells(row2count, tempint)) > mechtime Then
                                                                'this fuse does not protect the xfmr. highlight
                                                                'highlight fuseinservice cell
                                                                badfusearr(j) = 1
                                                                wks.Cells(rowcount, colarr(j)).Interior.ColorIndex = 3
                                                                Exit For
                                                            End If
                                                        Else
                                                            'Debug.Print ("out of range")
                                                        End If
                                                    End If
                                                row2count = row2count + 1
                                            Next row2
                                            
                                        End If
                                        
                                    End If
                                    datacol = datacol + 1
                                Next Value
                            End If
                        End If
                        
                    Next j
setsetting:
                    locstr = Mid(row, 1, 4)
                    
                    
                    If badfusearr(1) = 2 Then
                        settingsmd = smdarr(1)
                        settingspeed = speedarr(1)
                        settingsize = sizearr(1)
                    End If
                    
                    If badfusearr(0) = 2 Then
                        settingsmd = smdarr(0)
                        settingspeed = speedarr(0)
                        settingsize = sizearr(0)
                    End If
                    '^^Fuses are good. Use them to issue settings sheet
                    'If both fuses are good, use the fuse in service.
                    
                    If badfusearr(0) = 0 And strarr(0) <> "X" And strarr(0) <> "N/A" And strarr(0) <> "NA" Then
                        wks.Cells(rowcount, colarr(0)).Interior.ColorIndex = 6
                    End If
                    If badfusearr(1) = 0 And strarr(1) <> "X" And strarr(1) <> "N/A" And strarr(1) <> "NA" And sizearr(1) <> "X" Then
                        wks.Cells(rowcount, colarr(1)).Interior.ColorIndex = 6
                    End If
                    '^Highlight fuses in yellow if they cannot be read, data not found,
                    'or no current point in fuse data falls in the mech damage curve
                    
                    If badfusearr(0) <> 2 And badfusearr(1) <> 2 Then
                        settingsmd = getsmd(findbestfuse(CInt(highkv), infinitebuscurrent, kconst, dividingcurrent))
                        settingspeed = getspeed(findbestfuse(CInt(highkv), infinitebuscurrent, kconst, dividingcurrent))
                        settingsize = getsize(findbestfuse(CInt(highkv), infinitebuscurrent, kconst, dividingcurrent))
                    End If
                    
                    If Len(Dir("Z:\Relay Decatur\Xfmr Fuse_files\_2017_HS_FUSE_RECORD_UPDATE\Division 1 Fuse Calculations minmelt\" & wks.Cells(rowcount, 1), vbDirectory)) = 0 Then
                            MkDir ("Z:\Relay Decatur\Xfmr Fuse_files\_2017_HS_FUSE_RECORD_UPDATE\Division 1 Fuse Calculations minmelt\" & wks.Cells(rowcount, 1))
                    End If
                    
                    If Len(Dir("Z:\Relay Decatur\Xfmr Fuse_files\_2017_HS_FUSE_RECORD_UPDATE\Division 1 Fuse Calculations minmelt\" & wks.Cells(rowcount, 1) & "\" & locstr & "_Dist-trf-Recl.xlsx")) <> 0 Then
                        Set settingbook = Workbooks.Open("Z:\Relay Decatur\Xfmr Fuse_files\_2017_HS_FUSE_RECORD_UPDATE\Division 1 Fuse Calculations minmelt\" & wks.Cells(rowcount, 1) & "\" & locstr & "_Dist-trf-Recl.xlsx")
                    Else
                        Set settingbook = Workbooks.Open("Z:\Relay Decatur\Xfmr Fuse_files\_2017_HS_FUSE_RECORD_UPDATE\Division 1 Fuse Calculations minmelt\settings_template.xlsx")
                    End If
                        '^Open the settingsheet if it exists. Otherwise, make a new one
                    
                    Set settingsheet = settingbook.Sheets(1)
                    'settingsheet.SaveAs ("Z:\Relay Decatur\Xfmr Fuse_files\_2017_HS_FUSE_RECORD_UPDATE\Division 1 Fuse Calculations minmelt\" & wks.Cells(rowcount, 1) & "\" & locstr & "_Dist-trf-Recl")
                    
                    tempint2 = 8
                    tempint = 0
                    While settingsheet.Cells(tempint2, 2) <> ""
                        tempint = tempint + 1
                        tempint2 = 8 + (4 * tempint)
                    Wend
                    
                    settingsheet.Cells(2, 2) = wks.Cells(rowcount, 1)
                    tempint2 = 8 + (4 * tempint)
                    settingsheet.Cells(tempint2, 2) = "XFMR#" & Mid(wks.Cells(rowcount, loc), Len(wks.Cells(rowcount, loc)), 1)
                    tempint2 = 9 + (4 * tempint)
                    settingsheet.Cells(tempint2, 2) = settingsmd
                    tempint2 = 10 + (4 * tempint)
                    settingsheet.Cells(tempint2, 2) = settingspeed
                    tempint2 = 11 + (4 * tempint)
                    settingsheet.Cells(tempint2, 2) = settingsize
                    settingsheet.Cells(2, 7) = Date
                    '^^Update settings sheet
                    
                    
                    wks.Cells(rowcount, needswork) = "N"
                    wks.Cells(rowcount, settings) = settingsmd & " " & settingspeed & " " & settingsize
                    wks.Cells(rowcount, settings).Interior.ColorIndex = 4
                    
                    Debug.Print ("   " & xfmrloc & " " & "Issued Fuse: " & settingsmd & " " & settingspeed & " " & settingsize)
                    
                    Call settingbook.Close(True, "Z:\Relay Decatur\Xfmr Fuse_files\_2017_HS_FUSE_RECORD_UPDATE\Division 1 Fuse Calculations minmelt\" & wks.Cells(rowcount, 1) & "\" & locstr & "_Dist-trf-Recl")
                    
            End If

NEXTROW:
         If Err.Number > 0 Then
            Call Err.Clear
         End If
         rowcount = rowcount + 1
         Next row
         

        

    Next i
    
Exit Sub


End Sub

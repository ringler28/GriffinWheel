Sub manualEntry()

    Dim homeWB As Workbook
    Dim wmsResults As Workbook
    Dim wmsFile As String
    Dim resultsFolder As String
    Dim currentImg As String
    Dim imgFolder As String
    Dim imgFile As String
    Dim imgFileName As String
    Dim fullimgPath
    Dim row As Integer
    Dim line As Integer
    Dim offsetTimeMin As Integer
    Dim offsetTimeSec As Integer
    Dim dayDiff As Integer
    Dim hourDiff As Integer
    Dim minDiff As Integer
    Dim prevminDiff As Integer
    Dim missedCount As Integer
    Dim countMax As Integer
    Dim count As Integer
    Dim countPct As Integer
    Dim imgFS As Object
    Dim imgTime As Date
    Dim imgAdjTime As Date
    Dim wmsTime As Date
    Dim prevwmsTime As Date
    Dim i As Integer
    
    Set homeWB = Workbooks("RunningWMS_tapesize_comparison.xlsm")
    resultsFolder = homeWB.path & "\WMSresults"
    wmsFile = Dir(resultsFolder & "\*.csv")
    imgFolder = homeWB.path & "\OCRimages"
    
    'open each results overview file in WMSresults folder
    'Do While wmsFile <> ""
    
        ''choose results file to open''
           wmsFile = "20140707.csv"
        '''''''''''''''''''''''''''''''
        
        Workbooks.Open (resultsFolder & "\" & wmsFile)
        Set wmsResults = Workbooks(wmsFile)
    
    Do
        row = 2
        line = 2
        
        'iterate through each row in current opened wms log file
        Do
            'check the date modified of each ocr image file in image folder
            imgFile = Dir(imgFolder & "\*.bmp")
            Do While imgFile <> ""
        
                Set imgFS = CreateObject("Scripting.FileSystemObject")
                imgFileName = imgFS.GetFileName(imgFile)
                currentImg = "\" & imgFileName
                fullimgPath = imgFolder & currentImg
                Call lookupImgTime
            
                    wmsTime = wmsResults.Sheets(1).Range("M" & row)
                    prevwmsTime = wmsResults.Sheets(1).Range("M" & line)
                    dayDiff = DateDiff("d", wmsTime, imgAdjTime)
                    hourDiff = DateDiff("h", wmsTime, imgAdjTime)
                    minDiff = DateDiff("n", wmsTime, imgAdjTime)
                    prevminDiff = DateDiff("n", prevwmsTime, imgAdjTime)
                
                    'ensure time comparison is of same day, hour, and min
                    If dayDiff = 0 And hourDiff = 0 And minDiff = 0 And prevminDiff >= 0 Then
                        'compare time part (seconds) that image written time is after previous wms write time and before current
                        'If minDiff = 0 And prevminDiff = 0 Then
                            If DatePart("s", imgAdjTime) > DatePart("s", prevwmsTime) And DatePart("s", imgAdjTime) <= DatePart("s", wmsTime) Then
                                wmsResults.Sheets(1).Range("P" & row) = imgAdjTime
                                UserForm2.Image1.Picture = LoadPicture(fullimgPath)
                                UserForm2.TextBox2.Value = "0714"
                                UserForm2.TextBox3.Value = row
                                UserForm2.TextBox2.SetFocus
                                UserForm2.Show vbModal
                                SetAttr fullimgPath, vbNormal
                                Kill (fullimgPath)
                                    If i = 1 Then
                                        row = row + 1
                                        GoTo jump1
                                    ElseIf i = 0 Then
                                        GoTo jump2
                                    End If
                            End If
                        'ElseIf minDiff = 0 And prevminDiff > 0 Then
                            'If DatePart("s", imgAdjTime) <= DatePart("s", wmsTime) Then
                                'wmsResults.Sheets(1).Range("P" & row) = imgAdjTime
                               'UserForm2.Image1.Picture = LoadPicture(imgFolder & currentImg)
                               'UserForm2.Show vbModal
                            'End If
                        'ElseIf minDiff = -1 And prevminDiff = 0 Then
                            'If DatePart("s", imgAdjTime) > DatePart("s", prevwmsTime) Then
                                'wmsResults.Sheets(1).Range("P" & row) = imgAdjTime
                                'UserForm2.Image1.Picture = LoadPicture(imgFolder & currentImg)
                                'UserForm2.Show vbModal
                            'End If
                        'End If
                        
                    End If
jump2:
                    imgFile = Dir()
                Loop
            
                row = row + 1
                line = row - 1
                row = row - 1
                
jump1:
            Do
                row = row + 1
            Loop Until wmsResults.Sheets(1).Range("P" & row) = ""
            
        Loop Until wmsResults.Sheets(1).Range("A" & row) = ""
        
        
        ''save and close without alert''
        'wmsResults.SaveAs
        'wmsResults.Close False
        
        'wmsFile = Dir()
        
        UserForm3.Show vbModal
        
        row = 2
    Loop Until wmsResults.Sheets(1).Range("A" & row) = ""


End Sub

Sub lookupImgTime()

'input time offset from OCR image write time to WMS log file write time
offsetTimeMin = 7
offsetTimeSec = 29
    
    Set imgFS = CreateObject("Scripting.FileSystemObject")
    imgTime = imgFS.GetFile(imgFolder & currentImg).DateLastModified
    imgAdjTime = DateAdd("n", offsetTimeMin, imgTime)
    imgAdjTime = DateAdd("s", offsetTimeSec, imgAdjTime)
    Set imgFS = Nothing
    
End Sub



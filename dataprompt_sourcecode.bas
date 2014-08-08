
Sub openAll()

    Dim dataFolder As String
    Dim dataFile As String
    
    dataFolder = "CopyRawDataHere"
    
    'Change .xslm to proper file extension'
    dataFile = Dir(dataFolder & "\*.xlsm")

    Do While dataFile <> ""
        Workbooks.Open fileName:=dataFolder & "\" & dataFile
        dataFile = Dir()
    Loop

End Sub

Public Function fname() As String
    
    fname = ActiveWorkbook.Name

End Function

Sub scanSplit()

Dim newRun As Integer
Dim dataFolder2 As String
Dim dataFile2 As String
Dim dataWMS As Workbook
Dim resultsWMS As Workbook
Dim resultsView As Workbook

line = 10


'create counter for progress bar'
Call filecount
    countDone = 0
    countMax = Range("BA2")


'Define the name of the final viewing workbook'
    Set resultsView = Workbooks("dataprompt.xlsm")

    'Specify folder path where data is located'
    dataFolder2 = Range("BA1") & "\" & Range("BA5")
    
    '##Change .csv to proper file extension of files to be used'
    dataFile2 = Dir(dataFolder2 & "\*.csv")

    'Define name of calculation workbook'
    'Uses the same folder path that was entered for the "CopyRawDataHere" prompt'
    Set resultsWMS = Workbooks.Open(Range("BA1") & "\WMS_Calc_Wb.xlsm")

Do While dataFile2 <> ""
        
        Workbooks.Open fileName:=dataFolder2 & "\" & dataFile2
        Set dataWMS = Workbooks(dataFile2)
        
            'copy workbook name to results view'
            line = line + 1
            resultsView.Sheets(1).Range("A" & line) = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
               
      
            'use Evan Holcomb's code logic to split multiple scans'
            newRun = 3000
        
            Do
            
                newRun = newRun + 1
            
            Loop Until dataWMS.Sheets(1).Range("A" & newRun) = "Date" Or dataWMS.Sheets(1).Range("A" & newRun) = ""
        
            If dataWMS.Sheets(1).Range("A" & newRun) = "Date" Then
        
                newRun = newRun - 1
                dataWMS.Sheets(1).Range("$A$1:$Z" & newRun).Select
                Selection.EntireRow.Delete
                ActiveWorkbook.SaveAs fileName:=dataFolder2 & "\" & Range("A" & line) & "_II.csv"
        
            End If
            
              'count number of iterations performned'
            countDone = countDone + 1
            pctDone = countDone / countMax
            With main
                .Frame1.Caption = Format(pctDone, "0%")
                .Label16.Width = pctDone * (.Frame1.Width - 10)
            End With
            DoEvents
        
            dataWMS.Close False
        
        dataFile2 = Dir()
    Loop
    
resultsWMS.Close False

UserForm2.Show


End Sub
Sub scrapFinder()

  'Define all variables to be used in function'
    Dim dataFolder As String
    Dim dataFile As String
    Dim splitFolder As String
    Dim splitFile As String
    Dim dataWMS As Workbook
    Dim resultsWMS As Workbook
    Dim resultsView As Workbook
    Dim WB As Workbook
    Dim row As Integer
    Dim line As Integer
    Dim countMax As Integer
    Dim countDone As Integer
    Dim pctDone As Single
    Dim newRun As Integer
            
    
    'Set starting point for augmented assignment to be used in while loop'
    line = 10
    row = 10
        
    'create counter for progress bar'
    Call filecountScrap
        countDone = 0
        countMax = Range("BA4")
    
    'Define the name of the final viewing workbook'
    Set resultsView = Workbooks("dataprompt.xlsm")

    'Specify folder path where data is located'
    dataFolder = Range("BA1") & "\" & Range("BA5")
    
    '##Change .csv to proper file extension of files to be used'
    dataFile = Dir(dataFolder & "\GC*.csv")

    'Define name of calculation workbook'
    'Uses the same folder path that was entered for the "CopyRawDataHere" prompt'
    Set resultsWMS = Workbooks.Open(Range("BA1") & "\WMS_Calc_Wb.xlsm ")
    
    'copy and paste result headings from WMS results worksheet'
    resultsWMS.Sheets("Results").Range("$A$1:$AZ$1").Copy
    resultsView.Sheets(1).Range("B10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

    'clears previous compiled results'
    resultsView.Sheets(1).Range("$A$11:$AZ$500").ClearContents

    Do While dataFile <> ""
        
        Workbooks.Open fileName:=dataFolder & "\" & dataFile
        Set dataWMS = Workbooks(dataFile)
               
        'copy single scan raw data to calculation workbook'
         newRun = 3000
        
            Do
            
                newRun = newRun + 1
            
            Loop Until dataWMS.Sheets(1).Range("A" & newRun) = "Date" Or dataWMS.Sheets(1).Range("A" & newRun) = ""
        
            newRun = newRun - 1
            dataWMS.Sheets(1).Range("$A$1:$L" & newRun).Copy
            
            'copy workbook name to results view'
            line = line + 1
            resultsView.Sheets(1).Range("A" & line) = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
            
            'paste to WMS calculation raw data worksheet'
            resultsWMS.Sheets("Enter Raw Data Here").Range("A1").PasteSpecial
                                        
            'copy results from WMS results worksheet'
            resultsWMS.Sheets("Results").Range("$A$2:$AZ$2").Copy
        
            'paste data into data prompt workbook'
            row = row + 1
            resultsView.Sheets(1).Range("B" & row).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                    
        'count number of iterations performned'
        countDone = countDone + 1
        pctDone = countDone / countMax
        With main
            .Frame1.Caption = Format(pctDone, "0%")
            .Label16.Width = pctDone * (.Frame1.Width - 10)
        End With
        DoEvents
        
        'clear results copied into results tab on calc wb so data from past scan is not merged into new'
        resultsWMS.Sheets("Enter Raw Data Here").Range("$A$3000:$L" & newRun).ClearContents
         
        dataWMS.Close False
        
        dataFile = Dir()
    Loop


'closes calculation workbook'
resultsWMS.Close False


'sets dataprompt results view as active workbook'
resultsView.Sheets(1).Activate

main.Hide

UserForm1.Show


End Sub

Sub validNumFinder()

  'Define all variables to be used in function'
    Dim dataFolder As String
    Dim dataFile As String
    Dim splitFolder As String
    Dim splitFile As String
    Dim dataWMS As Workbook
    Dim resultsWMS As Workbook
    Dim resultsView As Workbook
    Dim WB As Workbook
    Dim row As Integer
    Dim line As Integer
    Dim countMax As Integer
    Dim countDone As Integer
    Dim pctDone As Single
    Dim newRun As Integer
    Dim numLen As Integer
    Dim fileName As String
    Dim i As Integer
    Dim j As String
    Dim k As Object
    Dim h As String
            
    
    'Set starting point for augmented assignment to be used in while loop'
    line = 10
    row = 10
        
    'create counter for progress bar'
    Call filecountValid
        countDone = 0
        countMax = Range("BA4")
        
Application.ScreenUpdating = False
    
    'Define the name of the final viewing workbook'
    Set resultsView = Workbooks("dataprompt.xlsm")

    'Specify folder path where data is located'
    dataFolder = Range("BA1") & "\" & Range("BA5")
    
    '##Change .csv to proper file extension of files to be used'
    dataFile = Dir(dataFolder & "\GC*.csv")

    'Define name of calculation workbook'
    'Uses the same folder path that was entered for the "CopyRawDataHere" prompt'
    Set resultsWMS = Workbooks.Open(Range("BA1") & "\WMS_Calc_Wb.xlsm ")
    
    'copy and paste result headings from WMS results worksheet'
    resultsWMS.Sheets("Results").Range("$A$1:$AZ$1").Copy
    resultsView.Sheets(1).Range("B10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False

    'clears previous compiled results'
    resultsView.Sheets(1).Range("$A$11:$AZ$500").ClearContents

Do While dataFile <> ""
               
        'check to see if filename is valid length
        Set k = CreateObject("Scripting.FileSystemObject")
        h = k.GetFileName(dataFile)
        j = Left(h, InStr(h, ".") - 1)
        i = Len(Replace(j, " ", ""))
        
    If i = 11 Then
                
        Workbooks.Open fileName:=dataFolder & "\" & dataFile
        Set dataWMS = Workbooks(dataFile)

        'copy single scan raw data to calculation workbook'
         newRun = 3000
        
            Do
            
                newRun = newRun + 1
            
            Loop Until dataWMS.Sheets(1).Range("A" & newRun) = "Date" Or dataWMS.Sheets(1).Range("A" & newRun) = ""
        
            newRun = newRun - 1
            dataWMS.Sheets(1).Range("$A$1:$L" & newRun).Copy
            
            'copy workbook name to results view'
            line = line + 1
            resultsView.Sheets(1).Range("A" & line) = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
            
            'paste to WMS calculation raw data worksheet'
            resultsWMS.Sheets("Enter Raw Data Here").Range("A1").PasteSpecial
                                        
            'copy results from WMS results worksheet'
            resultsWMS.Sheets("Results").Range("$A$2:$AZ$2").Copy
        
            'paste data into data prompt workbook'
            row = row + 1
            resultsView.Sheets(1).Range("B" & row).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                    
        'count number of iterations performned'
        countDone = countDone + 1
        pctDone = countDone / countMax
        With main
            .Frame1.Caption = Format(pctDone, "0%")
            .Label16.Width = pctDone * (.Frame1.Width - 10)
        End With
        DoEvents
        
        'clear results copied into results tab on calc wb so data from past scan is not merged into new'
        resultsWMS.Sheets("Enter Raw Data Here").Range("$A$3000:$L" & newRun).ClearContents
         
        dataWMS.Close False
        
    End If

        dataFile = Dir()
Loop


'closes calculation workbook'
resultsWMS.Close False

Application.ScreenUpdating = True

'sets dataprompt results view as active workbook'
resultsView.Sheets(1).Activate

main.Hide

UserForm1.Show


End Sub



Sub copydataAll()


    'Define all variables to be used in function'
    Dim dataFolder As String
    Dim dataFile As String
    Dim splitFolder As String
    Dim splitFile As String
    Dim dataWMS As Workbook
    Dim resultsWMS As Workbook
    Dim resultsView As Workbook
    Dim WB As Workbook
    Dim row As Integer
    Dim line As Integer
    Dim countMax As Integer
    Dim countDone As Integer
    Dim pctDone As Single
    Dim newRun As Integer
            
    
    'Set starting point for augmented assignment to be used in while loop'
    line = 10
    row = 10
        
    'create counter for progress bar'
    Call filecount
        countDone = 0
        countMax = Range("BA2")
    
    'Define the name of the final viewing workbook'
    Set resultsView = Workbooks("dataprompt.xlsm")

    'Specify folder path where data is located'
    dataFolder = Range("BA1") & "\" & Range("BA5")
    
    '##Change .csv to proper file extension of files to be used'
    dataFile = Dir(dataFolder & "\*.csv")

    'Define name of calculation workbook'
    'Uses the same folder path that was entered for the "CopyRawDataHere" prompt'
    Set resultsWMS = Workbooks.Open(Range("BA1") & "\WMS_Calc_Wb.xlsm ")
    
    'copy and paste result headings from WMS results worksheet'
    resultsWMS.Sheets("Results").Range("$A$1:$AZ$1").Copy
    resultsView.Sheets(1).Range("B10").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
    'clears previous compiled results'
    resultsView.Sheets(1).Range("$A$11:$AZ$500").ClearContents

    Do While dataFile <> ""
        
        Workbooks.Open fileName:=dataFolder & "\" & dataFile
        Set dataWMS = Workbooks(dataFile)
               
        'copy single scan raw data to calculation workbook'
         newRun = 3000
        
            Do
            
                newRun = newRun + 1
            
            Loop Until dataWMS.Sheets(1).Range("A" & newRun) = "Date" Or dataWMS.Sheets(1).Range("A" & newRun) = ""
        
            newRun = newRun - 1
            dataWMS.Sheets(1).Range("$A$1:$L" & newRun).Copy
            
            'copy workbook name to results view'
            line = line + 1
            resultsView.Sheets(1).Range("A" & line) = Left(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 1)
            
            'paste to WMS calculation raw data worksheet'
            resultsWMS.Sheets("Enter Raw Data Here").Range("A1").PasteSpecial
                                        
            'copy results from WMS results worksheet'
            resultsWMS.Sheets("Results").Range("$A$2:$AZ$2").Copy
        
            'paste data into data prompt workbook'
            row = row + 1
            resultsView.Sheets(1).Range("B" & row).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                    
        'count number of iterations performned'
        countDone = countDone + 1
        pctDone = countDone / countMax
        With main
            .Frame1.Caption = Format(pctDone, "0%")
            .Label16.Width = pctDone * (.Frame1.Width - 10)
        End With
        DoEvents
        
        'clear results copied into results tab on calc wb so data from past scan is not merged into new'
        resultsWMS.Sheets("Enter Raw Data Here").Range("$A$3000:$L" & newRun).ClearContents
         
        dataWMS.Close False
        
        dataFile = Dir()
    Loop


'closes calculation workbook'
resultsWMS.Close False


'sets dataprompt results view as active workbook'
resultsView.Sheets(1).Activate

main.Hide

UserForm1.Show

End Sub

Sub clearSplits()

'create subroutine to clear split folder after every mass pull
Kill (splitFolder & "\*.csv")


End Sub

Sub clearData()
'clear results copied into results tab on calc wb so data from past scan is not merged into new'
resultsWMS.Sheets("Enter Raw Data Here").Range("$A$4:$Z$4000").Delete xlUp

End Sub

Sub filecount()
 
 Dim dataFolder As String, path As String, count As Integer
    dataFolder = Range("BA1") & "\" & Range("BA5")

    'Change .csv to proper file extension'
    path = dataFolder & "\*.csv"

    dataFile = Dir(path)

    Do While dataFile <> ""
        count = count + 1
        dataFile = Dir()
    Loop

    Range("BA2").Value = count
    
End Sub

Sub filecountScrap()
 
 Dim dataFolder As String, path As String, count As Integer
    dataFolder = Range("BA1") & "\" & Range("BA5")

    'Change .csv to proper file extension'
    path = dataFolder & "\GC*.csv"

    dataFile = Dir(path)

    Do While dataFile <> ""
        count = count + 1
        dataFile = Dir()
    Loop

    Range("BA4").Value = count
    
End Sub
Sub filecountValid()
 
 Dim dataFolder As String, path As String, count As Integer
 Dim i As Integer
 Dim j As String
 Dim k As Object
 Dim h As String

    dataFolder = Range("BA1") & "\" & Range("BA5")

    'Change .csv to proper file extension'
    path = dataFolder & "\GC*.csv"

    dataFile = Dir(path)

    Do While dataFile <> ""
        Set k = CreateObject("Scripting.FileSystemObject")
        h = k.GetFileName(dataFile)
        j = Left(h, InStr(h, ".") - 1)
        i = Len(Replace(j, " ", ""))
            If i = 11 Then
                count = count + 1
            End If
            
        dataFile = Dir()
    Loop

    Range("BA4").Value = count

End Sub

Sub copydata()


Dim dataWMS As Workbook
Dim resultsWMS As Workbook
Dim row As Integer

'## Open both workbooks:
Set dataWMS = Workbooks.Open(Range("AA1") & ".xlsm")
Set resultsWMS = Workbooks.Open("WMS_spreadsheet.xlsm ")

'copy data from WMS'
dataWMS.Sheets(1).Range("$A$1:$C$3").Copy

'paste to WMS results worksheet'
row = 2
resultsWMS.Sheets("Enter Raw Data Here").Range("A" & row).PasteSpecial

'Close x:
'Specify False to close without saving'

dataWMS.Close False

End Sub

Sub copyresults()


Dim resultView As Workbook
Dim resultsWMS As Workbook
Dim row As Integer

'## Open both workbooks:
Set resultView = Workbooks("dataprompt.xlsm")
Set resultsWMS = Workbooks("WMS_spreadsheet.xlsm ")


'copy results from WMSspreadsheet'
resultsWMS.Sheets("Results").Range("$A$1:$Q$2").Copy

'paste to WMS results worksheet'
row = 1
row = row + 1
resultsWView.Sheets(1).Range("A" & row).PasteSpecial

'Close x:
resultsWMS.Close

End Sub


Sub GoToChart()

    Dim gRow As Integer
    Call filecount
    gRow = Range("AA2") + 10
    wnum = Range("AA2")
    
    ActiveWorkbook.Charts.Add
    With ActiveWorkbook.ActiveChart
        'Data?
        .ChartType = xlXYScatter
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "=""Diameter"""
        .SeriesCollection(1).XValues = "=Sheet1!$C$11:$C" & gRow
        .SeriesCollection(1).Values = "=Sheet1!$AB$11:$B" & gRow

     
        'Titles
        .HasTitle = True
        .ChartTitle.Characters.Text = "Scatter Chart"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "X values"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Y values"
        .Axes(xlCategory).HasMajorGridlines = True

        'Formatting
        .Axes(xlCategory).HasMinorGridlines = False
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).HasMinorGridlines = False
        .HasLegend = False

    End With
     

End Sub




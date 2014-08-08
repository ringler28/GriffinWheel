Sub filecount()
 
 Dim dataFolder As String, path As String, count As Integer
    dataFolder = ThisWorkbook.path & "\OCRimages"

    'Change .bmp to proper file extension'
    path = dataFolder & "\*.bmp"

    dataFile = Dir(path)

    Do While dataFile <> ""
        count = count + 1
        dataFile = Dir()
    Loop
    
    UserForm1.TextBox1 = count
    UserForm1.Show
    
End Sub

Sub dataCombine()

Dim row As Integer
Dim line As Integer

Application.ScreenUpdating = False

row = 2

Do

line = 2

    Do

        If Worksheets("WMSdata").Range("A" & line) = Worksheets("SQL_WMScomparison").Range("A" & row) Then
    
            Worksheets("WMSdata").Range("B" & line).Copy
            Worksheets("SQL_WMScomparison").Range("C" & row).PasteSpecial
            
            'Copy wms diameter
            'Worksheets("WMSdata").Range("D" & line).Copy
            'Worksheets("SQL_WMScomparison").Range("D" & row).PasteSpecial
            
            'Copy date column
            'Worksheets("WMSdata").Range("C" & line).Copy
            'Worksheets("SQL_WMScomparison").Range("E" & row).PasteSpecial
            
        ElseIf Worksheets("WMSdata").Range("A" & line) <> Worksheets("SQL_WMScomparison").Range("A" & row) Then
        
        End If
        
        line = line + 1
    
    Loop Until Worksheets("WMSdata").Range("A" & line).Value = ""

    row = row + 1
    
Loop Until Worksheets("SQL_WMScomparison").Range("A" & row).Value = ""

Application.ScreenUpdating = True


End Sub


Sub cleanData()
    Dim row As Integer
    Dim line As Integer

 row = 2
 
 'ActiveWorkbook.Worksheets("SQL_WMScomparison").Activate

    Do
        If Worksheets("SQL_WMScomparison").Range("B" & row) = 0 Or Worksheets("SQL_WMScomparison").Range("C" & row) = 0 Or Worksheets("SQL_WMScomparison").Range("D" & row) > 5 Then
            Worksheets("SQL_WMScomparison").Range("A" & row).EntireRow.Delete
            GoTo jump1
        End If
    
    row = row + 1
    
jump1:
    Loop Until Worksheets("SQL_WMScomparison").Range("A" & row) = ""

End Sub

Sub showChart()
       ActiveWorkbook.Charts("PLOT-Tape Size Variance").Activate
End Sub



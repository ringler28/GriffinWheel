Sub m9or6()
    
Dim row As Integer
Dim line As Integer
Dim count9or6 As Integer
Dim countPass9or6 As Integer
Dim pctTotal9or6 As Integer
Dim countPass As Integer
Dim countTotal As Integer
Dim pctFail9or6 As Integer

row = 1
line = 1

    Sheets("9or6").Select
    Columns("A:A").Select
    Range("A64").Activate
    Selection.Delete Shift:=xlToLeft
    Sheets("OCR_Validation_At_Final").Select

    Do While Workbooks("OCR_Validation_Macro.xlsm").Sheets("OCR_Validation_At_Final").Range("A" & row) <> ""
       
         row = row + 1
        If InStr(1, Workbooks("OCR_Validation_Macro.xlsm").Sheets("OCR_Validation_At_Final").Range("B" & row).Value, "PLC WN", vbTextCompare) > 0 And InStr(1, Workbooks("OCR_Validation_Macro.xlsm").Sheets("OCR_Validation_At_Final").Range("B" & row).Value, "PLC WN = 00000", vbTextCompare) = 0 Then
            If InStr(Workbooks("OCR_Validation_Macro.xlsm").Sheets("OCR_Validation_At_Final").Range("B" & row), "9") > 0 Or InStr(Workbooks("OCR_Validation_Macro.xlsm").Sheets("OCR_Validation_At_Final").Range("B" & row), "6") > 0 Then
                Workbooks("OCR_Validation_Macro.xlsm").Sheets("OCR_Validation_At_Final").Range("B" & row).Select
                Selection.Copy
                Workbooks("OCR_Validation_Macro.xlsm").Sheets("9or6").Range("A" & line).PasteSpecial xlPasteValues
                line = line + 1
            End If
        End If
         
        
    Loop
    
line = 1
count9or6 = 0
    
    Do While Workbooks("OCR_Validation_Macro.xlsm").Sheets("9or6").Range("A" & line) <> ""
        count9or6 = count9or6 + 1
        line = line + 1
    Loop
    
    ActiveWindow.ScrollRow = 1
    
    Sheets("Main_Page").Select
    
    Range("C4").Value = count9or6
    
    MsgBox count9or6
End Sub

Sub m0()
Dim i As Integer

Sheets("OCR_Validation_At_Final").Select

    Range("Table1[[#Headers],[message]]").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=2, Criteria1:= _
        "=*00000*", Operator:=xlAnd
    Range("B2:B10000").Select
    Selection.Copy
    Sheets("00000").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("OCR_Validation_At_Final").Select
    Range("Table1[[#Headers],[message]]").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=2
    
    i = 1

    Do While Sheets("00000").Range("A" & i) <> ""
        
        i = i + 1
        
    Loop
    
i = i - 1

Sheets("Main_Page").Select

Range("C3").Value = i

MsgBox i
End Sub

Sub sucm()
   Dim i As Integer
   
   
   Sheets("OCR_Validation_At_Final").Select
   
    Range("Table1[[#Headers],[message]]").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=2, Criteria1:= _
        "=*Validated against PLC.*", Operator:=xlAnd
    Range("B2:B10000").Select
    Selection.Copy
    Sheets("Pass").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("OCR_Validation_At_Final").Select
    Application.CutCopyMode = False
    Range("Table1[[#Headers],[message]]").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=2
    
 

i = 1

    Do While Sheets("Pass").Range("A" & i) <> ""
        
        i = i + 1
        
    Loop
    
i = i - 1

Sheets("Main_Page").Select

Range("C5").Value = i

MsgBox i
    
End Sub

Sub totm()


Dim i As Integer
Dim j As Integer

i = 2
j = 0

    Do While Sheets("OCR_Validation_At_Final").Range("A" & i) <> ""
        
        i = i + 1
        j = j + 1
        
    Loop

Sheets("Main_Page").Select

Range("C2").Value = j

MsgBox j

End Sub



Sub find9or6()
    
Dim row As Integer
Dim line As Integer
Dim count9or6 As Integer
Dim countPass9or6 As Integer
Dim pctTotal9or6 As Integer
Dim countPass As Integer
Dim countTotal As Integer
Dim pctFail9or6 As Integer

row = 1
line = 1

    Sheets("9or6").Select
    Columns("A:A").Select
    Range("A64").Activate
    Selection.Delete Shift:=xlToLeft
    Sheets("OCR_Validation_At_Final").Select

    Do While Workbooks("OCR_Validation_Macro.xlsm").Sheets("OCR_Validation_At_Final").Range("A" & row) <> ""
       
         row = row + 1
        If InStr(1, Workbooks("OCR_Validation_Macro.xlsm").Sheets("OCR_Validation_At_Final").Range("B" & row).Value, "PLC WN", vbTextCompare) > 0 And InStr(1, Workbooks("OCR_Validation_Macro.xlsm").Sheets("OCR_Validation_At_Final").Range("B" & row).Value, "PLC WN = 00000", vbTextCompare) = 0 Then
            If InStr(Workbooks("OCR_Validation_Macro.xlsm").Sheets("OCR_Validation_At_Final").Range("B" & row), "9") > 0 Or InStr(Workbooks("OCR_Validation_Macro.xlsm").Sheets("OCR_Validation_At_Final").Range("B" & row), "6") > 0 Then
                Workbooks("OCR_Validation_Macro.xlsm").Sheets("OCR_Validation_At_Final").Range("B" & row).Select
                Selection.Copy
                Workbooks("OCR_Validation_Macro.xlsm").Sheets("9or6").Range("A" & line).PasteSpecial xlPasteValues
                line = line + 1
            End If
        End If
         
        
    Loop
    
line = 1
count9or6 = 0
    
    Do While Workbooks("OCR_Validation_Macro.xlsm").Sheets("9or6").Range("A" & line) <> ""
        count9or6 = count9or6 + 1
        line = line + 1
    Loop
    
    ActiveWindow.ScrollRow = 1
    
    Sheets("Main_Page").Select
    
    Range("C4").Value = count9or6
    
    'MsgBox count9or6
        
End Sub

Sub countSuccess()
Attribute countSuccess.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Button11_Click Macro
'

'
   Dim i As Integer
   
   Sheets("OCR_Validation_At_Final").Select

    Range("Table1[[#Headers],[message]]").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=2, Criteria1:= _
        "=*Validated against PLC.*", Operator:=xlAnd
    Range("B2:B10000").Select
    Selection.Copy
    Sheets("Pass").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("OCR_Validation_At_Final").Select
    Application.CutCopyMode = False
    Range("Table1[[#Headers],[message]]").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=2
    
 

i = 1

    Do While Sheets("Pass").Range("A" & i) <> ""
        
        i = i + 1
        
    Loop
    
i = i - 1

Sheets("Main_Page").Select

Range("C5").Value = i

'MsgBox i
    
    
End Sub

Sub countZero()
Attribute countZero.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Button13_Click Macro
'

'
Dim i As Integer

Sheets("OCR_Validation_At_Final").Select

    Range("Table1[[#Headers],[message]]").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=2, Criteria1:= _
        "=*00000*", Operator:=xlAnd
    Range("B2:B10000").Select
    Selection.Copy
    Sheets("00000").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("OCR_Validation_At_Final").Select
    Range("Table1[[#Headers],[message]]").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=2
    
    i = 1

    Do While Sheets("00000").Range("A" & i) <> ""
        
        i = i + 1
        
    Loop
    
i = i - 1

Sheets("Main_Page").Select

Range("C3").Value = i

'MsgBox i
    
End Sub

Sub transferData()
Attribute transferData.VB_ProcData.VB_Invoke_Func = " \n14"
'

'

Dim i As Integer

    
    Sheets("OCR_Validation_At_Final").Select
    Range("A1").Select
    
If Range("A1") <> "" Then
    Selection.ListObject.ListColumns(2).Delete
    Range("A2").Select
    Selection.ListObject.ListColumns(1).Delete
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 21.14
    Columns("B:B").ColumnWidth = 81.43
ElseIf Range("A1") = "" Then
End If
    ChDir Range("folderName")
    Workbooks.Open Filename:= _
        Range("folderName") & "\" & Range("fileName")
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    Rows("6:6").RowHeight = 22.5
    Rows("1:6").Select
    Range("A6").Activate
    Selection.Delete Shift:=xlUp
    
i = 1
    
    Do While Range("A" & i) <> ""
    
        i = i + 1
    
    Loop
    
i = i - 1

    Range("A1:H" & i).Select
    Selection.UnMerge
    Columns("B:C").Select
    Selection.Delete Shift:=xlToLeft
    Selection.ColumnWidth = 84.71
    Range("A1:B" & i).Select
    Selection.Copy
    Windows("OCR_Validation_Macro.xlsm").Activate
    ActiveSheet.Paste
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$B" & i), , xlYes).Name = _
        "Table1"
    Range("Table1[#All]").Select
    ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleLight1"
    
    Application.DisplayAlerts = False
    
    Windows("OCR_Validation_At_Final.xlsx").Activate
    ActiveWindow.Close False

    Application.DisplayAlerts = True

    Sheets("Main_Page").Select

End Sub


Sub countTotal()

Dim i As Integer
Dim j As Integer

i = 2
j = 0

    Do While Sheets("OCR_Validation_At_Final").Range("A" & i) <> ""
        
        i = i + 1
        j = j + 1
        
    Loop

Sheets("Main_Page").Select

Range("C2").Value = j

'MsgBox j
    
End Sub


Sub removeDuplicates()
'
' removeDuplicates Macro
'

'

Dim k As Integer
Dim l As Integer

Sheets("OCR_Validation_At_Final").Select

l = 2
Do While Range("A" & l) <> ""
l = l + 1
Loop

    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-1],17,11)"
    
    Range("Table1").Select
    Selection.Copy
    Sheets.Add After:=Sheets(Sheets.count)
    Columns("B:B").ColumnWidth = 113.71
    Columns("A:A").ColumnWidth = 25
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Columns("C:C").Select
    ActiveSheet.Range("$A$1:$C" & l).removeDuplicates Columns:=3, Header:=xlNo

    Sheets("OCR_Validation_At_Final").Select
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 20.86
    Columns("B:B").ColumnWidth = 74.71
    Columns("B:B").ColumnWidth = 77.14
    Sheets(6).Select
    
k = 1
Do While Range("A" & k) <> ""
    k = k + 1
Loop

    Range("A1:B" & k).Select
    Selection.Copy
    Sheets("OCR_Validation_At_Final").Select
    Range("A2").Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "date"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "message"
    Range("A1:B" & k).Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$B" & k), , xlYes).Name = _
        "Table1"
    Range("Table1[#All]").Select
    ActiveSheet.ListObjects("Table1").TableStyle = "TableStyleMedium2"
    
        Range("Table1[#Headers]").Select
    With Selection.Font
        .Name = "Tahoma"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With

    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
        Range("Table1").Select
    With Selection.Interior
        .PatternThemeColor = xlThemeColorDark1
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = -0.149998474074526
    End With
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=MOD(ROW(),2)"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249946592608417
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    
    Sheets(6).Select
    
    Application.DisplayAlerts = False
    ActiveWindow.SelectedSheets.Delete
    Sheets("OCR_Validation_At_Final").Select
    Application.DisplayAlerts = True
    
    Sheets("Main_Page").Select
    
End Sub

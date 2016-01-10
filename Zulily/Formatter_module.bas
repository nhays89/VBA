Attribute VB_Name = "Module1"
' Keyboard Shortcut: Ctrl+Shift+V
Dim ncolns As Integer
Dim nrows As Integer
Dim i As Integer
Dim lastrow As Integer
Dim poqty As Long
Dim pounder As Long
Dim poover As Long
Dim qtyrcvd As Long
Dim j As Integer
Dim s1 As Range
Dim rangenamefound As Boolean

Sub main()

Dim rangeVal As String
rangeVal = FindRange()
Call Formatter(rangeVal)
End Sub

Function FindRange() As String
Dim initialproxy As String
With Range("A1")
rangenamefound = False

ncolns = Range(.Cells(1, 1), .End(xlToRight)).Columns.Count

  Do Until rangenamefound
    For j = 1 To ncolns
        If .Cells(1, j).Value Like "PO Pushed" Then
        rangenamefound = True
        initialproxy = .Cells(2, j).Address
        Exit Do
        End If
    Next j
   Loop
End With
FindRange = initialproxy
End Function

Sub Formatter(rangeVal As String)

With Range(rangeVal)
lastrow = .Cells.SpecialCells(xlCellTypeLastCell).Row
lastrow = lastrow - 1
End With
Set s1 = ActiveWorkbook.ActiveSheet.Range("I2").Cells(1, 1)
With Range(rangeVal)
    For i = lastrow To 1 Step -1
    If .Cells(i, 1).Value = "" Then
    poqty = s1.Cells(i, 10)
    rcvdqty = s1.Cells(i, 12)
    pounder = s1.Cells(i, 17)
    poover = s1.Cells(i, 16)
    s1.Cells(i, 10) = rcvdqty
    s1.Cells(i, 12) = poqty
    s1.Cells(i, 17) = rcvdqty
    s1.Cells(i, 16) = poqty
    End If
    Next i
End With
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("G:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Columns("K:K").Select
    Selection.Delete Shift:=xlToLeft
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Columns("M:M").Select
    Selection.Cut
    Columns("L:L").Select
    Selection.Insert Shift:=xlToRight
    Columns("N:N").Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Columns("I:M").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1:M1").Select
    Selection.Font.Bold = True
    Columns("A:M").Select
    Selection.AutoFilter
    Columns("L:L").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Columns("M:M").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
        Formula1:="=0.1", Formula2:="=10000"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16751204
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("B2").Select
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("B2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("L1:L107"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A2").Select
End Sub



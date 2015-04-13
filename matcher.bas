Attribute VB_Name = "Module2"
Sub emmanuel()
Attribute emmanuel.VB_ProcData.VB_Invoke_Func = "N\n14"

Dim rcvdqty As Long
Dim overageqty As Long
Dim isfound As Boolean
Dim counter1 As Integer
Dim counter2 As Integer
Dim nrows As Integer
Dim c1 As Range
Dim b2 As Range
Dim z As Integer

Set c1 = Worksheets("Sheet1").Range("c2")
Set b2 = Worksheets("Sheet2").Range("b2")
If b2.Cells(1, 1).Value = "" Then
MsgBox ("please insert data into cells " & Cells("A1") & "")
End If


With c1
nrows = Range(.Cells(1, 1), .End(xlDown)).Rows.Count
nrows = nrows + 1
End With

counter1 = 1
    Do Until b2.Cells(counter1, 1).Value = ""
    isfound = False
    counter2 = 1
        Do Until isfound = True
            For counter2 = 1 To nrows
            overageqty = 0
            rcvdqty = 0
            z = 0
            z = c1.Cells(counter2, 11).Value
                    If c1.Cells(counter2, 1).Value = b2.Cells(counter1, 1).Value And z = 0 Then
                    isfound = True
                    rcvdqty = b2.Cells(counter1, 3).Value + c1.Cells(counter2, 9).Value
                    c1.Cells(counter2, 9).Value = rcvdqty
                    overageqty = b2.Cells(counter1, 3).Value
                    c1.Cells(counter2, 11).Value = overageqty
                    Exit Do
                    ElseIf c1.Cells(counter2 + 1, 1) = "" Then
                    isfound = True
                    Exit Do
                    Else
                    End If
            Next
        Loop
        counter1 = counter1 + 1
    Loop
End Sub

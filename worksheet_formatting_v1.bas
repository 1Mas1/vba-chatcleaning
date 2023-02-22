Attribute VB_Name = "worksheet_formatting"

Sub worksheet_formatting_main()

Call clearFilter
Call removeFilter
Call deleteFirstRow
Call convertToRange
Call ArrangeColumns
Call splitDateTime
Call columnWidth
Call formatting
Call applyFilter
Call freezeRow

End Sub

Sub ArrangeColumns()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("MainSheet")

ws.Range("A1").EntireColumn.Insert
ws.Range("A1").value = "Chat Index"

Dim lastColumn As Long
lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column

Dim column As Range
For Each column In ws.Range("A1:" & Split(ws.Cells(1, lastColumn).Address, "$")(1) & "1").Cells
    If column.value = "Timestamp: Time" Then
        column.EntireColumn.Cut
        ws.Range("B1").Insert shift:=xlToRight
        ws.Range("B1").value = "Date"
        lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column
        Exit For
    End If
Next column

ws.Range("C1").EntireColumn.Insert
ws.Range("C1").value = "Time"

ws.Range("D1").EntireColumn.Insert
ws.Range("D1").value = "Blank"

For Each column In ws.Range("A1:" & Split(ws.Cells(1, lastColumn).Address, "$")(1) & "1").Cells
    If column.value = "Body" Then
        column.EntireColumn.Cut
        ws.Range("E1").Insert shift:=xlToRight
        lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column
        Exit For
    End If
Next column

For Each column In ws.Range("A1:" & Split(ws.Cells(1, lastColumn).Address, "$")(1) & "1").Cells
    If column.value = "From" Then
        column.EntireColumn.Cut
        ws.Range("F1").Insert shift:=xlToRight
        ws.Range("F1").value = "From User"
        lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column
        Exit For
    End If
Next column

ws.Range("G1").EntireColumn.Insert
ws.Range("G1").value = "From Attributed"

For Each column In ws.Range("A1:" & Split(ws.Cells(1, lastColumn).Address, "$")(1) & "1").Cells
    If column.value = "To" Then
        column.EntireColumn.Cut
        ws.Range("H1").Insert shift:=xlToRight
        ws.Range("H1").value = "To User"
        lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column
        Exit For
    End If
Next column

ws.Range("I1").EntireColumn.Insert
ws.Range("I1").value = "To Attributed"

For Each column In ws.Range("A1:" & Split(ws.Cells(1, lastColumn).Address, "$")(1) & "1").Cells
    If column.value = "Participants" Then
        column.EntireColumn.Cut
        ws.Range("J1").Insert shift:=xlToRight
        lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column
        Exit For
    End If
Next column

For Each column In ws.Range("A1:" & Split(ws.Cells(1, lastColumn).Address, "$")(1) & "1").Cells
    If column.value = "Source" Then
        column.EntireColumn.Cut
        ws.Range("K1").Insert shift:=xlToRight
        lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column
        Exit For
    End If
Next column

End Sub

Sub deleteFirstRow()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("MainSheet")

ws.Rows(1).Delete

End Sub

Sub convertToRange()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("MainSheet")

Dim table As ListObject
For Each table In ws.ListObjects
    table.Range.Copy
    table.Unlist
Next table

Application.CutCopyMode = False

End Sub

Sub clearFilter()
'removes all filters on activesheet

On Error Resume Next
ActiveSheet.ShowAllData

End Sub

Sub removeFilter()
'removes the filter from the activesheet

On Error Resume Next
ActiveSheet.AutoFilterMode = False

End Sub

Sub formatting()
'if this sub is called after cleaning the columns, then the chat index will be blank. This uses the column titled '#' to find the lastrow

Dim lastrow As Long
Dim lastColumn As Long
Dim col As Range


Set col = Rows(1).Find("#", LookIn:=xlValues, lookat:=xlWhole)

lastrow = Cells(Rows.Count, col.column).End(xlUp).Row
lastColumn = Cells(1, Columns.Count).End(xlToLeft).column

Dim rngAll As Range
Set rngAll = Range(Cells(1, 1), Cells(lastrow, lastColumn))

Dim rngTopRow As Range
Set rngTopRow = Range(Cells(1, 1), Cells(1, lastColumn))

Dim rngSecondRowDown As Range
Set rngSecondRowDown = Range(Cells(2, 1), Cells(lastrow, lastColumn))

With rngAll
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    .Borders.ColorIndex = xlAutomatic
    
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Borders(xlInsideHorizontal).Weight = xlThin
    .Borders(xlInsideHorizontal).ColorIndex = xlAutomatic
    
    .Borders(xlInsideVertical).LineStyle = xlContinuous
    .Borders(xlInsideVertical).Weight = xlThin
    .Borders(xlInsideVertical).ColorIndex = xlAutomatic
End With

'sets the colour, font and row size of the first row
With rngTopRow
    .Interior.Color = RGB(48, 84, 150)
    .Font.Color = vbWhite
    .Font.Bold = True
    .RowHeight = 40
End With

'sets colour, borders and row size of rows 2 to lastrow
With rngSecondRowDown
    .Interior.Color = RGB(255, 255, 255)
    .RowHeight = 50
End With

'wraps text in columns E to I
Range("E:I").WrapText = True

'sets font to calibri size 11
rngAll.Font.Name = "Calibri"
rngAll.Font.Size = 11

End Sub

Sub splitDateTime()
'if this sub is called after cleaning the columns, then the chat index will be blank. This uses the column titled '#' to find the lastrow

Dim lastrow As Long
Dim col As Range

Set col = Rows(1).Find("#", LookIn:=xlValues, lookat:=xlWhole)
lastrow = Cells(Rows.Count, col.column).End(xlUp).Row

For i = 2 To lastrow
    Cells(i, 3).value = Mid(Cells(i, 2).value, 12, 16)
    Cells(i, 2).value = Left(Cells(i, 2).value, 10)
Next i

End Sub

Sub columnWidth()

Columns("a").columnWidth = 15
Columns("b").columnWidth = 11
Columns("c:d").columnWidth = 15
Columns("e").columnWidth = 55
Columns("f:i").columnWidth = 20
Columns("j").columnWidth = 40


End Sub

Sub applyFilter()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("MainSheet")
Dim rngAll As Range

Dim lastrow As Long
Dim lastColumn As Long
Dim col As Range

Set col = Rows(1).Find("#", LookIn:=xlValues, lookat:=xlWhole)

lastrow = Cells(Rows.Count, col.column).End(xlUp).Row
lastColumn = Cells(1, Columns.Count).End(xlToLeft).column

Set rngAll = Range(Cells(1, 1), Cells(lastrow, lastColumn))
rngAll.AutoFilter

End Sub

Sub freezeRow()
'freeze top row

ActiveWindow.SplitRow = 1
ActiveWindow.FreezePanes = True

End Sub


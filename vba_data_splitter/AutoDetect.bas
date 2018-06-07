Attribute VB_Name = "AutoDetect"

Sub AutoDetect()

'Splitting the data is done based on values in a specific column. This code collects all values occuring in the column from all sheets.

Dim i As Long 'general variable used by for loops
Dim pointer As Long 'the pointer points at the next empty cell in the values list on the info sheet
Dim split_col As String 'name of the column the splitting is based on
Dim split_col_num As Long 'the number of the column the splitting is based on
Dim lines_count As Long 'number of lines on a sheet

'to avoid error messages stopping the script
Application.DisplayAlerts = False

'clearing the values range before populating it with the values collected from the split column
Sheets("info").Range("A14:A2000").Clear
Sheets("info").Range("A14:A2000").Interior.Color = RGB(255, 242, 204)
Sheets("info").Range("A13:A2000").BorderAround ColorIndex:=1, Weight:=3

'the values range starts in row 14 so the pointer is set to 14
pointer = 14

'fetching the name of the split column
split_col = ThisWorkbook.Worksheets("info").Cells(11, 1)

'looping through all sheets of this workbook except for the info sheet
For Each ws In ThisWorkbook.Sheets
If ws.name = "info" Then

Else
'finding the number of the split column based on it's title
split_col_num = Application.Match(split_col, ws.Range("A1:DA1"), False)
'counting the lines in the sheet
lines_count = Application.CountA(ws.Range("A2:A500000"))

'looping through all cells of the split column and if that value is not yet in the values range on the info sheet then we add it
For i = 2 To lines_count
If Application.CountIf(Sheets("info").Range(Cells(14, 1), Cells(pointer, 1)), ws.Cells(i, split_col_num)) = 0 Then
Sheets("info").Cells(pointer, 1).Value = ws.Cells(i, split_col_num)
pointer = pointer + 1
End If

Next i

End If

Next ws

'turning alerts on again
Application.DisplayAlerts = True

End Sub

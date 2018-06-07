Sub Split()

Dim i, j As Long 'general variables used by for loops
Dim split_val_count As Long 'number of values in the split values range (ie number of files to be created)
Dim file_name As String 'static part of the file name
Dim split_col As String 'name of the column the splitting is based on
Dim split_col_num As Long 'the number of the column the splitting is based on
Dim cur_split_val As String 'variable to store the current split value
Dim links As Variant 'used for breaking links after splitting
Dim wb As Workbook 'used to refer to the workbook that contains the filtered data

'to avoid error messages stopping the script
Application.DisplayAlerts = False

'collecting the following data from the info sheet: number of split values, static part of the file name, name (title) of the split column
split_val_count = Application.CountA(ThisWorkbook.Worksheets("info").Range("A14:C1000"))
file_name = ThisWorkbook.Worksheets("info").Cells(11, 2)
split_col = ThisWorkbook.Worksheets("info").Cells(11, 1)

'removing any existing filters
'ActiveSheet.AutoFilterMode = False

'looping though all split values
For i = 1 To split_val_count
cur_split_val = Worksheets("info").Cells(i + 13, 1)

'creating a new workbook - the filtered data will be copied into this workbook
Set wb = Workbooks.Add

'copy all sheets into the newly created workbook
For Each ws In ThisWorkbook.Sheets
    ws.Copy after:=wb.Sheets(wb.Sheets.Count)
Next ws

'to avoid error messages stopping the script
Application.DisplayAlerts = False

'Sheet1 is a default blank sheet in the newly created workbook. This sheet and the info sheet don't need to be saved so they're deleted
wb.Sheets("Sheet1").Delete
wb.Sheets("info").Delete


'looping through all sheets in the newly created workbook (ie the copies of the sheets in this workbook, except for the info sheet)
For Each ws In wb.Worksheets

'finding the number of the split column based on it's title
split_col_num = Application.Match(split_col, ws.Range("A1:DA1"), False)

'filtering for everything except for the current split value and deleting visible lines (ie leaving only lines where the split column contains the current split value
ws.AutoFilterMode = False
ws.Range("A1").AutoFilter field:=split_col_num, Criteria1:="<>" & cur_split_val
ws.Range("A2:A500000").SpecialCells(xlCellTypeVisible).EntireRow.Delete
On Error Resume Next
'remove filter
ws.ShowAllData

Next ws

'if the "Break Links" is set to yet on the info sheet then break all links before saving the file
If ThisWorkbook.Sheets("info").Cells(11, 3) = "Yes" Then
links = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
For j = 1 To UBound(links)
wb.BreakLink _
    name:=links(j), _
    Type:=xlLinkTypeExcelLinks
Next j
End If

'saving the created workbook, file name contains the static part given by the user and the current split value
With wb
    .SaveAs ThisWorkbook.Path & "\" & file_name & " " & cur_split_val & ".xlsx", ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    .Close 0

End With

'turning alerts on again
Application.DisplayAlerts = True

Next i

End Sub


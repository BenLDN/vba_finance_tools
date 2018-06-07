Attribute VB_Name = "MergeWorkbooks"
'Merging tabls from multiple files in a directory
Sub MergeWorkbooks()

Dim Filename As String
Dim Sheet As Worksheet

'Turning off screen updating to prevent random flashing and to speed up the process
Application.ScreenUpdating = False

'Loop through all workbooks in the directory where THIS workbook is located
'Skip this workbook
'Open each WB and the loop through the worksheets in that WB
'Copy that sheet into THIS workbook

WBname = Dir(ActiveWorkbook.Path & "\" & "*.xls*")

Do While WBname <> ""
    If WBname = ThisWorkbook.Name Then
    Else
        Workbooks.Open Filename:=ActiveWorkbook.Path & "\" & WBname, ReadOnly:=True
        For Each Sheet In ActiveWorkbook.Sheets
        Sheet.Copy After:=ThisWorkbook.Sheets(1)
        Next Sheet
        Workbooks(WBname).Close
    End If
    WBname = Dir()
Loop

Application.ScreenUpdating = True

End Sub

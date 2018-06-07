Attribute VB_Name = "Refresh_Sensitivity_Analysis"
Sub Refresh_Sensitivity_Analysis()
Dim i, j As Integer
Dim g_original, WACC_original As Double

'Saving the current values of g and WACC so they can be restored after we have run the analyis
g_original = Worksheets("FCF").Cells(32, 18).Value
WACC_original = Worksheets("FCF").Cells(34, 18).Value

'Loop throgh all g and WACC combinations
'Update g and WACC in cells R32 and R34 based
'Get the EV from cell E41
'Copy EV to the cell determined by the current g and WACC values

For i = 0 To 6
For j = 0 To 10
Worksheets("FCF").Cells(32, 18).Value = Worksheets("FCF").Cells(52, 3 + i).Value
Worksheets("FCF").Cells(34, 18).Value = Worksheets("FCF").Cells(53 + j, 2).Value
Worksheets("FCF").Cells(53 + j, 3 + i).Value = Worksheets("FCF").Cells(41, 5)
Next j
Next i

'Restoring the original values of g and WACC
Worksheets("FCF").Cells(32, 18).Value = g_original
Worksheets("FCF").Cells(34, 18).Value = WACC_original

End Sub

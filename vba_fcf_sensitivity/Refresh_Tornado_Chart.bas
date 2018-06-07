Attribute VB_Name = "Refresh_Tornado_Chart"
Sub Refresh_Tornado_Chart()
Dim i As Integer
Dim g, g2021, wacc, gm, opex, tax, capex As Double
Dim cellref As String

'Saving the current values of the assumptions
g = Worksheets("FCF").Cells(32, 18).Value
g2021 = Worksheets("FCF").Cells(15, 18).Value
wacc = Worksheets("FCF").Cells(34, 18).Value
gm = Worksheets("FCF").Cells(18, 18).Value
opex = Worksheets("FCF").Cells(20, 18).Value
tax = Worksheets("FCF").Cells(25, 18).Value
capex = Worksheets("FCF").Cells(30, 18).Value

'Loop throgh all assumptions we want to change
'Copy "low" values of this assumption to the appropriate cell
'Get the EV from cell E41 and copy it to the "low" column
'Do the same for the "mid" and "high" values of the assumption
'Assumptions are set to their original values at the end of each loop

For i = 1 To 7

cellref = Worksheets("Tornado").Cells(i + 1, 14).Value

Worksheets("FCF").Range(cellref).Value = Worksheets("Tornado").Cells(i + 1, 2).Value
Worksheets("Tornado").Cells(i + 1, 6).Value = Worksheets("FCF").Cells(41, 5).Value

Worksheets("FCF").Range(cellref).Value = Worksheets("Tornado").Cells(i + 1, 3).Value
Worksheets("Tornado").Cells(i + 1, 7).Value = Worksheets("FCF").Cells(41, 5).Value

Worksheets("FCF").Range(cellref).Value = Worksheets("Tornado").Cells(i + 1, 4).Value
Worksheets("Tornado").Cells(i + 1, 8).Value = Worksheets("FCF").Cells(41, 5).Value

'Restoring the original values of the assumptions
'Since the slices of the tornado chart use an "all things being equal principle...
'...we restore original assumption values after each loop

'the inefficiency caused by this is negligible as the data set is small

Worksheets("FCF").Cells(32, 18).Value = g
Worksheets("FCF").Cells(15, 18).Value = g2021
Worksheets("FCF").Cells(34, 18).Value = wacc
Worksheets("FCF").Cells(18, 18).Value = gm
Worksheets("FCF").Cells(20, 18).Value = opex
Worksheets("FCF").Cells(25, 18).Value = tax
Worksheets("FCF").Cells(30, 18).Value = capex

Next i

'Sort data on "Tornado" sheet by column O (sum of the absolute values of the low and high diff)

Worksheets("Tornado").Columns("A:O").Sort key1:=Range("O2"), order1:=xlAscending, Header:=xlYes

End Sub

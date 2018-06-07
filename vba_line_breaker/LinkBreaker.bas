Attribute VB_Name = "LinkBreaker"
Sub LinkBreaker()

Dim j As Integer
Dim links As Variant
Dim myfolder, myfile As String

Application.DisplayAlerts = False

myfolder = ThisWorkbook.Path

myfile = Dir(myfolder & "\*.xlsx")

Do While myfile <> ""
Set wb = Workbooks.Open(Filename:=myfolder & "\" & myfile, UpdateLinks:=0)
myfile = Dir

links = wb.LinkSources(Type:=xlLinkTypeExcelLinks)

On Error Resume Next

For j = 1 To UBound(links)
wb.BreakLink _
    Name:=links(j), _
    Type:=xlLinkTypeExcelLinks

Next j

wb.Save
wb.Close

Loop

Application.DisplayAlerts = True

End Sub

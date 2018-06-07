'Add all remaining fields to values as Sum

Sub PivotAddValues()
    Dim piv As PivotTable
    Dim field As Long
    For Each piv In ActiveSheet.PivotTables
        For field = 1 To piv.PivotFields.Count
            With piv.PivotFields(field)
                If .Orientation = 0 Then
                    .Orientation = xlDataField
                    .Function = xlSum
                End If
            End With
        Next
    Next
End Sub

Attribute VB_Name = "mod_exc_DataTables"
Public Sub FillTableDownByCopyingBlanksFromAbove()
    Dim rng, rRow, rCel As Range
    Set rng = ActiveSheet.UsedRange
    For Each rRow In rng.Rows
        For Each rCel In rRow.Cells
            With rCel
                If (.Value = "") And (.row > 1) Then
                    .Value = rng.Cells(.row - 1, .Column).Value
                End If
            End With
        Next
    Next
End Sub

Public Sub FillColumnDownByCopyingBlanksFromAbove()
    Dim rCel As Range
    For Each rCel In ActiveSheet.Columns(ActiveCell.Column).Cells
        With rCel
            If (.Value = "") And (.row > 1) Then
                .Value = ActiveSheet.Cells(.row - 1, .Column).Value
            End If
        End With
    Next
End Sub


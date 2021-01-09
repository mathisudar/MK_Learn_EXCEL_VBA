Sub DeleteRelativeRef()
    ActiveCell.Offset(4, 0).Rows("1:1").EntireRow.Select
    Selection.Delete Shift:=xlUp
End Sub

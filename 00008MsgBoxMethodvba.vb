Option Explicit

Sub MyMsgBoxes()

'VBA is an object;MsgBox is a Method
VBA.MsgBox "Hello Mathi"
VBA.MsgBox 2
VBA.MsgBox 2 + 3

'Concatenate Strings
VBA.MsgBox "SURAJ " & "MATHI"

'Get the cell value in the MsgBox
VBA.MsgBox Worksheets(1).Range("C10").Value


End Sub
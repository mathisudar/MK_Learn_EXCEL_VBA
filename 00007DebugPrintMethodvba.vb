Option Explicit

Sub MyFirstProc()

Worksheets(1).Range("A1").Value = "MATHI"
' To get the output in Immediate Window
Debug.Print Worksheets(1).Range("A1").Value

Worksheets(1).Range("A2").Value = "SURAJ"
Debug.Print Worksheets(1).Range("A2").Value


End Sub


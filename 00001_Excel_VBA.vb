Sub Macro1()
'
' Macro1 Macro
' ActiveCell.FormulaR1C1 = "Mathi"
' Range("A1").Select
' Selection.Font.Bold = True

' Color = 65535


    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Mathi"
    Range("A1").Select
    Selection.Font.Bold = True
    Range("A1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

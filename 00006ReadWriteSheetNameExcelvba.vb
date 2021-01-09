'Read the Workbook name using index
? Application.Workbooks(2).Name
MK_Workbook_CH2.xlsm

'Read the WorkSheet name using index
? Application.Workbooks(2).Worksheets(1).Name
MK_Sh1

? Application.Workbooks(2).Worksheets(2).Name
MK_Sh2

'Set the WorkSheet name using index
Application.Workbooks(2).Worksheets(3).Name = "MK_Sh3"


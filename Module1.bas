Attribute VB_Name = "Module1"
Sub deletecells()

'Delete Extra
 Columns("V:AF").EntireColumn.Delete
 Columns("T").EntireColumn.Delete
 
'Changing the 3rd-25the row Height
Cells.Select
    Range("A1").Activate
    Selection.RowHeight = 14.5

'Changing the 3rd-25the row Colunms
Columns("A:H").ColumnWidth = 6
Columns("I").ColumnWidth = 35
Columns("J:P").AutoFit
Columns("Q:T").ColumnWidth = 5
Columns("B").AutoFit

 
'BOLD first row
Range("A1:Z1").Font.Bold = True
Cells.Select
    Range("A1:Z1").Activate
    With Selection
        .HorizontalAlignment = xlLeft
    End With
Range("B2").Font.Bold = True
Range("B2").Interior.Color = RGB(240, 252, 3)
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Range_A1_D6()

Range("A1:D6").Select

End Sub

Sub Active_range()

Range(ActiveCell, "D8").Select

End Sub

Sub Offset_Full()

Range("A1").Offset(RowOffset:=1, ColumnOffset:=1).Select

End Sub

Sub Offset_Short()

Range("A4").Offset(1, 3).Select

End Sub

Sub offset_back()

Range("B4").Offset(-1, -1).Select

End Sub

Sub resize_row_column()

Range("A1").Resize(2, 2).Select

End Sub

Attribute VB_Name = "Module2"

Sub DelBlankRows()

   Columns("A:A").Select
   Selection.SpecialCells(xlCellTypeBlanks).Select
   Selection.EntireRow.Delete

    Range("A1").Select

   End Sub

Sub TotalCompiled()

Sheets(2).Activate

Dim lr As Long
lr = Range("E" & Rows.Count).End(xlUp).Row
Range("E" & lr + 1).Formula = "=SUM(E2:E" & lr & ")"


End Sub

Sub LocaleCopyV2()
   With Sheets(2)
      .Range(.Range("F" & Rows.Count).End(xlUp).Offset(1), .Range("A" & Rows.Count).End(xlUp).Offset(, 5)).Value = Sheets(1).Range("G1").Value
   End With
End Sub

Sub NameCopyV2()
   With Sheets(2)
      .Range(.Range("G" & Rows.Count).End(xlUp).Offset(1), .Range("A" & Rows.Count).End(xlUp).Offset(, 6)).Value = Sheets(1).Range("F1").Value
   End With
End Sub

Sub DateAdd()
   With Sheets(2)
      .Range(.Range("H" & Rows.Count).End(xlUp).Offset(1), .Range("A" & Rows.Count).End(xlUp).Offset(, 7)).Value = Sheets(1).Range("H1").Value
   End With
End Sub






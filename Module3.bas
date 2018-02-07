Attribute VB_Name = "Module3"


Sub Copy()
Dim Last_Row As Long
Last_Row = Range("A" & Rows.Count).End(xlUp).Row
Sheets(2).Range("A2:H" & Last_Row).Copy Sheets(2).Range(Last_Row & "A:H")

End Sub









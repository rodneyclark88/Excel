Attribute VB_Name = "Module2"

Sub DateAdd()

Dim rInt As Range
    Dim rCell As Range
    Dim tCell As Range

    Set rInt = Sheets(2).Range("A2:A150")
    If Not rInt Is Nothing Then
        For Each rCell In rInt
            Set tCell = rCell.Offset(0, 7)
            If IsEmpty(tCell) Then
               tCell = Now
                tCell.NumberFormat = "mmm d, yyyy"
            End If
       Next
    End If

End Sub


Sub DelBlankRows()

   Columns("A:A").Select
   Selection.SpecialCells(xlCellTypeBlanks).Select
   Selection.EntireRow.Delete

    Range("A1").Select

   End Sub

'Copy Name & Location of Order
'Sub OrderInfo()
'Method 1
'Sheets(1).Range("F1").Copy Destination:=Sheets(2).Range("E1")

'Application.CutCopyMode = False
'End Sub

Sub GetName()

Dim rInt As Range
    Dim rCell As Range
    Dim tCell As Range

    Set rInt = Sheets(2).Range("A2:A150")
    'If Not rInt Is Nothing Then
        For Each rCell In rInt
            Set tCell = rCell.Offset(0, 6)
            tCell = Sheets(1).Range("F1")
       Next
    

End Sub
Sub GetLocale()

Dim rInt As Range
    Dim rCell As Range
    Dim tCell As Range

    Set rInt = Sheets(2).Range("A2:A150")
    'If Not rInt Is Nothing Then
        For Each rCell In rInt
            Set tCell = rCell.Offset(0, 5)
            tCell = Sheets(1).Range("G1")
       Next

End Sub

Sub TotalCompiled()

Sheets(2).Activate

Dim LR As Long
LR = Range("E" & Rows.Count).End(xlUp).Row
Range("E" & LR + 1).Formula = "=SUM(E2:E" & LR & ")"


End Sub



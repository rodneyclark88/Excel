Attribute VB_Name = "Module1"

Sub RemoveBlankRows()
Sheets(2).Select
'Deletes any row with blank cells located inside a designated range


Dim Rng As Range

'Store blank cells inside a variable
  On Error GoTo NoBlanksFound
    Set Rng = Range("A2:E105").SpecialCells(xlCellTypeBlanks)
  On Error GoTo 0

'Delete entire row of blank cells found
  Rng.EntireRow.Delete
  

Exit Sub

NoBlanksFound:
  'MsgBox "No Blank cells were found"

End Sub

Sub FormatRange()

'Remove borders, Center Totals/Price/Qty, Paste Values

Dim Rng As Range

Set Rng = ActiveSheet.Range("A2:E1025")

Rng.Borders.LineStyle = xlNone

Range("C2:E150").Select
    With Selection
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

Worksheets(2).Range("A2:E1025").Copy

Worksheets(2).Range("A2").PasteSpecial Paste:=xlPasteValues

Worksheets(2).Columns("C:E").AutoFit

Worksheets(2).Range("A1").Select

Application.CutCopyMode = False

End Sub

Sub SumColE()
Dim LRow As Long
LRow = Range("E65536").End(xlUp).Row + 1
Cells(LRow, 1).Formula = "=SUM(E1:E" & LRow - 1 & ")"

End Sub


'This will help to overcome errors when deleting selections of ranges:

On Error Resume Next

'If you change the value in column B:

If Target.Column = "C" Then

    'If the end result of the change is not empty:
    
    If Target.Value > 0 Then
    
    Target.Offset(0, 5).Value = "=TODAY()" 'Gets the current date in cell to the left in column A
    
    Target.Offset(0, 5).Value = Target.Offset(0, 5).Value 'Replaces the formula with its value
    
    Else 'If the end result of the change is empty (deleted):
    
    Target.Offset(0, 5).Value = "" 'Removes the date to the left in column A
    
    End If

End If

End Sub


Sub BestCopy1()

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
'Application.EnableEvents = False

Sheets(1).Activate

Dim c As Range
Dim destROW As Long
For Each c In Sheets(1).Range("C2:C6")
    destROW = Sheets(2).Cells(Rows.Count, "A").End(xlUp).Row + 1
    If c.Value > 0 Then c.Offset(, -2).Resize(, 5).Copy Sheets(2).Range("A" & destROW)
Next c

Call BestCopy2
Call BestCopy3
Call BestCopy4
Call BestCopy5
Call RemoveBlankRows
Call FormatRange
Call DateAdd
Call LocaleCopyV2
Call NameCopyV2
Call DelBlankRows
Call TotalCompiled


Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
'Application.EnableEvents = True

End Sub

Sub BestCopy2()

Sheets(1).Activate

Dim c As Range
Dim destROW As Long
For Each c In Sheets(1).Range("C8:C18")
    destROW = Sheets(2).Cells(Rows.Count, "A").End(xlUp).Row + 1
    If c.Value > 0 Then c.Offset(, -2).Resize(, 5).Copy Sheets(2).Range("A" & destROW)
Next c

End Sub

Sub BestCopy3()

Sheets(1).Activate

Dim c As Range
Dim destROW As Long
For Each c In Sheets(1).Range("C20:C45")
    destROW = Sheets(2).Cells(Rows.Count, "A").End(xlUp).Row + 1
    If c.Value > 0 Then c.Offset(, -2).Resize(, 5).Copy Sheets(2).Range("A" & destROW)
Next c

End Sub

Sub BestCopy4()

Sheets(1).Activate

Dim c As Range
Dim destROW As Long
For Each c In Sheets(1).Range("C47:C93")
    destROW = Sheets(2).Cells(Rows.Count, "A").End(xlUp).Row + 1
    If c.Value > 0 Then c.Offset(, -2).Resize(, 5).Copy Sheets(2).Range("A" & destROW)
Next c

End Sub

Sub BestCopy5()

Sheets(1).Activate

Dim c As Range
Dim destROW As Long
For Each c In Sheets(1).Range("C95:C104")
    destROW = Sheets(2).Cells(Rows.Count, "A").End(xlUp).Row + 1
    If c.Value > 0 Then c.Offset(, -2).Resize(, 5).Copy Sheets(2).Range("A" & destROW)
Next c

End Sub


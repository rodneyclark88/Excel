Attribute VB_Name = "Module1"

Sub Location()



End Sub


Sub RemoveBlankRows()
Sheets(2).Select
'PURPOSE: Deletes any row with blank cells located inside a designated range
'SOURCE: www.TheSpreadsheetGuru.com

Dim rng As Range

'Store blank cells inside a variable
  On Error GoTo NoBlanksFound
    Set rng = Range("A2:E105").SpecialCells(xlCellTypeBlanks)
  On Error GoTo 0

'Delete entire row of blank cells found
  rng.EntireRow.Delete
  

Exit Sub

NoBlanksFound:
  MsgBox "No Blank cells were found"

End Sub

Sub FormatRange()

'Remove borders, Center Totals/Price/Qty, Paste Values

Dim rng As Range

Set rng = ActiveSheet.Range("A2:E150")

rng.Borders.LineStyle = xlNone

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

Worksheets(2).Range("A2:E150").Copy

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

    Private Sub Worksheet_Change(ByVal Target As Range)

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
Application.EnableEvents = False

Sheets(2).Activate
Range("A2:E150").ClearContents

Sheets(1).Activate

Dim c As Range
For Each c In Sheets(1).Range("C2:C6")
    If c.Value > 0 Then c.Offset(, -2).Resize(, 5).Copy Sheets(2).Range(c.Offset(, -2).Address)
Next c

Call BestCopy2
Call BestCopy3
Call BestCopy4
Call BestCopy5
Call RemoveBlankRows
Call FormatRange

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True

End Sub

Sub BestCopy2()

Sheets(1).Activate

Dim c As Range
For Each c In Sheets(1).Range("C9:C18")
    If c.Value > 0 Then c.Offset(, -2).Resize(, 5).Copy Sheets(2).Range(c.Offset(, -2).Address)
Next c

End Sub

Sub BestCopy3()

Sheets(1).Activate

Dim c As Range
For Each c In Sheets(1).Range("C20:C45")
    If c.Value > 0 Then c.Offset(, -2).Resize(, 5).Copy Sheets(2).Range(c.Offset(, -2).Address)
Next c

End Sub

Sub BestCopy4()

Sheets(1).Activate

Dim c As Range
For Each c In Sheets(1).Range("C47:C93")
    If c.Value > 0 Then c.Offset(, -2).Resize(, 5).Copy Sheets(2).Range(c.Offset(, -2).Address)
Next c

End Sub

Sub BestCopy5()

Sheets(1).Activate

Dim c As Range
For Each c In Sheets(1).Range("C95:C104")
    If c.Value > 0 Then c.Offset(, -2).Resize(, 5).Copy Sheets(2).Range(c.Offset(, -2).Address)
Next c

End Sub

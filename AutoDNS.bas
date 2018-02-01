Attribute VB_Name = "Module11"
Sub FormatToNumber()

Application.ScreenUpdating = False


Range("B2").Select
Range(ActiveCell, ActiveCell.End(xlDown)).Select

With Selection
    .NumberFormat = "0"
    .Value = .Value
End With

Call SelectToExport
Call ExportToTemplate
Call CopyToTabs
Call SaveTemplate
Call Filter50

Application.ScreenUpdating = True

End Sub


Sub SelectToExport()

With Worksheets("Source")
    lastrow = .Cells(Rows.Count, "A").End(xlUp).Row
    .Range("A1:J" & lastrow).Select
End With

End Sub


Sub ExportToTemplate()
   '
   ' CopyOpenItems Macro
   ' Copy open items to sheet.
   '
   ' Keyboard Shortcut: Ctrl+Shift+O
   '
   Dim wbTarget            As Workbook 'workbook where the data is to be pasted
   Dim wbThis              As Workbook 'workbook from where the data is to be copied
   Dim strName             As String   'name of the source sheet/ target workbook
    
   'set to the current active workbook (the source book)
   Set wbThis = ActiveWorkbook
    
   'get the active sheetname of the book
   strName = ActiveSheet.Name
    
   'open a workbook that has same name as the sheet name
   Set wbTarget = Workbooks.Open("C:\Users\rodneyc\Desktop\Source.xlsm")
   
   'Activate first tab
   Worksheets("Export Database Tab").Activate
    
   'select cell A1 on the target book
   Range("B2").Select
    
   'clear existing values form target book
   Range("B2:K250").ClearContents

   'activate the source book
   wbThis.Activate
    
   'clear any thing on clipboard to maximize available memory
   Application.CutCopyMode = False
    
   'copy the range from source book
   With Worksheets("Source")
    lastrow = .Cells(Rows.Count, "A").End(xlUp).Row
    .Range("A2:J" & lastrow).Select
    Selection.Copy
   End With
   
   'wbThis.Range("A12:M62").Copy
   'paste the data on the target book
   
   wbTarget.Activate

   ActiveSheet.Range("B2").Select
   Selection.PasteSpecial
    
   'clear any thing on clipboard to maximize available memory
   Application.CutCopyMode = False
    
   'save the target book
   wbTarget.Save
    
   'close the workbook
   wbTarget.Close

   'activate the source book again
   wbThis.Activate
    
   'clear memory
   Set wbTarget = Nothing
   Set wbThis = Nothing
    
End Sub


Sub CopyToTabs()

Workbooks.Open ("C:\Users\rodneyc\Desktop\Source.xlsm")

With Worksheets("Export Database Tab")
    lastrow = .Cells(Rows.Count, "B").End(xlUp).Row
    .Range("B2:K" & lastrow).Select
    Selection.Copy
   End With

'Activate the destination worksheet
Sheets(2).Activate
'Select the target range
Range("B4").Select
'Paste in the target destination
ActiveSheet.Paste

'Repeat to all Tabs

Sheets(3).Activate
Range("B4").Select
ActiveSheet.Paste

Sheets(4).Activate
Range("B4").Select
ActiveSheet.Paste

Sheets(5).Activate
Range("B4").Select
ActiveSheet.Paste

Sheets(6).Activate
Range("B4").Select
ActiveSheet.Paste

Application.CutCopyMode = False

End Sub

Sub SaveTemplate()

Sheets(1).Activate
ActiveWorkbook.SaveAs Filename:="C:\Users\rodneyc\Desktop\ExportReady.xlsm"

End Sub

Sub Filter50()

For Each cell In Sheets(2).Range("E4:E200")
If cell.Value < 50 Then
    matchrow = cell.Row
    Rows(matchrow).EntireRow.Delete
    
End If

Next
         
End Sub

Sub RowDelete()

lastrow = ActiveSheet.Worksheets("Over $50-Over 90 days").Cells(Rows.Count, 1).End(xlUp).Row

For i = lastrow To 2 Step by - 1
If ThisWorkbook.Worksheets("Over $50-Over 90 days").Cells(i, 6).Value < 50.01 Then

Rows(i).Delete

End If

Next

ThisWorkbook.Worksheets("Over $50-Over 90 days").Cells(1, 1).Select



End Sub




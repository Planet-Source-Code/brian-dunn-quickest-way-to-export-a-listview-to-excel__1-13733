<div align="center">

## Quickest way to export a listview to Excel


</div>

### Description

This is a faster way to take a listview control and display its contents in a new Excel workbook.

A common mistake in using OLE to manipulate Excel is to send data values one cell at a time. However, if you are exporting listview, it is much faster to create a two-dimensional array of the data and then send the entire array to Excel all at once. This method can be applied to grids, recordsets, or any other table-like data.

This code will also allow the user to select multiple, non-contiguous rows for export. Hidden columns are not exported, either. Also, if the ColumnHeader.Tag properties have been set to "string", "number", or "date", the Excel columns will be formatted as such.
 
### More Info
 
A reference to a ListView control.

The listview allows multiple row selection.

True if it worked, False if not


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Dunn](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-dunn.md)
**Level**          |Intermediate
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-dunn-quickest-way-to-export-a-listview-to-excel__1-13733/archive/master.zip)





### Source Code

```
Public Function ExportToExcel(lvw As MSComctlLib.ListView) As Boolean
 Dim objExcel As Excel.Application
 Dim objWorkbook As Excel.Workbook
 Dim objWorksheet As Excel.Worksheet
 Dim objRange As Excel.Range
 Dim lngResults As Long
 Dim i As Integer
 Dim intCounter As Integer
 Dim intStartRow As Integer
 Dim strArray() As String
 Dim intVisibleColumns() As Integer
 Dim intColumns As Integer
 Dim itm As ListItem
 'If there are no selected items in the listview control
 If lvw.SelectedItem Is Nothing Then
 MsgBox "There aren't any items in the listview selected." _
  , vbOKOnly + vbInformation, "Export Failed"
 GoTo ExitFunction
 End If
 'Ask the user if they want to export just the selected items
 lngResults = MsgBox("Do you want to export only the selected rows to Excel? " _
 , vbYesNoCancel + vbQuestion, "Select Rows For Export")
 If lngResults = vbCancel Then
 GoTo ExitFunction
 End If
 Screen.MousePointer = vbHourglass
 'Try to create an instance of Excel
 On Error Resume Next
 Set objExcel = New Excel.Application
 If Err.Number > 0 Then
 MsgBox "Microsoft Excel is not loaded on this machine.", vbOKOnly + vbCritical, "Error Loading Excel"
 GoTo ExitFunction
 End If
 On Error GoTo HANDLE_ERROR
 ' Don't allow user to affect workbook
 objExcel.Interactive = False
 If objExcel.Visible = False Then
 objExcel.Visible = True
 End If
 objExcel.WindowState = xlMaximized
 Set objWorkbook = objExcel.Workbooks.Add
 Set objWorksheet = objWorkbook.Sheets(1)
 intCounter = 0
 Set objRange = objWorksheet.Rows(1)
 objRange.Font.Size = 10
 objRange.Font.Bold = True
 For i = 1 To lvw.ColumnHeaders.Count
 If lvw.ColumnHeaders(i).Width <> 0 Then
  ' Create an array of visible column indexes
  intColumns = intColumns + 1
  ReDim Preserve intVisibleColumns(1 To intColumns)
  intVisibleColumns(intColumns) = i
  objRange.Cells(1, intColumns) = lvw.ColumnHeaders(i).Text
  With objWorksheet.Columns(intColumns)
  Select Case LCase$(lvw.ColumnHeaders(i).Tag)
  ' If tag is empty, format as text
  Case "string", ""
   .NumberFormat = "@"
  Case "number"
   .NumberFormat = "#,##0.00_);(#,##0.00)"
   .HorizontalAlignment = xlRight
  Case "date"
   .NumberFormat = "mm/dd/yyyy"
   .HorizontalAlignment = xlRight
  End Select
  End With
 End If
 Next i
 ' Dimension array to number of listitems
 ReDim strArray(1 To lvw.ListItems.Count, 1 To intColumns)
 intCounter = 0
 intStartRow = 2
 For Each itm In lvw.ListItems
 ' A response of vbNo meant to export all the items
 If lngResults = vbNo Or itm.Selected Then
  ' increment the number of selected rows
  intCounter = intCounter + 1
  For i = 1 To intColumns
  If intVisibleColumns(i) = 1 Then
   strArray(intCounter, 1) = itm.Text
  Else
   strArray(intCounter, i) = itm.SubItems(intVisibleColumns(i) - 1)
  End If
  Next i
 End If
 Next itm
 ' Send entire array to Excel range
 With objWorksheet
 .Range(.Cells(2, 1), _
  .Cells(2 + intCounter - 1, intColumns)) = strArray
 End With
 objWorksheet.Columns.AutoFit
 objExcel.Interactive = True
 ExportToExcel = True
ExitFunction:
 Screen.MousePointer = vbDefault
 Exit Function
HANDLE_ERROR:
 MsgBox "Export to Excel failed. Encountered thej following Error" & vbCrLf & vbCrLf & _
   Err.Number & ": " & Err.DESCRIPTION, vbOKOnly + vbCritical, "Error Exporting To Excel"
 Set objRange = Nothing
 Set objWorksheet = Nothing
 Set objWorkbook = Nothing
 objExcel.Quit
 GoTo ExitFunction
End Function
```


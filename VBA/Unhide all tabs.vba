Option Explicit

Sub UnhideAllSheets()

'Unhide all sheets in workbook.
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
    ws.Visible = xlSheetVisible
Next ws

End Sub

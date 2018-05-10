Option Explicit

Sub resource()

'enabling and disabling events
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.DisplayAlerts = False

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.DisplayAlerts = True

'defining dimensions or variables
Dim wrksht As Worksheet
Dim p As Integer
Dim z As String
Dim a As Long
Dim b As Double

'select a worksheet
Worksheets("Sheet3").Activate
Worksheets("Sheet3").Select
Sheets("Sheet4").Activate
Sheet1.Activate
Sheet1.Select

'Select a range (cell) in the active worksheet
ActiveSheet.Range("C5").Select
Worksheets("Sheet3").Range("C5").Select
Range("E10").Select

'find last row
Dim lr As Integer
lr = Worksheets("sheet1").Cells(Rows.Count, 1).End(xlUp).Row

'delete a table + Unlist a table and remove formats
Worksheets(1).Select
ActiveSheet.ListObjects("Table3").Delete

ActiveSheet.ListObjects("Table4").Unlist
ActiveSheet.UsedRange.ClearFormats

'defining an array
Dim nr As Variant
nr = Array("DATA_1", "DATA_2", "DATA_3", "DATA_4", "DATA_5", "DATA_6")

'save and close active workbook
ActiveWorkbook.Close True

'copying the formula in first row to lastrow of preceding columns and pasting as values from the next row _
 that leaves formula only in the first row (excel computes faster as a result)
 'can write efficient way to copy paste (Pending)
Dim lastrow1 As Long
lastrow1 = Worksheets("Data").Cells(Rows.Count, 1).End(xlUp).Row

Range("V9:AF9").Copy
Range("V10:AF" & lastrow1).Select
ActiveSheet.Paste
Range("V10:AF" & lastrow1).Copy
Range("V10").Select
Selection.PasteSpecial xlPasteValues

'Open workbook and switch windows (Ctrl+Tab)
Windows("XYZ Damage Reduction Source Data Query.xlsx").Activate
Workbooks.Open "\\USXS1031\Groups\EMA-XYZ\Public\XYZ Damage Data\XYZ Damage Reduction Source Data Query.xlsx"

'closing a workbook
Workbooks("XYZ Damage Reduction Source Data Query.xlsx").Close

'refresh pivot table
ActiveSheet.PivotTables("PivotTable1").RefreshTable

'filtering a value and deleting rows and sort ascending
ActiveSheet.UsedRange.Select
Selection.AutoFilter
ActiveSheet.UsedRange.AutoFilter Field:=11, Criteria1:="2016"
ActiveSheet.Range("$A$1:$AS$" & lastrow3).Offset(1, 0).SpecialCells _
    (xlCellTypeVisible).EntireRow.Delete
    
ActiveSheet.AutoFilterMode = False
               'sort
Range("A2:AS" & lastrow3).Sort key1:=Range("K2:K" & lastrow3), _
   order1:=xlAscending, Header:=xlNo

End Sub

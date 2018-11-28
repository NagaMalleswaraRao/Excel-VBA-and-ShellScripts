Option Explicit

Sub resource()

'enabling and disabling events
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.DisplayAlerts = False
Application.AskToUpdateLinks = False

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.DisplayAlerts = True
Application.AskToUpdateLinks = True

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

'closing a workbook
Workbooks("XYZ Damage Reduction Source Data Query.xlsx").Close
     
'save and close active workbook
ActiveWorkbook.Close True

'copying the formula in first row to lastrow of preceding columns and pasting as values from the next row _
'that leaves formula only in the first row (excel computes faster as a result)
Dim LastRow_WkSt1 As Long
LastRow_WkSt1 = Sheets("WkSt1").Cells(Rows.Count, 26).End(xlUp).Row

Worksheets("WkSt1").Range("AA7:IG" & LastRow_WkSt1).FillDown
With Worksheets("WkSt1").Range("AA8:IG" & LastRow_WkSt1)
    .Value = .Value
End With

'Move and append data from Wkst2 to Wkst4 
LastRow_WkSt2 = Sheets("WkSt2").Cells(Rows.Count, 1).End(xlUp).Row
LastRow_WkSt4 = Sheets("WkSt4").Cells(Rows.Count, 2).End(xlUp).Row
Rowz = LastRow_WkSt4 + LastRow_WkSt2 - 6 '6 is the row no. of header from the top 

Worksheets("WkSt4").Range(Cells(LastRow_WkSt4 + 1, 1), Cells(Rowz, 31)).Value = _
Worksheets("WkSt2").Range("A7:AE" & LastRow_WkSt2).Value                         

'Open workbook and switch windows (Ctrl+Tab)
Windows("XYZ Damage Reduction Source Data Query.xlsx").Activate
Workbooks.Open "\\USXS1031\Groups\EMA-XYZ\Public\XYZ Damage Data\XYZ Damage Reduction Source Data Query.xlsx"

'refresh pivot table in another sheet
Sheets("WkSt3").PivotTables("PivotTable4").PivotCache.Refresh

'''Deleting the old "current month" units, sorting data to remove blanks in middle
Dim rng As Range, rng2 As Range
LastRow_Units2 = Sheets("Units").Cells(Rows.Count, 13).End(xlUp).Row
Set rng = Worksheets("Units").Range("L9:R" & LastRow_Units2)

rng.AutoFilter

With rng
        .AutoFilter Field:=2, Criteria1:=Year
        .AutoFilter Field:=3, Criteria1:=Month
        .Offset(1, 0).SpecialCells(xlCellTypeVisible).ClearContents
End With

ActiveSheet.ShowAllData

''Sort on Year and Month to remove blanks in middle
With ActiveSheet.Sort
     .SortFields.Add Key:=Range("M9"), Order:=xlAscending
     .SortFields.Add Key:=Range("N9"), Order:=xlAscending
     .SetRange rng
     .Header = xlYes
     .Apply
End With

''Reapplying filter, if no filter is present. Later we can bring updated Units here
rng.AutoFilter
If ActiveSheet.AutoFilterMode Then
Else
   rng.AutoFilter
End If

'Filter multiple items in a field
LastRow_Units2 = Sheets("Units").Cells(Rows.Count, 13).End(xlUp).Row
Set rng2 = Worksheets("Units").Range("L9:S" & LastRow_Units2)

With rng2
        .AutoFilter Field:=1, Criteria1:=Array( _
 "Asheville", "Springfield", "Memphis"), Operator:=xlFilterValues
End With

'Refresh Data Connections
Dim objConn As Variant
For Each objConn In ThisWorkbook.Connections
    objConn.Refresh
Next

'Clear contents
LastRow_WkSt4 = Sheets("WkSt4").Cells(Rows.Count, 2).End(xlUp).Row
Worksheets("WkSt4").Range("A2:AE" & LastRow_WkSt4).ClearContents

'Paste the time stamp in a sheet cell after macro has run
Worksheets("Refresh").Range("J16") = Now

End Sub

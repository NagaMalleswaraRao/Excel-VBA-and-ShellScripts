Option Explicit

Sub resource()

'enabling and disabling events
Application.ScreenUpdating = False 'Runs macro code without showing what it's doing
Application.EnableEvents = False 'Suppresses doubleclick, sheet selection events of Worksheet code
Application.DisplayAlerts = False 'Suppresses alerts like "Do you want to save this wkbk?" etc.
Application.AskToUpdateLinks = False 'Suppresses "Do you want to update links (formula references another wkbk)?"

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.DisplayAlerts = True
Application.AskToUpdateLinks = True

'Toggle calculation mode
Application.Calculation = xlCalculationAutomatic
Application.Calculation = xlCalculationManual     
     
'defining dimensions or variables
Dim wrksht As Worksheet
Dim p As Integer
Dim z As String
Dim a As Long
Dim b As Double

'define public variables. These can be used throughout the workbook across modules
Public LastRow_1 As Long
Public wkbk1 As Workbook
Public str1 As String

'select a worksheet (try to avoid this as much as possible)
Worksheets("Sheet3").Activate
Worksheets("Sheet3").Select
Sheets("Sheet4").Activate

'Select a range (cell) in the active worksheet (try to avoid this as much as possible)
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
                    
'''Splitting comma separated cell value to an array and writing the elements back to Excel range                    
Dim aray As Variant, str As String
Dim i As Integer, m As Integer
'read comma separated cell value from Excel
str = Range("B2").Value
'read the individual elements to an array
aray = VBA.split(str, ",")

i = 3
'Write values of the array to Excel range
For m = LBound(aray) To UBound(aray)
    Cells(i, 3).Value = aray(m)
    i = i + 1
Next
                    
'closing a workbook
Workbooks("XYZ Damage Reduction Source Data Query.xlsx").Close
     
'save and close active workbook
ActiveWorkbook.Close True

'Delete data from A to AE, leave a formula row from Af to BC                         
LastRow_WkSt40 = Worksheets("WkSt4").Cells(Rows.Count, 2).End(xlUp).Row

If LastRow_WkSt40 > 2 Then
    Worksheets("WkSt4").Range("A2:AE" & LastRow_WkSt40).ClearContents
    Worksheets("WkSt4").Range("AF3:BC" & LastRow_WkSt40).ClearContents
End If                        

'copying the formula in first row to lastrow of preceding columns and pasting as values from the next row _
'that leaves formula only in the first row (excel computes faster as a result)
Dim LastRow_WkSt1 As Long
LastRow_WkSt1 = Sheets("WkSt1").Cells(Rows.Count, 26).End(xlUp).Row

'This converts cell value 0012 to 12 (avoid this if you have numeric text and use the code below)
Worksheets("WkSt1").Range("AA7:IG" & LastRow_WkSt1).FillDown
With Worksheets("WkSt1").Range("AA8:IG" & LastRow_WkSt1)
    .Value = .Value
End With

'This leaves the formats and data types as is
With Worksheets("WkSt1").Range("AA8:IG" & LastRow_WkSt1)
    .Copy
    .PasteSpecial xlPasteValues
End With
Application.CutCopyMode = False

'Move and append data from Wkst2 to Wkst4 
LastRow_WkSt2 = Sheets("WkSt2").Cells(Rows.Count, 1).End(xlUp).Row
LastRow_WkSt4 = Sheets("WkSt4").Cells(Rows.Count, 2).End(xlUp).Row
Rowz = LastRow_WkSt4 + LastRow_WkSt2 - 6 '6 is the row no. of header from the top 

'This converts 0012 to 12 (avoid this if you have numeric text and use the code below)
Worksheets("WkSt4").Range(Cells(LastRow_WkSt4 + 1, 1), Cells(Rowz, 31)).Value = _
Worksheets("WkSt2").Range("A7:AE" & LastRow_WkSt2).Value                         

'No need to define "Rowz" for the below snippet
Worksheets("WkSt2").Range("A7:AE" & LastRow_WkSt2).Copy
Worksheets("WkSt4").Cells(LastRow_WkSt4 + 1, 1).PasteSpecial xlPasteValues
Application.CutCopyMode = False

'While avoiding 'Select' or 'Activate', when you have to use Cells in a Range, use the below snippet
Sheets("A").Range(Sheets("A").Cells(lr + 1, 14), Sheets("A").Cells(lr + 1, 20)).Value = "Lol"

'Open workbook and switch windows (Ctrl+Tab)
Windows("XYZ Damage Reduction Source Data Query.xlsx").Activate
Workbooks.Open "\\USXS1031\Groups\EMA-XYZ\Public\XYZ Damage Data\XYZ Damage Reduction Source Data Query.xlsx"

'refresh pivot table in another sheet
Worksheets("WkSt3").PivotTables("PivotTable4").PivotCache.Refresh

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
Worksheets("Units").Sort.SortFields.Clear

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

'Filter Multiple items in multiple fields
'rng is Range to filter, arr is an array  
With rng
     .AutoFilter Field:=3, Criteria1:=arr1, Operator:=xlFilterValues
     .AutoFilter Field:=4, Criteria1:=arr2, Operator:=xlFilterValues
     .AutoFilter Field:=5, Criteria1:=arr3, Operator:=xlFilterValues
End With


'Refresh all Data Connections
Dim objConn As Variant
For Each objConn In ThisWorkbook.Connections
    objConn.Refresh
Next

'Paste the time stamp in a wkst cell after the macro has run
Worksheets("Refresh").Range("J16") = Now

'call a macro (UnhideAllSheets) in this workbook
Call UnhideAllSheets

'''call a macro (tgt_macro) in a opened workbook
Dim str as string, wkbk_name as string
Dim wkbk as workbook
'read the filepath from a cell in Thisworkbook and open it and assign it to a variable
'sample filepath: L:\Public\Finance\2019\Input\Template_file.xlsb
wkbk_name = Thisworkbook.Worksheets("Macros").Cells(10, 3).Value
Workbooks.Open wkbk_name
Set wkbk = ActiveWorkbook
str = "'Template_file.xlsb'!tgt_macro"
wkbk.Application.Run (str)

End Sub

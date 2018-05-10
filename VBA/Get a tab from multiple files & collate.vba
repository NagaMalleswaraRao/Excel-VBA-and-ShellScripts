Sub collate()
  
'Put the templates and file -to collate them- in a new folder
Dim wk As Workbook, wk2 As Workbook
Dim wkst As Worksheet
Dim str As String, str2 As String, str3 As String
Dim i As Integer, lr As Integer
'Objective: out of all the templates, get only one tab and
'rename the tab with Location indicator

Set wk = ThisWorkbook
'Map tab contains file path of templates and names to be used for tabs after collating
lr = wk.Worksheets("Map").Cells(Rows.Count, 7).End(xlUp).Row
str2 = wk.Worksheets("Map").Cells(4, 2).Value

For i = 4 To lr
    str = wk.Worksheets("Map").Cells(i, 10).Value
    str3 = wk.Worksheets("Map").Cells(i, 8).Value
    Workbooks.Open (str)
    Set wk2 = ActiveWorkbook
        For Each wkst In wk2.Worksheets
            If wkst.Name = str2 Then
                'copy the tab and move before Map tab
                wkst.Copy wk.Worksheets("Map")
                'Rename tab in the collated file
                wk.ActiveSheet.Name = str3
            End If
        Next
    wk2.Close False
Next i

'After collating, hide the mapping tab
wk.Worksheets("Map").Visible = False

End Sub

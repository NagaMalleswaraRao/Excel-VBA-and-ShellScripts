Sub collate()
  
'Put the templates and this file in a new folder

Dim wk As Workbook, wk2 As Workbook
Dim wkst As Worksheet
Dim str As String, str2 As String, str3 As String
Dim i As Integer, lr As Integer
'Objective: out of all the templates, get only current EBIT and _
 rename the tab with Location indicator

Set wk = ThisWorkbook
lr = wk.Worksheets("Map").Cells(Rows.Count, 7).End(xlUp).Row
str2 = wk.Worksheets("Map").Cells(4, 2).Value

For i = 4 To lr
    str = wk.Worksheets("Map").Cells(i, 10).Value
    str3 = wk.Worksheets("Map").Cells(i, 8).Value
    Workbooks.Open (str)
    Set wk2 = ActiveWorkbook
        For Each wkst In wk2.Worksheets
            If wkst.Name = str2 Then
                wkst.Copy wk.Worksheets("Map")
                wk.ActiveSheet.Name = str3
            End If
        Next
    wk2.Close False
Next i

wk.Worksheets("Map").Visible = False

End Sub

Sub disint()

'Objective: Copies each wkst(template) as a new workbook, rename tab and file name
Dim i As Integer
Dim filename As String
Dim xpath As String
Dim xWs As Worksheet
xpath = Application.ActiveWorkbook.Path

For Each xWs In ThisWorkbook.Worksheets
    If xWs.Name <> "Map" Then
        'xWs.Name is similar to Goa-Jan
        i = Application.WorksheetFunction.Find("-", xWs.Name) - 1
        filename = "0118-EBIT " & Left(xWs.Name, i)
        xWs.Copy
        Application.ActiveWorkbook.SaveAs filename:=xpath & "\" & filename & ".xlsx"
        ActiveSheet.Name = "Jan EBIT"
        ActiveWorkbook.Save
        Application.ActiveWorkbook.Close False
    End If
Next

End Sub

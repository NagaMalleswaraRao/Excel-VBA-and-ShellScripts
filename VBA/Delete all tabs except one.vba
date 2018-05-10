Sub delete_templates()

Dim wk As Worksheet
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.DisplayAlerts = False

'Unhide all sheets in workbook.
For Each wk In ActiveWorkbook.Worksheets
    wk.Visible = xlSheetVisible
Next wk

'Objective: delete all except 'Map' tab
For Each wk In ActiveWorkbook.Worksheets
    If wk.Name <> "Map" Then
        wk.Delete
    End If
Next

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.DisplayAlerts = True

End Sub

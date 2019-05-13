Sub edit_links()

''' Change links
Dim old_link As String, new_link As String
old_link = ActiveWorkbook.Worksheets("Data").Range("I3")
new_link = ActiveWorkbook.Worksheets("Data").Range("I4")

ActiveWorkbook.ChangeLink old_link, new_link, xlLinkTypeExcelLinks

End Sub

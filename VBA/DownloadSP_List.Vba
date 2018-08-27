Sub ImportSharePointList()
    Application.DisplayAlerts = False
    
    Dim objMyList As ListObject
    Dim objWksheet As Worksheet
    Dim strSPServer As String
    Const SERVER As String = "home.flux-na.com/sites/StoreAudit"
    Const LISTNAME As String = "{F6706998-D2D3-4FAA-B69F-9D4F5ADDD95E}"
    Const VIEWNAME As String = "{C3E49D8F-19A5-44ED-AC37-76AC9B501692}"

   ' The SharePoint server URL pointing to
   ' the SharePoint list to import into Excel.
    strSPServer = "http://" & SERVER & "/_vti_bin"
    ' delete table from Temp tab
    Dim WS As Worksheet
    Worksheets("Temp").Delete
    Set WS = Sheets.Add(after:=Sheets(Worksheets.Count))
    WS.Name = "Temp"
    
    Set objWksheet = Worksheets("Temp")
    ' Add a list range to the newly created worksheet
    ' and populated it with the data from the SharePoint list.
    Set objMyList = objWksheet.ListObjects.Add(xlSrcExternal, _
        Array(strSPServer, LISTNAME, VIEWNAME), False, , Range("A1"))

    Application.DisplayAlerts = True
End Sub


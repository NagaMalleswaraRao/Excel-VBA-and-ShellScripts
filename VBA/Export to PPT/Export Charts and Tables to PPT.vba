Option Explicit
Sub Export_Charts_And_Tables_To_PPT()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.DisplayAlerts = False

Dim tempath As String, Newprespath As String
tempath = "L:\Private\TEAM\Naga\Transportation - Structural Costs Automation\Template wsca.pptx"
Newprespath = "L:\Private\TEAM\Naga\Transportation - Structural Costs Automation\Warehousing - Structural Cost Analysis and Action Planning.pptx"

Dim pptapp As PowerPoint.Application
Set pptapp = CreateObject("Powerpoint.Application")
pptapp.Visible = True
Dim pres As PowerPoint.Presentation
Set pres = pptapp.Presentations.Open(tempath)
Dim ppslide As PowerPoint.Slide
Dim ppshape As PowerPoint.Shape

Worksheets("List of Locations").Select
Range("A1").Select
Dim lastrow As Integer
lastrow = Worksheets("List of Locations").Cells(Rows.Count, 1).End(xlUp).Row

Dim x As Integer, i As Integer
Dim slidecount As Integer
Dim str As String

For x = 2 To lastrow

    Worksheets("List of Locations").Select
    str = Range("A" & x)
    i = 4 * x - 6
    Set ppshape = pres.Slides(i).Shapes(1)
    With ppshape.TextFrame.TextRange
    .Font.Size = 34.7
    .Font.Color = RGB(72, 209, 210)
    .Font.Bold = msoTrue
    .Text = str
    End With
    
    Set ppshape = pres.Slides(i + 1).Shapes(1)
    With ppshape.TextFrame.TextRange
    .Font.Size = 28
    .Font.Color = vbBlack
    .Font.Bold = msoTrue
    .Text = str & " - Current Condition"
    End With
    
    Set ppshape = pres.Slides(i + 2).Shapes(1)
    With ppshape.TextFrame.TextRange
    .Font.Size = 28
    .Font.Color = vbBlack
    .Font.Bold = msoTrue
    .Text = str & " - Current Condition"
    End With
    
    Set ppshape = pres.Slides(i + 3).Shapes(1)
    With ppshape.TextFrame.TextRange
    .Font.Size = 28
    .Font.Color = vbBlack
    .Font.Bold = msoTrue
    .Text = str & " - Actions"
    End With
    
    Dim pf1 As PivotField, pf2 As PivotField, pf3 As PivotField, pf4 As PivotField, pf5 As PivotField
    
    Worksheets("Charts").Select
    Set pf1 = ActiveSheet.PivotTables("PivotTable2").PivotFields("Location")
    pf1.ClearAllFilters
    pf1.CurrentPage = str
    Set pf2 = ActiveSheet.PivotTables("PivotTable5").PivotFields("Location")
    pf2.ClearAllFilters
    pf2.CurrentPage = str
    Set pf3 = ActiveSheet.PivotTables("PivotTable8").PivotFields("Location")
    pf3.ClearAllFilters
    pf3.CurrentPage = str
    Set pf4 = ActiveSheet.PivotTables("PivotTable10").PivotFields("Location")
    pf4.ClearAllFilters
    pf4.CurrentPage = str
    
    Dim chartz
    chartz = Array("chart 1", "chart 3", "chart 5", "chart 7")
    
    ActiveSheet.ChartObjects(chartz(0)).Activate
    ActiveChart.ChartArea.Copy
    Set ppslide = pres.Slides(i + 1)
    ppslide.Select
    ActiveChart.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    ppslide.Shapes.Paste.Select
    With pptapp.ActiveWindow.Selection.ShapeRange
    .Left = 10
    .Height = 180
    .Top = 110
    .Width = 340
    End With
    
    ActiveSheet.ChartObjects(chartz(1)).Activate
    ActiveChart.ChartArea.Copy
    Set ppslide = pres.Slides(i + 1)
    ppslide.Select
    ActiveChart.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    ppslide.Shapes.Paste.Select
    With pptapp.ActiveWindow.Selection.ShapeRange
    .Left = 375
    .Height = 180
    .Top = 110
    .Width = 340
    End With
    
    ActiveSheet.ChartObjects(chartz(2)).Activate
    ActiveChart.ChartArea.Copy
    Set ppslide = pres.Slides(i + 1)
    ppslide.Select
    ActiveChart.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    ppslide.Shapes.Paste.Select
    With pptapp.ActiveWindow.Selection.ShapeRange
    .Left = 10
    .Height = 180
    .Top = 305
    .Width = 340
    End With
    
    ActiveSheet.ChartObjects(chartz(3)).Activate
    ActiveChart.ChartArea.Copy
    Set ppslide = pres.Slides(i + 1)
    ppslide.Select
    ActiveChart.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    ppslide.Shapes.Paste.Select
    With pptapp.ActiveWindow.Selection.ShapeRange
    .Left = 375
    .Height = 180
    .Top = 305
    .Width = 340
    End With
    
    Worksheets("YTD Data").Select
    Set pf5 = ActiveSheet.PivotTables("PivotTable2").PivotFields("Location")
    pf5.ClearAllFilters
    pf5.CurrentPage = str
    
    Range("M1:S18").Copy
    Set ppslide = pres.Slides(i + 2)
    ppslide.Shapes.PasteSpecial DataType:=2  '2 = ppPasteEnhancedMetafile
    Dim myShape As Object
    Set myShape = ppslide.Shapes(ppslide.Shapes.Count)
  
    With pptapp.ActiveWindow.Selection.ShapeRange
    .Left = 10
    .Height = 580
    .Top = 120
    .Width = 535
    End With
    
Next x

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.DisplayAlerts = True

pres.SaveAs Newprespath

End Sub

Option Explicit

Sub extracting()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.DisplayAlerts = False

Dim i As Long, j As Long, k As Long, m As Long
Dim lr As Long, lr1 As Long, lr2 As Long
lr = Worksheets("Filenames").Cells(Rows.Count, 1).End(xlUp).Row
lr2 = Worksheets("Filenames").Cells(Rows.Count, 2).End(xlUp).Row
Dim str As String, strm As String, strn As String

If lr2 <> 1 Then
k = lr2 + 1
Else
k = 1
End If

For i = k To lr
    str = Worksheets("Filenames").Range("A" & i)
    Dim fso As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream

    Set fso = New Scripting.FileSystemObject
    Set ts = fso.OpenTextFile("\\XSWS1831\Groups\HH\Private\TEAM\Naga\EDI Data Folder\" & str)

    Worksheets("Temp").Select
    Columns(1).ClearContents
    Range("A1").Select
    Do Until ts.AtEndOfStream
        ActiveCell.Value = ts.ReadLine
        ActiveCell.Offset(1, 0).Select
    Loop
    
    lr1 = Worksheets("Temp").Cells(Rows.Count, 1).End(xlUp).Row
    Worksheets("Output").Range("A" & i + 1).Value = str
        For j = 1 To lr1
            
            '<powerTrackReferenceNumber>1148038938</powerTrackReferenceNumber>
            If Left(Trim(Cells(j, 1).Value), 27) = "<powerTrackReferenceNumber>" Then
                strm = Trim(Cells(j, 1).Value)
                strn = Mid(strm, 28, Application.WorksheetFunction.Find("</", strm) - 28)
                Worksheets("Output").Range("AT" & i + 1).Value = strn
            End If
           
            '<fileReferenceNumber>200785370</fileReferenceNumber>
            If Left(Trim(Cells(j, 1).Value), 21) = "<fileReferenceNumber>" Then
                strm = Trim(Cells(j, 1).Value)
                strn = Mid(strm, 22, Application.WorksheetFunction.Find("</", strm) - 22)
                Worksheets("Output").Range("AU" & i + 1).Value = strn
            End If
            
            '<serviceAmount>28.86</serviceAmount>
            If Left(Trim(Cells(j, 1).Value), 15) = "<serviceAmount>" Then
                strm = Trim(Cells(j, 1).Value)
                strn = Mid(strm, 16, Application.WorksheetFunction.Find("</", strm) - 16)
                Worksheets("Output").Range("AS" & i + 1).Value = strn
            End If
            
        Next j
        
        'Save the extraction files after reading 25 files
        Worksheets("FileNames").Range("B" & i) = "Done"
        m = i Mod 25
        If m = 0 Then
        Worksheets("FileNames").Range("C" & i) = "Save"
        ActiveWorkbook.Save
        End If
        
Next i

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.DisplayAlerts = True

End Sub

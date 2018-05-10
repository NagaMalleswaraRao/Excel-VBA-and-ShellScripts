'Ask a question, if the answer is Yes, call the "Update macro".
Sub Refresh()
Dim a As Integer
a = MsgBox("Did you make sure mapping (RAW, Juarez mapping) is same as in CTS?", vbYesNo + vbQuestion, "Please Respond")

If a = vbYes Then
    Call update
End If
End Sub


Sub update()
Dim i As Integer
Dim objConn As Variant

'Refresh Data connections
For Each objConn In ThisWorkbook.Connections
    objConn.Refresh   
Next

'Refresh Pivot caches
Dim PC As PivotCache
For Each PC In ActiveWorkbook.PivotCaches
    PC.Refresh        
Next PC

'Input timestamp of refresh in a cell
Worksheets("Refresh").Range("J16") = Now

End Sub

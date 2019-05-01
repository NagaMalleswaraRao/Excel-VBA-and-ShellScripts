Sub error_checking()
'created two X and Y in columns C, D; trying to get output at
'column E (X/Y). Sample data is below.
'X | Y
'2 | 3
'3 | 0
'4 | 0
'0 | 9
'8 | 4

Dim i As Integer

For i = 4 To 10

On Error GoTo Err_handler
    Cells(i, 5) = Cells(i, 3) / Cells(i, 4)

Done:
GoTo Continue_if_no_error
    
Err_handler:
    On Error GoTo -1
    Cells(i, 5) = "NA"
    
Continue_if_no_error:
Next i

End Sub


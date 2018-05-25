Option Explicit

Sub FilenameExtract()
''' Reading all the file names in a file folder to a worksheet in "This Workbook"
''' Helps in: If any new file comes in, the extraction would work on the new ones only

Dim fList As String
Dim fList2() As String
Dim fName As String
Dim i As Long

''' Get the location \\USWS1031\Groups\HH\Private\TEAM\Naga\EDI files folder
fName = Dir("\\XSWS1091\Groups\HH\Private\TEAM\Naga\EDI files folder\*.txt")
i = 1

Do While fName <> ""
    ' Store the current file in the string fList.
    'fList = fList & vbNewLine & fName
    
    Worksheets("FileNames").Range("A" & i) = fName
    i = i + 1
    ' Get the next .txt file
    fName = Dir()
    ' The variable fName now contains the name of the next .txt file
Loop

End Sub

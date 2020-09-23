Attribute VB_Name = "mdlXMLFunc"
Option Explicit

Public Function ExportToXML(STable As String, FileName As String, Optional CTable As String, Optional Headers As Boolean) As Boolean
Dim TotalTable As String
Dim Whole As String
Dim i As Integer
Dim j As Integer
Dim TName As String
Dim FFile As Integer

Dim Doing As Integer
Dim CDoing As Integer

Dim TotalColumns As Integer
Dim TotalRows As String

Dim CInfo As String             'Column Info all
Dim TInfo As String             'All table items
Dim CTSplit() As String         'Split of the Table Items & Column Items
Dim SCInfo() As String          'Split of the column headers
Dim CInfoI As Integer           'Total column header strings
Dim STInfo() As String          'Split of the Table Info
Dim TInfoI As Integer           'Total Table Info strings

Dim CH() As String          'Column Header array
Dim LI() As String          'List item array

For i = 1 To TotalTables                'This find the information of the actuall table from the table name
    If Tables(i, 0) = STable Then
        TotalTable = Tables(i, 4)
        Exit For
    End If
Next i


CTSplit = Split(TotalTable, Chr(212) & Chr(232) & Chr(212))     'Splits the table information
CInfo = CTSplit(0)                                              'Gets the column header
TInfo = CTSplit(1)                                              'Gets the table info
CInfoI = CountSubStrings(CInfo, vbTab)                          'Finds the total column header items
SCInfo = Split(CInfo, vbTab)                                    'Splits the column header info
TInfoI = CountSubStrings(TInfo, vbTab)                          'Finds the total table items
STInfo = Split(TInfo, vbTab)                                    'Splits the table items

Doing = 1

TotalColumns = (CInfoI + 1) / 8             'Finds the ammount of columns there acutally is
TotalRows = (TInfoI + 1) / TotalColumns     'Finds the total ammount of rows there is


ReDim CH(TotalColumns, 7) As String         'Re dimensions the column header array
ReDim LI(TotalRows, TotalColumns) As String 'Re dimensions the list item array

CDoing = 1

For i = 0 To CInfoI Step 8                       'Loads each of the 8 components of a column into the array
    CH(CDoing, 0) = SCInfo(i)                    'Done in a step of 8 to speed up the loop
    CH(CDoing, 1) = SCInfo(i + 1)
    CH(CDoing, 2) = SCInfo(i + 2)
    CH(CDoing, 3) = SCInfo(i + 3)
    CH(CDoing, 4) = SCInfo(i + 4)
    CH(CDoing, 5) = SCInfo(i + 5)
    CH(CDoing, 6) = SCInfo(i + 6)
    CH(CDoing, 7) = SCInfo(i + 7)
    CDoing = CDoing + 1
Next i

Doing = 0

For i = 1 To TotalRows                      'Puts all of the list items into the LI array
    For j = 1 To TotalColumns
        LI(i, j) = STInfo(Doing)
        Doing = Doing + 1
    Next j
Next i

FFile = FreeFile                            'Finds a free file

STable = Replace(STable, " ", "-")          'Replaces any space for - cause XML cant handel spaces in tags

If DoesFileExist(FileName) = True Then      'If the file exists then it will ask if you want to write over
    If MsgBox(FileName & " already exists" & vbCrLf & "Do you wish to overwrite it?", vbExclamation + vbOKCancel, "Overwrite file?") = vbCancel Then
        Exit Function
    End If
End If
'Creates the whole XML file
Whole = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Whole = Whole & vbCrLf & "<!-- This was produced using BlackVault Databasing System, a product of Black Storm Software -->"
Whole = Whole & vbCrLf & "<" & STable & ">"
If Headers = True Then
    Whole = Whole & vbCrLf & vbTab & "<cheaders>"

    For i = 1 To TotalColumns
        Whole = Whole & vbCrLf & vbTab & vbTab & "<column>"
        Whole = Whole & vbCrLf & vbTab & vbTab & vbTab & "<name>" & CH(i, 0) & "</name>"
        Whole = Whole & vbCrLf & vbTab & vbTab & vbTab & "<required>" & CH(i, 1) & "</required>"
        Whole = Whole & vbCrLf & vbTab & vbTab & vbTab & "<dtype>" & CH(i, 2) & "</dtype>"
        Whole = Whole & vbCrLf & vbTab & vbTab & vbTab & "<cypte>" & CH(i, 3) & "</cypte>"
        Whole = Whole & vbCrLf & vbTab & vbTab & vbTab & "<sql>" & CH(i, 4) & "</sql>"
        Whole = Whole & vbCrLf & vbTab & vbTab & vbTab & "<vrules>" & CH(i, 5) & "</vrules>"
        Whole = Whole & vbCrLf & vbTab & vbTab & vbTab & "<default>" & CH(i, 6) & "</default>"
        Whole = Whole & vbCrLf & vbTab & vbTab & vbTab & "<maxlen>" & CH(i, 7) & "</maxlen>"
        Whole = Whole & vbCrLf & vbTab & vbTab & "</column>"
    Next i

    Whole = Whole & vbCrLf & vbTab & "</cheaders>"
    Whole = Whole & vbCrLf & vbTab & "<table>"
End If
For i = 1 To TotalRows
    If Headers = True Then
        Whole = Whole & vbCrLf & vbTab & vbTab & "<row>"
    Else
        Whole = Whole & vbCrLf & vbTab & vbTab & "<row>"
    End If
    For j = 1 To TotalColumns       'Adds the column header as the tags and then the list item and closes the tags
        If Headers = True Then
            Whole = Whole & vbCrLf & vbTab & vbTab & vbTab & "<" & CH(j, 0) & ">" & LI(i, j) & "</" & CH(j, 0) & ">"
        Else
            Whole = Whole & vbCrLf & vbTab & vbTab & "<" & CH(j, 0) & ">" & LI(i, j) & "</" & CH(j, 0) & ">"
        End If
    Next j
    If Headers = True Then
        Whole = Whole & vbCrLf & vbTab & vbTab & "</row>"
    Else
        Whole = Whole & vbCrLf & vbTab & "</row>"
    End If
Next i

If Headers = True Then
    Whole = Whole & vbCrLf & vbTab & "</table>"
End If
Whole = Whole & vbCrLf & "</" & STable & ">"            'Closes the XML

Open FileName For Output As #FFile
    Print #FFile, Trim(Whole)
Close #FFile

End Function


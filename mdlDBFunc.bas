Attribute VB_Name = "mdlDBFunc"
Option Explicit


Public Function LoadFromFile(SFile As String) As Boolean
Dim TSplit() As String
Dim DBProp As String
Dim Tables() As String
Dim Querries() As String

Dim FFile As Integer

Dim Whole As String
Dim Temp As String

FFile = FreeFile

Open SFile For Input As #FFile
    While Not EOF(FFile)
        Line Input #FFile, Temp
        Whole = Whole & vbNewLine & Temp
    Wend
Close #FFile

If Len(Whole) = 0 Then
    MsgBox "There was no data in the file being loaded", vbCritical, "Error: No Data"
    Exit Function
End If

TSplit = Split(Whole, Chr(252) & Chr(232) & Chr(212) & Chr(232) & Chr(252))

DBProperties = TSplit(0)
DBPreferences = TSplit(1)
DBRelationships = TSplit(2)
DBUsers = TSplit(3)
DBMacros = TSplit(4)
LoadAllQuerries (TSplit(5))
LoadAllTables (TSplit(6))

End Function

Public Function SaveToFile(SFile As String) As Boolean
Dim i As Integer
Dim Splitter As String
Dim FFile As Integer

SaveToFile = True

FFile = FreeFile

Splitter = Chr(252) & Chr(232) & Chr(212) & Chr(232) & Chr(252)

TotalText = DBProperties & Splitter & DBPreferences & Splitter & DBRelationships & Splitter & DBUsers & Splitter & DBMacros & Splitter

For i = 1 To TotalQuerries
    If i = TotalQuerries Then
        TotalText = TotalText & Querrys(i, 0) & Chr(222) & Chr(232) & Chr(222) & Querrys(i, 1)
    Else
        TotalText = TotalText & Querrys(i, 0) & Chr(222) & Chr(222) & Chr(222) & Querrys(i, 1) & Chr(222) & Chr(232) & Chr(222)
    End If
Next i

TotalText = TotalText & Splitter

For i = 1 To TotalTables
    If i = TotalTables Then
        TotalText = TotalText & Tables(i, 0) & Chr(222) & Chr(232) & Chr(222) & Tables(i, 1) & Chr(222) & Chr(232) & Chr(222) & Tables(i, 2) & Chr(222) & Chr(232) & Chr(222) & Tables(i, 3) & Chr(222) & Chr(232) & Chr(222) & Tables(i, 4)
    Else
        TotalText = TotalText & Tables(i, 0) & Chr(222) & Chr(232) & Chr(222) & Tables(i, 1) & Chr(222) & Chr(232) & Chr(222) & Tables(i, 2) & Chr(222) & Chr(232) & Chr(222) & Tables(i, 3) & Chr(222) & Chr(232) & Chr(222) & Tables(i, 4) & Chr(222) & Chr(232) & Chr(222)
    End If
Next i

Open SFile For Output As #FFile
    Print #FFile, Trim(TotalText)
Close #FFile

End Function


Public Function LoadAllQuerries(TQuerrys As String)
Dim NumItems As Integer
Dim NumQuerrys As Integer
Dim TSplit() As String
Dim i As Integer
Dim Doing As Integer
Dim QDoing As Integer
Doing = 0
QDoing = 1

NumItems = CountSubStrings(TQuerrys, Chr(222) & Chr(232) & Chr(222))
NumQuerrys = (NumItems + 1) / 3

TSplit = Split(TQuerrys, Chr(222) & Chr(232) & Chr(222))

ReDim Querries(NumQuerrys, 2)

TotalQuerries = NumQuerrys

For i = 0 To NumItems
    Querries(QDoing, Doing) = TSplit(i)
    Select Case Doing
        Case 0
            Doing = 1
        Case 1
            Doing = 2
        Case 2
            Doing = 0
            QDoing = QDoing + 1
    End Select
Next i
End Function


Public Function LoadAllTables(TTables As String)
Dim NumItems As Integer
Dim NumTables As Integer
Dim TSplit() As String
Dim i As Integer
Dim Doing As Integer
Dim TDoing As Integer
Doing = 0
TDoing = 1

NumItems = CountSubStrings(TTables, Chr(222) & Chr(232) & Chr(222))
NumTables = (NumItems + 1) / 5

TSplit = Split(TTables, Chr(222) & Chr(232) & Chr(222))

ReDim Tables(NumTables, 4)

TotalTables = NumTables

For i = 0 To NumItems
    
    If Doing = 0 And DoesTableExist(TSplit(i)) = True Then      'If a table with the same name is found then it will add <duplicate name> to it so that no problems are created
        Do Until DoesTableExist(TSplit(i)) = False
            TSplit(i) = TSplit(i) & " <Duplicate Name>"
        Loop
    End If
    
    Tables(TDoing, Doing) = TSplit(i)
    Select Case Doing
        Case 0
            Doing = 1
        Case 1
            Doing = 2
        Case 2
            Doing = 3
        Case 3
            Doing = 4
        Case 4
            Doing = 0
            TDoing = TDoing + 1
    End Select
Next i
End Function

Public Function DoesTableExist(TableName As String) As Boolean         'Used to make sure that no duplicate tables are created
Dim i As Integer
DoesTableExist = False
For i = 1 To TotalTables
    If TableName = Tables(i, 0) Then
        DoesTableExist = True
        Exit Function
    End If
Next i
End Function

Public Function DoesRelExist(TableOne As String, TableTwo As String) As Boolean
Dim i As Integer
Dim TRels() As String
Dim TRelsCount As Integer
Dim TSRel() As String
DoesRelExist = False

TRels = Split(DBRelationships, Chr(212) & Chr(232) & Chr(212))
TRelsCount = CountSubStrings(DBRelationships, Chr(212) & Chr(232) & Chr(212))

For i = 0 To TRelsCount
    TSRel = Split(TRels(i), vbTab)
    If TableOne = TSRel(0) And TableTwo = TSRel(2) Then
        DoesRelExist = True
    ElseIf TableOne = TSRel(2) And TableTwo = TSRel(0) Then
        DoesRelExist = True
    Else
        DoesRelExist = False
    End If
Next i
End Function

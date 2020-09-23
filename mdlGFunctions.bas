Attribute VB_Name = "mdlGFunctions"
Option Explicit

Public Function CountSubStrings(Target As String, Template As String) As Integer        'Used to count the number of splits in a string
Dim Pos1 As Integer
Dim Pos2 As Integer
Dim Count As Integer

If Len(Target) = 0 Or Len(Template) = 0 Or Len(Template) > Len(Target) Then
    CountSubStrings = -1
    Exit Function
End If
Count = 0
Pos2 = 1
Do
    Pos1 = InStr(Pos2, Target, Template, vbTextCompare)
    If Pos1 > 0 Then
        Count = Count + 1
        Pos2 = Pos1 + 1
    End If
Loop Until Pos1 = 0
CountSubStrings = Count
End Function

Public Function ContainsText(Template As String, Data As String) As Boolean
Dim Position As Integer
Position = InStr(1, Data, Template, vbTextCompare)
If Position = -1 Or Position = 0 Then
    ContainsText = False
Else
    ContainsText = True
End If
End Function

Public Function FindPosition(TStart As Integer, ToFind As String, Whole As String) As Integer       'Finds where a certain string is inside a string
FindPosition = InStr(TStart, Whole, ToFind)
End Function

Public Function DoesFileExist(DirPath As String) As Boolean         'Checks to see if a file exists
On Error GoTo ERROR
Dim Temp As String
Temp = Dir(DirPath, vbDirectory)
If Temp = "" Then
    DoesFileExist = False
Else
    DoesFileExist = True
End If
Exit Function
ERROR:
    DoesFileExist = False
End Function

Public Function CreatePWDHash(Password As String) As String
Dim MD5 As clsMD5
Dim ReversedP As String
Dim Hash As String
Dim ReversedH As String
Dim i As Integer
Set MD5 = New clsMD5
For i = 1 To Len(Password)                                          'Reverses the password to reduce the
    ReversedP = Mid(Password, i, 1) & ReversedP                     'likely hood of brute force attacks working
Next i

Hash = MD5.MD5(ReversedP)

Set MD5 = Nothing
For i = 1 To Len(Hash)                                              'Reverses the hash to make it even more
    ReversedH = Mid(Hash, i, 1) & ReversedH                         'difficult to crack
Next i
CreatePWDHash = ReversedH
End Function

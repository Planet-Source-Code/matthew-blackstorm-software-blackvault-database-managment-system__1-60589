Attribute VB_Name = "mdlDVald"
Option Explicit

Public Function Validate(DType As Integer, VRules As String, ToValidate As Variant) As Boolean
Dim NumRules As Integer                                                 'Holds the number of rules
Dim Rules() As String                                                   'Holds the split of the rules
Dim DRule As String                                                     'Holds the type of rule it is
Dim CData As Variant                                                    'Holds the data to check the data against
Dim i As Integer

Validate = True
NumRules = CountSubStrings(VRules, ";")
Rules = Split(VRules, ";")

For i = 0 To NumRules - 1
    If ContainsOperators(VRules) = True Then                            'If the rules contains operators
        If ValidateOperators(DType, VRules, ToValidate) = True Then     'this sends them to be validated
            Validate = True
            Exit Function
        Else
            Validate = False
            Exit Function
        End If
    Else                                                                'Validates rules with no operators
        DRule = UCase(Left(Rules(i), 2))
        CData = Mid(Rules(i), 4, Len(Rules(i)) - 4)
        If CheckRules(DType, CData, DRule, ToValidate) = True Then
            Validate = True
            Exit Function
        Else
            Validate = False
            Exit Function
        End If
    End If
Next i

End Function

Private Function ContainsOperators(Rule As String) As Boolean
ContainsOperators = False
If ContainsText("]and[", LCase(Rule)) = True Or ContainsText("]or[", LCase(Rule)) = True Or ContainsText("]not[", LCase(Rule)) = True Then
    ContainsOperators = True
End If
End Function

Private Function ValidateOperators(DType As Integer, ARule As String, Data As Variant) As Boolean
Dim Operator As Integer
Dim OPPos As Integer

Dim Rule As String

Dim FRule As String
Dim FValid As Boolean

Dim SRule As String
Dim SValid As Boolean

ValidateOperators = False

FValid = False
SValid = False

If FindPosition(1, "]and[", LCase(ARule)) > 1 Then                  'Checks if there is an And operator
    OPPos = FindPosition(1, "]and[", LCase(ARule))
    Operator = 0
ElseIf FindPosition(1, "]or[", LCase(ARule)) > 1 Then               'Checks if there is an Or operator
    OPPos = FindPosition(1, "]or[", LCase(ARule))
    Operator = 1
ElseIf FindPosition(1, "]not[", LCase(ARule)) > 1 Then              'Checks if there s a Not operator
    OPPos = FindPosition(1, "]not[", LCase(ARule))
    Operator = 2
End If

Rule = Left(ARule, 2)                                               'Find the Main rule to follow
FRule = Mid(ARule, 4, (OPPos - 4))                                  'Find the first rule

If Operator = 1 Then                                                'Find the second rule
    SRule = Mid(ARule, (OPPos + 4), (Len(ARule) - OPPos - 5))
Else
    SRule = Mid(ARule, (OPPos + 5), (Len(ARule) - OPPos - 6))
End If

Select Case Operator
    Case 0
        FValid = CheckRules(DType, FRule, Rule, Data)               'Validates the And operator
        SValid = CheckRules(DType, SRule, Rule, Data)
        If FValid And SValid = True Then
            ValidateOperators = True
        End If
    Case 1
        FValid = CheckRules(DType, FRule, Rule, Data)               'Validates the Or operator
        SValid = CheckRules(DType, SRule, Rule, Data)
        If FValid Or SValid = True Then
            ValidateOperators = True
        End If
    Case 2
        FValid = CheckRules(DType, FRule, Rule, Data)               'Validtes the Not operator
        SValid = CheckRules(DType, SRule, Rule, Data)
        If FValid = True And SValid = False Then
            ValidateOperators = True
        End If
End Select

End Function

Private Function CheckRules(DType As Integer, CheckData As Variant, Rule As String, Data As Variant) As Boolean
CheckRules = False
Select Case DType
    Case 0                  '*** String
        Select Case Rule
            Case "CT"       'Contains
                Select Case LCase(CStr(CheckData))
                    Case "{alpha}"
                        If LCase(CStr(CheckData)) = CStr(CheckData) Then
                            If ContainsAlpha(False, CStr(Data)) = True Then
                                CheckRules = True
                            End If
                        Else
                            If ContainsAlpha(True, CStr(Data)) = True Then
                                CheckRules = True
                            End If
                        End If
                    Case "{numeric}"
                        If LCase(CStr(CheckData)) = CStr(CheckData) Then
                            If ContainsNumeric(CStr(Data)) = True Then
                                CheckRules = True
                            End If
                        End If
                    Case Else
                        If ContainsText(CStr(CheckData), CStr(Data)) = True Then
                            CheckRules = True
                        End If
                End Select
            Case "DC"       'Doesnt Contain
                Select Case LCase(CStr(CheckData))
                    Case "{alpha}"
                        If LCase(CStr(CheckData)) = CStr(CheckData) Then
                            If ContainsAlpha(False, CStr(Data)) = False Then
                                CheckRules = True
                            End If
                        Else
                            If ContainsAlpha(True, CStr(Data)) = False Then
                                CheckRules = True
                            End If
                        End If
                    Case "{numeric}"
                        If LCase(CStr(CheckData)) = CStr(CheckData) Then
                            If ContainsNumeric(CStr(Data)) = False Then
                                CheckRules = True
                            End If
                        End If
                    Case Else
                        If ContainsText(CStr(CheckData), CStr(Data)) = False Then
                            CheckRules = True
                        End If
                End Select
            Case "PF"       'Prefix Of
                Select Case LCase(CStr(CheckData))
                    Case "{alpha}"
                        If LCase(CStr(CheckData)) = CStr(CheckData) Then
                            If ContainsAlpha(False, CStr(Left(Data, 1))) = True Then
                                CheckRules = True
                            End If
                        Else
                            If ContainsAlpha(True, CStr(Left(Data, 1))) = True Then
                                CheckRules = True
                            End If
                        End If
                    Case "{numeric}"
                        If LCase(CStr(CheckData)) = CStr(CheckData) Then
                            If ContainsNumeric(CStr(Left(Data, 1))) = True Then
                                CheckRules = True
                            End If
                        End If
                    Case Else
                        If ContainsText(CStr(CheckData), CStr(Left(Data, Len(CheckData)))) = True Then
                            CheckRules = True
                        End If
                    End Select
            Case "SF"       'Suffix Of
                Select Case LCase(CStr(CheckData))
                    Case "{alpha}"
                        If LCase(CStr(CheckData)) = CStr(CheckData) Then
                            If ContainsAlpha(False, CStr(Right(Data, 1))) = True Then
                                CheckRules = True
                            End If
                        Else
                            If ContainsAlpha(True, CStr(Right(Data, 1))) = True Then
                                CheckRules = True
                            End If
                        End If
                    Case "{numeric}"
                        If LCase(CStr(CheckData)) = CStr(CheckData) Then
                            If ContainsNumeric(CStr(Right(Data, 1))) = True Then
                                CheckRules = True
                            End If
                        End If
                    Case Else
                        If ContainsText(CStr(CheckData), CStr(Right(Data, Len(CheckData)))) = True Then
                            CheckRules = True
                        End If
                    End Select
        End Select
    Case 2, 3, 4, 5         '*** Byte, Integer, Long Int, Double
        CheckData = CLng(CheckData)
        Data = CLng(Data)
        If IsNumType(CDbl(Data), DType) = True Then
            Select Case Rule
                Case "LS"       'Less Than
                    If Data < CheckData Then CheckRules = True
                Case "GR"       'Greate Than
                    If Data > CheckData Then CheckRules = True
                Case "EQ"       'Equal To
                    If Data = CheckData Then CheckRules = True
                Case "NE"       'Not Equal To
                    If Data <> CheckData Then CheckRules = True
            End Select
        Else
            CheckRules = False
            Exit Function
        End If
    Case 6                  '*** Time
        If IsDate(CStr(Data)) = True Then
            Select Case Rule
                Case "BF"       'Before
                    If CDate(Data) < CDate(CheckData) Then
                        CheckRules = True
                        Exit Function
                    End If
                Case "AF"       'After
                    If CDate(Data) > CDate(CheckData) Then
                        CheckRules = True
                        Exit Function
                    End If
                Case "EQ"       'Equal To
                    If CDate(Data) = CDate(CheckData) Then
                        CheckRules = True
                        Exit Function
                    End If
                Case "NE"       'Not Equal To
                    If CDate(Data) <> CDate(CheckData) Then
                        CheckRules = True
                        Exit Function
                    End If

            End Select
        Else
            CheckRules = False
            Exit Function
        End If
    Case 7                  '*** Day
        If IsDay(CStr(Data)) = True Then
            If CheckDay(Rule, CStr(CheckData), CStr(Data)) = True Then
                CheckRules = True
                Exit Function
            End If
        Else
            CheckRules = False
            Exit Function
        End If
    Case 8                  '*** Date
        If IsDate(CStr(Data)) = True Then
            Select Case Rule
                Case "BF"       'Before
                    If CDate(CheckData) > CDate(Data) Then
                        CheckRules = True
                        Exit Function
                    End If
                Case "AF"       'After
                    If CDate(CheckData) < CDate(Data) Then
                        CheckRules = True
                        Exit Function
                    End If
                Case "EQ"       'Equal To
                    If CDate(CheckData) = CDate(Data) Then
                        CheckRules = True
                        Exit Function
                    End If
                Case "NE"       'Not Equal To
                    If CDate(CheckData) <> CDate(Data) Then
                        CheckRules = True
                        Exit Function
                    End If

            End Select
        Else
            CheckRules = False
            Exit Function
        End If
    Case 9                  '
        
    Case 10                 '*** URL
        If IsURLEmail(CStr(Data)) = True Then
            Select Case Rule
                Case "OT"
                    If ContainsText(CStr(CheckData), CStr(Data)) = True Then         'Of Domain Type
                        CheckRules = True
                        Exit Function
                    End If
                Case "NT"
                    If ContainsText(CStr(CheckData), CStr(Data)) = False Then        'Not Of Domain Type
                        CheckRules = True
                        Exit Function
                    End If
                Case "OD"
                    If ContainsText(CStr(CheckData), CStr(Data)) = True Then         'Of Domain
                        CheckRules = True
                        Exit Function
                    End If
                Case "ND"
                    If ContainsText(CStr(CheckData), CStr(Data)) = False Then        'Not Of Domain
                        CheckRules = True
                        Exit Function
                    End If
            End Select
        Else
            CheckRules = False
            Exit Function
        End If
    Case 11                 '*** Password
        Select Case Rule
            Case "CT"       'Contains
                Select Case LCase(CStr(CheckData))
                    Case "{alpha}"
                        If LCase(CStr(CheckData)) = CStr(CheckData) Then
                            If ContainsAlpha(False, CStr(Data)) = True Then
                                CheckRules = True
                            End If
                        Else
                            If ContainsAlpha(True, CStr(Data)) = True Then
                                CheckRules = True
                            End If
                        End If
                    Case "{numeric}"
                        If LCase(CStr(CheckData)) = CStr(CheckData) Then
                            If ContainsNumeric(CStr(Data)) = True Then
                                CheckRules = True
                            End If
                        End If
                    Case Else
                        If ContainsText(CStr(CheckData), CStr(Data)) = True Then
                            CheckRules = True
                        End If
                End Select
            Case "DC"       'Doesnt Contain
                Select Case LCase(CStr(CheckData))
                    Case "{alpha}"
                        If LCase(CStr(CheckData)) = CStr(CheckData) Then
                            If ContainsAlpha(False, CStr(Data)) = False Then
                                CheckRules = True
                            End If
                        Else
                            If ContainsAlpha(True, CStr(Data)) = False Then
                                CheckRules = True
                            End If
                        End If
                    Case "{numeric}"
                        If LCase(CStr(CheckData)) = CStr(CheckData) Then
                            If ContainsNumeric(CStr(Data)) = False Then
                                CheckRules = True
                            End If
                        End If
                    Case Else
                        If ContainsText(CStr(CheckData), CStr(Data)) = False Then
                            CheckRules = True
                        End If
                End Select
            Case "MN"       'Min Length
                If Len(CStr(Data)) > CInt(CheckData) Then CheckRules = True
            Case "MX"       'Max Length
                If Len(CStr(Data)) < CInt(CheckData) Then CheckRules = True
        End Select
End Select
End Function

Private Function ContainsAlpha(UpperCase As Boolean, Data As String) As Boolean
Dim i As Integer
ContainsAlpha = False
If UpperCase = False Then                               'Checks to see if there is uppercase in the string
    For i = 97 To 122                                   'loops through all uppercase letters
        If ContainsText(Chr(i), Data) = True Then
            ContainsAlpha = True
            Exit For
        End If
    Next i
Else
    For i = 65 To 90                                    'Checks to see if there is lowercase in the string
        If ContainsText(Chr(i), Data) = True Then       'loops through all lowercase letters
            ContainsAlpha = True
            Exit For
        End If
    Next i
End If
End Function

Private Function ContainsNumeric(Data As String) As Boolean
Dim i As Integer
ContainsNumeric = False
For i = 48 To 57                                        'Checks to see if there is numbers in the string
    If ContainsText(Chr(i), Data) = True Then
        ContainsNumeric = True
        Exit For
    End If
Next i
End Function

Public Function IsNumType(Data As Double, DType As Integer) As Boolean      'This checks to make sure that
Dim IntNum As Variant                                                       'the data is within the specified
IsNumType = False                                                           'paramaters of its data type
Select Case DType
    Case 2                                  'Checks that is is within a Byte
        If Data > 255 Or Data < 0 Then
            IsNumType = False
            Exit Function
        Else
            IntNum = CInt(Data)
            If Data = IntNum Then
                IsNumType = True
                Exit Function
            Else
                IsNumType = False
                Exit Function
            End If
        End If
    Case 3                                  'Checks that it is within an Integer
        If Data > 32768 Or Data < -32767 Then
            IsNumType = False
            Exit Function
        Else
            IntNum = CInt(Data)
            If Data = IntNum Then
                IsNumType = True
                Exit Function
            Else
                IsNumType = False
                Exit Function
            End If
        End If
    Case 4                                  'Makes sure its within a Long Integer
        If Data > 2147483648# Or Data < -2147483648# Then
            IsNumType = False
        Else
            IntNum = CLng(Data)
            If Data = IntNum Then
                IsNumType = True
                Exit Function
            Else
                IsNumType = False
                Exit Function
            End If
        End If
    Case 5
        
End Select
End Function

Private Function IsDay(Data As String) As Boolean           'Used to check if input is a valid day
IsDay = False
Select Case LCase(Data)
    Case "sunday"
        IsDay = True
    Case "monday"
        IsDay = True
    Case "tuesday"
        IsDay = True
    Case "wednesday"
        IsDay = True
    Case "thursday"
        IsDay = True
    Case "friday"
        IsDay = True
    Case "saturday"
        IsDay = True
    Case Else
        IsDay = False
End Select
End Function

Private Function CheckDay(Rule As String, CData As String, Data As String) As Boolean
CheckDay = False
Select Case Rule
    Case "BF"
        If GetDayPos(Data) < GetDayPos(CData) Then
            CheckDay = True
            Exit Function
        End If
    Case "AF"
        If GetDayPos(Data) > GetDayPos(CData) Then
            CheckDay = True
            Exit Function
        End If
    Case "EQ"
        If LCase(CData) = LCase(Data) Then
            CheckDay = True
            Exit Function
        End If
    Case "NE"
        If LCase(CData) <> LCase(Data) Then
            CheckDay = True
            Exit Function
        End If

End Select
End Function

Private Function GetDayPos(SDay As String) As Integer           'Used to get the day position with
Select Case LCase(SDay)                                         'Sunday being at position 0
    Case "sunday"
        GetDayPos = 0
    Case "monday"
        GetDayPos = 1
    Case "tuesday"
        GetDayPos = 2
    Case "wednesday"
        GetDayPos = 3
    Case "thursday"
        GetDayPos = 4
    Case "friday"
        GetDayPos = 5
    Case "saturday"
        GetDayPos = 6
End Select
End Function

Private Function IsURLEmail(URL As String) As Boolean
Dim objRegExp As RegExp                                     'Uses VBScrip Regular Expresions to prase
Set objRegExp = New RegExp                                  'the URL or Email
IsURLEmail = False
objRegExp.Test "^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,})+$"""
objRegExp.IgnoreCase = True
objRegExp.Global = False
IsURLEmail = objRegExp.Test(URL)
End Function

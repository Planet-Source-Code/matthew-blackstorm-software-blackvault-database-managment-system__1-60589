VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function PraseSQL(SQL As String, WasERROR As Boolean, ERROR As String, Optional CTable As String) As String
Select Case Left(Trim(UCase(SQL)), 6)
    Case "SELECT"
        PraseSQL = PraseStatment(SQL, 0, WasERROR, ERROR, CTable)
    Case "DELETE"
        PraseSQL = PraseStatment(SQL, 1, WasERROR, ERROR, CTable)
    Case "UPDATE"
        PraseSQL = PraseStatment(SQL, 2, WasERROR, ERROR, CTable)
    Case Else
        WasERROR = True
        ERROR = DisplayError(2001) 'Incorrect statment
        PraseSQL = "<error>"
        Exit Function
End Select
End Function


Private Function PraseStatment(SQL As String, StatType As Integer, WasERROR As Boolean, ERROR As String, Optional CTable As String) As String
Dim i As Integer
Dim MainStat As String
Dim FroPos As Integer
Dim WhePos As Integer
Dim OrdPos As Integer

Dim IsWheStat As Boolean
Dim WheStat As String

Dim IsOrdStat As Boolean
Dim OrdStat As String

Dim ContainsRel As Boolean

'Used to check if the SELECT and Where statment is valid
Dim SBOpened As Boolean
Dim RBOpened As Integer
Dim QOpened As Boolean

Dim TName As String
Dim OTable As Boolean
Dim FName As String
Dim Condition As String

Dim IType As Integer            'Used for the where statments

'Used to check if the FROM statment is valid
Dim FroStat As String
Dim FroTab() As String
Dim FroCount As Integer


ContainsRel = False

If ContainsText("FROM ", UCase(SQL)) = False Then
    WasERROR = True
    ERROR = DisplayError(2002)      'No from statment
    PraseStatment = "<error>"
    Exit Function
End If

If ContainsText("WHERE ", UCase(SQL)) = True Then
    IsWheStat = True
End If

If ContainsText("ORDER ", UCase(SQL)) = True Then
    IsOrdStat = True
End If

If IsWheStat = False And IsOrdStat = False Then
    FroPos = FindPosition(1, "FROM ", UCase(SQL))
    MainStat = Trim(Mid(SQL, 8, FroPos - 9))
    FroStat = Trim(Mid(SQL, FroPos + 5, Len(SQL) - FroPos - 5))
End If

If IsWheStat = True And IsOrdStat = False Then
    FroPos = FindPosition(1, "FROM ", UCase(SQL))
    WhePos = FindPosition(FroPos, "WHERE ", UCase(SQL))
    MainStat = Trim(Mid(SQL, 8, FroPos - 9))
    FroStat = Trim(Mid(SQL, FroPos + 5, WhePos - FroPos - 6))
    WheStat = Trim(Mid(SQL, WhePos + 6, Len(SQL) - 7))
End If

If IsWheStat = True And IsOrdStat = True Then
    FroPos = FindPosition(1, "FROM ", UCase(SQL))
    WhePos = FindPosition(FroPos, "WHERE ", UCase(SQL))
    OrdPos = FindPosition(WhePos, "ORDER ", UCase(SQL))
    MainStat = Trim(Mid(SQL, 8, FroPos - 9))
    FroStat = Trim(Mid(SQL, FroPos + 5, WhePos - FroPos - 6))
    WheStat = Trim(Mid(SQL, WhePos + 6, OrdPos - WhePos - 7))
    OrdStat = Trim(Mid(SQL, OrdPos + 7, Len(SQL) - OrdPos - 6))
End If

If IsWheStat = False And IsOrdStat = True Then
    FroPos = FindPosition(1, "FROM ", UCase(SQL))
    OrdPos = FindPosition(FroPos, "ORDER ", UCase(SQL))
    MainStat = Trim(Mid(SQL, 8, FroPos - 9))
    FroStat = Trim(Mid(SQL, FroPos + 5, OrdPos - FroPos - 6))
    OrdStat = Trim(Mid(SQL, OrdPos + 7, Len(SQL) - OrdPos - 6))
End If

SBOpened = False
RBOpened = 0
QOpened = False

OTable = True
If Trim(MainStat) <> "*" Then
    For i = 1 To Len(MainStat)                      'Loops through the main statment looking for errors
        Select Case Mid(MainStat, i, 1)
            Case "["                                'The opening of a new table or feild name
                If SBOpened = False Then            'Checks to see if there is one open already
                    SBOpened = True
                    If OTable = False Then
                        If Mid(MainStat, i - 1, 1) <> "!" Then
                            WasERROR = True                 'If ! is not used to seperate Table and Feild then will error
                            ERROR = DisplayError(2017)
                            PraseStatment = "<error>"
                            Exit Function
                        End If
                    End If
                Else
                    WasERROR = True                 'If so then it will error
                    ERROR = DisplayError(2007)
                    PraseStatment = "<error>"
                    Exit Function
                End If
            Case "]"                                'The closing of a new table or feild name
                If SBOpened = True Then             'Check to see if a bracket is open
                    SBOpened = False
                    If OTable = True And Mid(MainStat, i + 1, 2) = "![" Then
                        OTable = False
                    ElseIf OTable = False Then
                        OTable = True
                        If TableExists(TName) = True Then    'Checks to see if the table exists
                            If TableExists(TName, FName) = True Then     'Checks to see if the Feild exist
                                TName = ""
                                FName = ""
                                OTable = True
                            Else                    'If the Feild doesnt exist it will error
                                WasERROR = True
                                ERROR = DisplayError(2009)
                                PraseStatment = "<error>"
                                Exit Function
                            End If
                        Else                        'If a bracket is not open it will error
                            WasERROR = True
                            ERROR = DisplayError(2010)
                            PraseStatment = "<error>"
                            Exit Function
                        End If
                    Else
                        WasERROR = True             'No Feild referenced to will error
                        ERROR = DisplayError(2019)
                        PraseStatment = "<error>"
                        Exit Function
                    End If
                Else                                'If there is too many closing brackets it will error
                    WasERROR = True
                    ERROR = DisplayError(2008)
                    PraseStatment = "<error>"
                    Exit Function
                End If
            Case "("
                RBOpened = RBOpened + 1             'Used to make sure that there is the same ammount of
            Case ")"                                'Seperation bracets opend as closed
                RBOpened = RBOpened - 1
            Case Else
                If SBOpened = True Then             'If a Table/Feild name is open it will add the extra text
                    If OTable = True Then           'to the appropriate feild
                        TName = TName & Mid(MainStat, i, 1)
                    Else
                        FName = FName & Mid(MainStat, i, 1)
                    End If
                End If
        End Select
        If RBOpened < 0 Then                        'If there is an close brackets on its own then it will error
            WasERROR = True
            ERROR = DisplayError(2011)
            PraseStatment = "<error>"
            Exit Function
        End If
    Next i
End If

If RBOpened > 0 Then                            'If there is an open bracket without a close then it will error
    WasERROR = True
    ERROR = DisplayError(2012)
    PraseStatment = "<error>"
    Exit Function
End If

If SBOpened = True Then                         'If there is still a square bracket open it will error
    WasERROR = True
    ERROR = DisplayError(2013)
    PraseStatment = "<error>"
    Exit Function
End If

If ContainsText(",", FroStat) = True Then
    FroTab = Split(FroStat, ",")
    FroCount = 1
    If UBound(FroTab) > 1 Then
        WasERROR = True
        ERROR = DisplayError(2003)      'Querry across to many tables
        PraseStatment = "<error>"
        Exit Function
    End If
Else
    FroCount = 0
    FroStat = RemoveBrackets(FroStat)
    If TableExists(FroStat) = False Then
        WasERROR = True
        ERROR = DisplayError(2004)          'Table does not exist
        PraseStatment = "<error>"
        Exit Function
    End If
End If

If FroCount = 1 Then
    FroTab(0) = RemoveBrackets(FroTab(0))
    FroTab(1) = RemoveBrackets(FroTab(1))
    If TableExists(FroTab(0)) = True And TableExists(FroTab(1)) = True Then
        If DoesRelExist(FroTab(0), FroTab(1)) = True Then
            ContainsRel = True
        Else
            WasERROR = True
            ERROR = DisplayError(2005)
            PraseStatment = "<error>"        'One or both tables dont exist
            Exit Function
        End If
    Else
        WasERROR = True
        ERROR = DisplayError(2006)
        PraseStatment = "<error>"
        Exit Function
    End If
End If

IType = 1
RBOpened = 0
QOpened = False
SBOpened = False


For i = 1 To Len(WheStat)                       '*** Parses the WHERE Statment
    Select Case Mid(WheStat, i, 1)
        Case Chr(34)                            'Used to make sure that all quotes are closed
            If QOpened = True Then
                QOpened = False
            Else
                QOpened = False
            End If
        Case "["
            If SBOpened = False Then
                SBOpened = True
            Else
                WasERROR = True                 'Too many opening brackets
                ERROR = DisplayError(2015)
                PraseStatment = "<error>"
                Exit Function
            End If
        Case "]"
            If SBOpened = True Then
                SBOpened = False
                If IType = 1 And Mid(WheStat, i + 1, 2) = "![" Then
                    IType = 2
                ElseIf IType = 2 And Mid(WheStat, i + 1, 1) <> "!" Then
                    IType = 1
                Else
                    WasERROR = True                 'If ! is not used to seperate Table and Feild then will error
                    ERROR = DisplayError(2018)
                    PraseStatment = "<error>"
                    Exit Function
                End If
            Else
                WasERROR = True                 'Too many closing brackets
                ERROR = DisplayError(2016)
                PraseStatment = "<error>"
                Exit Function
            End If
        Case "("
            RBOpened = RBOpened + 1             'Used to make sure that there is the same ammount of
        Case ")"                                'Seperation bracets opend as closed
            RBOpened = RBOpened - 1
        Case Else
            If SBOpened = True Then             'If a Table/Feild name is open it will add the extra text
                If OTable = True Then           'to the appropriate feild
                    TName = TName & Mid(MainStat, i, 1)
                Else
                    FName = FName & Mid(MainStat, i, 1)
                End If
            End If
    End Select
    If RBOpened < 0 Then                        'If there is an close brackets on its own then it will error
        WasERROR = True
        ERROR = DisplayError(2020)
        PraseStatment = "<error>"
        Exit Function
    End If
Next i

If RBOpened > 0 Then                            'If there is an open bracket without a close then it will error
    WasERROR = True
    ERROR = DisplayError(2021)
    PraseStatment = "<error>"
    Exit Function
End If

If SBOpened = True Then                         'If there is still a square bracket open it will error
    WasERROR = True
    ERROR = DisplayError(20215)
    PraseStatment = "<error>"
    Exit Function
End If

If QOpened = True Then
    WasERROR = True
    ERROR = DisplayError(2014)
    PraseStatment = "<error>"
End If
PraseStatment = RunSQL(StatType, MainStat, FroStat, WheStat, OrdStat, WasERROR, ERROR, CTable)

End Function


Public Function RunSQL(StatType As Integer, MainStat As String, FroStat As String, WheStat As String, OrdStat As String, WasERROR As Boolean, ERROR As String, Optional CTable As String) As String
Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim RCHeaders() As String       'The Results Table column headers
Dim SRCHeaders As String        'Holds all the column headers as a string
Dim RTable() As String          'The Results Table
Dim SRTable As String           'Holds the whole table as a string
Dim RTColumns As Integer        'Total number of columns
Dim RTRows As Integer           'Total number of rows

Dim TCHeaders() As String       'The Real Tables column headers
Dim TTable() As String          'The Real Table
Dim TTColumns As Integer        'Total Number of columns
Dim TTRows As Integer           'Total number of rows

Dim LIDoing As Integer          'Used for working out which item is being used in an array

Dim MainLst() As String         'Lists all of the Main statments (ie. parts of the Select, Delete statments0
Dim MainCount As Integer        'Keeps a count of all the main statments
Dim MaiUpTo As Integer          'Used for working out where the next main statment is to go

Dim FroTab() As String          'List of all the from statments (only two possible ones)
Dim FroCount As Integer         'Keeps a count of all the from statments

Dim WheLst() As String          'List of all the where statment
Dim WheCount As Integer         'Keeps a count of all the where statments
Dim WheUpTo As Integer          'Holds where where upto count

Dim CInfo As String             'Column Info all
Dim TInfo As String             'All table items
Dim CTSplit() As String         'Split of the Table Items & Column Items
Dim SCInfo() As String          'Split of the column headers
Dim CInfoI As Integer           'Total column header strings
Dim STInfo() As String          'Split of the Table Info
Dim TInfoI As Integer           'Total Table Info strings
Dim CDoing As Integer

Dim RCount As Integer           'Total nubmer of rows returned
Dim RStart As Integer           'The row to start at
Dim RStop As Integer            'The row to stop at

'Used for spliting up the where statment
Dim RBO As Boolean              'Used to check if a round bracket is open i. ( )
Dim SBo As Boolean              'Used to check if a square bracket is open ie. [ ]
Dim QOp As Boolean              'Used to check if a quote is open

Dim SPos As Integer             'Used to hold where the item spliter is ie ]![ in a Where statment

Dim LineIsOk As Boolean         'Used to telling if a row in a table is ok when all where comparisions are done

'Used for spliting up the where statment
Dim ComPos As Integer

RCount = 0
RStart = 0
RStop = 0

MainCount = CountSubStrings(MainStat, ",") + 1
If MainCount < 1 Then
    MainCount = 1
End If
If StatType = 2 Then
    ReDim MainLst(MainCount, 2)
Else
    ReDim MainLst(MainCount, 1)
End If

If RemoveBrackets(MainStat) = "*" Then
    MainStat = RemoveBrackets(MainStat)
End If

If ContainsText(",", MainStat) = True And Trim(MainStat) <> "*" Then
    'MainLst = Split(MainStat, ",")
    QOp = False
    SBo = False
    RBO = False
    ComPos = 1
    MaiUpTo = 0
    
    For i = 1 To Len(MainStat)
        Select Case Mid(MainStat, i, 1)
            Case Chr(34)
                If QOp = True Then
                    QOp = False
                Else
                    QOp = True
                End If
            Case "["
                If QOp = False Then
                    SBo = True
                End If
            Case "]"
                If QOp = False Then
                    SBo = False
                End If
            Case "("
                If QOp = False Then
                    RBO = True
                End If
            Case ")"
                If QOp = False Then
                    RBO = False
                End If
            Case ","
                If QOp = False And SBo = False And RBO = False Then
                    MainLst(MaiUpTo, 0) = Mid(MainStat, FindPosition(ComPos, "[", MainStat) + 1, FindPosition(ComPos, "]![", MainStat) - FindPosition(ComPos, "[", MainStat) - 1)
                    SPos = FindPosition(ComPos, "]![", MainStat) + 3
                    MainLst(MaiUpTo, 1) = Mid(MainStat, SPos, FindPosition(SPos, "]", MainStat) - SPos)
                    If StatType = 2 Then
                        SPos = FindPosition(SPos, "]", MainStat) + 1
                        MainLst(0, 2) = Mid(MainStat, SPos, Len(MainStat) - SPos + 1)
                    End If
                    MaiUpTo = MaiUpTo + 1
                    ComPos = i
                End If
        End Select
    Next i
    If QOp = False And SBo = False And RBO = False Then
        MainLst(MaiUpTo, 0) = Mid(MainStat, FindPosition(ComPos, "[", MainStat) + 1, FindPosition(ComPos, "]![", MainStat) - FindPosition(ComPos, "[", MainStat) - 1)
        SPos = FindPosition(ComPos, "]![", MainStat) + 3
        MainLst(MaiUpTo, 1) = Mid(MainStat, SPos, FindPosition(SPos, "]", MainStat) - SPos)
        If StatType = 2 Then
            SPos = FindPosition(SPos, "]", MainStat) + 1
            MainLst(0, 2) = Mid(MainStat, SPos, Len(MainStat) - SPos + 1)
        End If
        MaiUpTo = MaiUpTo + 1
        ComPos = i
    End If
    MainCount = MaiUpTo
    
ElseIf MainStat <> "*" And UCase(Trim(Left(MainStat, 8))) <> "COUNT(*)" Then
    ReDim MainLst(1, 1)
    MainCount = 1
    MainLst(0, 0) = Mid(MainStat, FindPosition(1, "[", MainStat) + 1, FindPosition(1, "]![", MainStat) - FindPosition(1, "[", MainStat) - 1)
    SPos = FindPosition(1, "]![", MainStat) + 3
    MainLst(0, 1) = Mid(MainStat, SPos, FindPosition(SPos, "]", MainStat) - SPos)
    MainCount = 1
End If

If WheStat <> "" And UCase(Trim(Left(WheStat, 5))) <> "ROWS(" Then     'Checks to see if there is any where statments
    WheLst = Split(WheStat, ",")           'If there is then it starts this to split them
    WheCount = CountSubStrings(WheStat, ",") + 1
    
    ReDim WheLst(WheCount, 2)               'Redimensions the where list
    If ContainsText(",", WheStat) = True Then       'If there is more than one where statment then it will
        QOp = False                                 'Start looping through them and adding them to the list
        SBo = False
        RBO = False
        ComPos = 1
        WheUpTo = 0
        
        For i = 1 To Len(WheStat)
            Select Case Mid(WheStat, i, 1)
                Case Chr(34)
                    If QOp = True Then
                        QOp = False
                    Else
                        QOp = True
                    End If
                Case "["
                    If QOp = False Then
                        SBo = True
                    End If
                Case "]"
                    If QOp = False Then
                        SBo = False
                    End If
                Case "("
                    If QOp = False Then
                        RBO = True
                    End If
                Case ")"
                    If QOp = False Then
                        RBO = False
                    End If
                Case ","
                    If QOp = False And SBo = False And RBO = False Then
                        WheLst(WheUpTo, 0) = Mid(WheStat, FindPosition(ComPos, "[", WheStat) + 1, FindPosition(ComPos, "]![", WheStat) - FindPosition(ComPos, "[", WheStat) - 1)
                        SPos = FindPosition(ComPos, "]![", WheStat) + 3
                        WheLst(WheUpTo, 1) = Mid(WheStat, SPos, FindPosition(SPos, "]", WheStat) - SPos)
                        SPos = FindPosition(SPos, "]", WheStat) + 1
                        If FindPosition(i + 1, ",", WheStat) = True Then
                            WheLst(WheUpTo, 2) = Mid(WheStat, ComPos, Len(WheStat) - SPos)
                        Else
                            WheLst(WheUpTo, 2) = Mid(WheStat, SPos, FindPosition(SPos, ",", WheStat) - SPos)
                        End If
                        WheUpTo = WheUpTo + 1
                        ComPos = i
                    End If
            End Select
        Next i
        If QOp = False And SBo = False And RBO = False Then
            WheLst(WheUpTo, 0) = Mid(WheStat, FindPosition(ComPos, "[", WheStat) + 1, FindPosition(ComPos, "]![", WheStat) - FindPosition(ComPos, "[", WheStat) - 1)
            SPos = FindPosition(ComPos, "]![", WheStat) + 3
            WheLst(WheUpTo, 1) = Mid(WheStat, SPos, FindPosition(SPos, "]", WheStat) - SPos)
            SPos = FindPosition(SPos, "]", WheStat) + 1
            WheLst(WheUpTo, 2) = Mid(WheStat, SPos, Len(WheStat) - SPos + 1)
            WheUpTo = WheUpTo + 1
            ComPos = i
        End If
        WheCount = WheUpTo
    Else
        ReDim WheLst(1, 2)
        WheLst(0, 0) = Mid(WheStat, FindPosition(1, "[", WheStat) + 1, FindPosition(1, "]![", WheStat) - FindPosition(1, "[", WheStat) - 1)
        SPos = FindPosition(1, "]![", WheStat) + 3
        WheLst(0, 1) = Mid(WheStat, SPos, FindPosition(SPos, "]", WheStat) - SPos)
        SPos = FindPosition(SPos, "]", WheStat) + 1
        WheLst(0, 2) = Mid(WheStat, SPos, Len(WheStat) - SPos + 1)
        WheCount = 1
    End If
End If

If MainCount > 0 Or MainStat = "*" Or UCase(MainStat) = "COUNT(*)" Then
    ReDim RCHeaders(MainCount, 7)
    ReDim RTable(MainCount, 1)
    
    If CTable <> "" Then                                            'If there is a temporary table passed it will runn it again
        CTSplit = Split(CTable, Chr(212) & Chr(232) & Chr(212))     'Splits the table information
    Else
        FroStat = RemoveBrackets(FroStat)
        For i = 1 To TotalTables
            If FroStat = Tables(i, 0) Then
                If Tables(i, 1) = 0 Then
                    WasERROR = True
                    ERROR = DisplayError(2022)
                    RunSQL = "<error>"
                    Exit Function
                End If
                CTSplit = Split(Tables(i, 4), Chr(212) & Chr(232) & Chr(212))
                Exit For
            End If
        Next i
    End If
    
    CInfo = CTSplit(0)                                              'Gets the column header
    TInfo = CTSplit(1)                                              'Gets the table info
    CInfoI = CountSubStrings(CInfo, vbTab)                          'Finds the total column header items
    SCInfo = Split(CInfo, vbTab)                                    'Splits the column header info
    TInfoI = CountSubStrings(TInfo, vbTab)                          'Finds the total table items
    STInfo = Split(TInfo, vbTab)                                    'Splits the table items
    
    TTColumns = (CInfoI + 1) / 8                                    'Finds the ammount of columns there acutally is
    TTRows = (TInfoI + 1) / TTColumns                               'Finds the total ammount of rows there is

    ReDim TCHeaders(TTColumns, 7)
    ReDim TTable(TTRows, TTColumns)
    
    LIDoing = 0
    CDoing = 0
    
    For i = 0 To TTRows - 1                                      'Puts all of the list items into the Temp Table Array array
        For j = 0 To TTColumns - 1
            TTable(i, j) = STInfo(LIDoing)
            LIDoing = LIDoing + 1
        Next j
    Next i
    
    For i = 0 To CInfoI Step 8                              'Loads each of the 8 components of a column into the array
        TCHeaders(CDoing, 0) = SCInfo(i)                    'Done in a step of 8 to speed up the loop
        TCHeaders(CDoing, 1) = SCInfo(i + 1)
        TCHeaders(CDoing, 2) = SCInfo(i + 2)
        TCHeaders(CDoing, 3) = SCInfo(i + 3)
        TCHeaders(CDoing, 4) = SCInfo(i + 4)
        TCHeaders(CDoing, 5) = SCInfo(i + 5)
        TCHeaders(CDoing, 6) = SCInfo(i + 6)
        TCHeaders(CDoing, 7) = SCInfo(i + 7)
        CDoing = CDoing + 1
    Next i
    
    ReDim RTable(TTRows, CDoing - 1)
    For i = 0 To TTRows - 1                                 'Loops through the rows
        LineIsOk = False                                    'Resets the LineIsOk Variable
        For j = 0 To WheCount                               'Loops through the where statments
            For k = 0 To TTColumns - 1                      'Loops through the columns looking for the same column as in the where statment
                If WheCount > 0 Then                        'If there is more than one where statment the int will start the checks
                    If WheLst(j, 1) = TCHeaders(k, 0) Then
                        If WhereCompare(CInt(TCHeaders(k, 2)), TTable(i, k), WheLst(j, 2), WasERROR, ERROR) = True Then
                            LineIsOk = True
                        Else
                            LineIsOk = False
                        End If
                    End If
                Else
                    LineIsOk = True                         'If there is no where statments then it will say the line is ok
                End If
            Next k
        Next j
        If LineIsOk = True Then                                 'If the line is ok then it will add the selected items into the RunSQl function variable
            If UCase(Trim(Left(MainStat, 8))) = "COUNT(*)" Then
                RCount = RCount + 1
            Else
                Select Case StatType
                    Case 0
                        If MainStat <> "*" Then                                 'If its not selecting all the items
                            For j = 0 To MainCount                              'Loops through the main items
                                For k = 0 To TTColumns - 1                      'Loops through the column headers
                                    If MainLst(j, 1) = TCHeaders(k, 0) Then     'Checks to see if the selected main item and column header match
                                        If SRTable = "" Then                    'If so then it will add the selcted item
                                            SRTable = TTable(i, k)
                                        Else
                                            SRTable = SRTable & vbTab & TTable(i, k)      'If there is already data in then it will use a tab seperator
                                        End If
                                    End If
                                Next k
                            Next j
                        Else                                                'If its selecting all the column headers
                            For j = 0 To TTColumns - 1                      'Loops through the column headers
                                If SRTable = "" Then                        'If so then it will add the selcted item
                                    SRTable = TTable(i, j)
                                Else
                                    SRTable = SRTable & vbTab & TTable(i, j)      'If there is already data in then it will use a tab seperator
                                End If
                            Next j
                        End If
                    Case 1
                        For j = 0 To TTColumns - 1
                            If SRTable = "" Then                     'if so then it will add the selcted item
                                SRTable = TTable(i, j)
                            Else
                                SRTable = SRTable & vbTab & TTable(i, j)    'If there is already data in then it will use a tab seperator
                            End If
                            TTable(i, j) = "<deleted>"
                        Next j
                    Case 2
                        
                End Select
            End If
        End If
    Next i
Else
    
End If

If MainStat = "*" Then
    MainCount = TTColumns
End If

Select Case StatType
    Case 0
        For i = 0 To MainCount - 1
            For j = 0 To TTColumns - 1
                    If MainStat <> "*" And UCase(MainStat) <> "COUNT(*)" Then       'If its any old SELECT statment
                        If MainLst(i, 1) = TCHeaders(j, 0) Then
                            If SRCHeaders = "" Then
                                SRCHeaders = TCHeaders(j, 0) & vbTab & TCHeaders(j, 1) & vbTab & TCHeaders(j, 2) & vbTab & TCHeaders(j, 3) & vbTab & TCHeaders(j, 4) & vbTab & TCHeaders(j, 5) & vbTab & TCHeaders(j, 6) & vbTab & TCHeaders(j, 7)
                            Else
                                SRCHeaders = SRCHeaders & vbTab & TCHeaders(j, 0) & vbTab & TCHeaders(j, 1) & vbTab & TCHeaders(j, 2) & vbTab & TCHeaders(j, 3) & vbTab & TCHeaders(j, 4) & vbTab & TCHeaders(j, 5) & vbTab & TCHeaders(j, 6) & vbTab & TCHeaders(j, 7)
                            End If
                        End If
                    ElseIf UCase(MainStat) <> "COUNT(*)" Then       'If is not a count statment then
                        If SRCHeaders = "" Then
                            SRCHeaders = TCHeaders(j, 0) & vbTab & TCHeaders(j, 1) & vbTab & TCHeaders(j, 2) & vbTab & TCHeaders(j, 3) & vbTab & TCHeaders(j, 4) & vbTab & TCHeaders(j, 5) & vbTab & TCHeaders(j, 6) & vbTab & TCHeaders(j, 7)
                        Else
                            SRCHeaders = SRCHeaders & vbTab & TCHeaders(j, 0) & vbTab & TCHeaders(j, 1) & vbTab & TCHeaders(j, 2) & vbTab & TCHeaders(j, 3) & vbTab & TCHeaders(j, 4) & vbTab & TCHeaders(j, 5) & vbTab & TCHeaders(j, 6) & vbTab & TCHeaders(j, 7)
                        End If
                    End If
            Next j
            If MainStat = "*" Then      'If its a select all query then the j for loop will add all column headers
                Exit For                'To the results so no need to go again
            End If
        Next i
    Case 1
        
    Case 2
        
End Select

If UCase(Trim(Left(MainStat, 8))) = "COUNT(*)" Then
    SRCHeaders = "COUNT(*)" & vbTab & "0" & vbTab & "2" & vbTab & "0" & vbTab & "<none>" & vbTab & "<none>" & vbTab & "0" & vbTab & "0"
    SRTable = RCount
End If

If SRCHeaders <> "" Then
    RunSQL = SRCHeaders & Chr(212) & Chr(232) & Chr(212)
End If


If SRTable <> "" Then
    RunSQL = RunSQL & SRTable
Else
    RunSQL = RunSQL & "<null>"          'If there is no list items returned then it will return a null statment
End If
End Function

Private Function WhereCompare(DType As Integer, Data As String, CData As String, WasERROR As Boolean, ERROR As String) As Boolean
Dim TCData As String

WhereCompare = False

Select Case DType
    Case 0                                                  'If its a text compare then will run these checks
        If Left(CData, 1) = "=" Then                        'Equal to check
            TCData = Trim(Mid(CData, 3, Len(CData) - 3))
            If ContainsText(TCData, Data) = True Then
                WhereCompare = True
            End If
        End If
        
        If Left(CData, 2) = "<>" Then                       'Not equal to check
            TCData = Trim(Mid(CData, 4, Len(CData) - 4))
            If ContainsText(TCData, Data) = False Then
                WhereCompare = True
            End If
        End If
        
        If Left(UCase(CData), 2) = "UC" Then                'Upper case check
            TCData = Trim(Mid(CData, 4, Len(CData) - 4))
            If ContainsText(UCase(TCData), UCase(Data)) = True Then
                WhereCompare = True
            End If
        End If
        
        If Left(UCase(CData), 2) = "LC" Then                'Lower case check
            TCData = Trim(Mid(CData, 4, Len(CData) - 4))
            If ContainsText(LCase(TCData), LCase(Data)) = True Then
                WhereCompare = True
            End If
        End If
    Case 1
        
    Case 2, 3, 4, 5
        If Left(CData, 1) = "=" Then                        'Equal to check
            TCData = Trim(Mid(CData, 2, Len(CData) - 1))
            If Val(TCData) = Val(Data) Then
                WhereCompare = True
            End If
        End If
        
        If Left(CData, 2) = "<>" Then                       'Not equal to check
            TCData = Trim(Mid(CData, 3, Len(CData) - 2))
            If Val(TCData) <> Val(Data) Then
                WhereCompare = True
            End If
        End If
        
        If Left(CData, 1) = "<" Then                        'Equal to check
            TCData = Trim(Mid(CData, 1, Len(CData) - 1))
            If Val(TCData) < Val(Data) Then
                WhereCompare = True
            End If
        End If
        
        If Left(CData, 1) = ">" Then                       'Not equal to check
            TCData = Trim(Mid(CData, 1, Len(CData) - 1))
            If Val(TCData) > Val(Data) Then
                WhereCompare = True
            End If
        End If
        
        If Left(CData, 2) = "<=" Then                        'Equal to check
            TCData = Trim(Mid(CData, 3, Len(CData) - 2))
            If Val(TCData) <= Val(Data) Then
                WhereCompare = True
            End If
        End If
        
        If Left(CData, 2) = ">=" Then                       'Not equal to check
            TCData = Trim(Mid(CData, 3, Len(CData) - 2))
            If Val(TCData) >= Val(Data) Then
                WhereCompare = True
            End If
        End If
    Case 6, 7, 8
        
    Case 9
        
    Case 10
        
    Case 11
End Select
End Function

Private Function RemoveBrackets(ReFrom As String) As String
Dim Temp As String
Temp = ReFrom
Temp = Replace(Temp, "[", "")
Temp = Replace(Temp, "]", "")
Temp = Replace(Temp, vbNewLine, "")
Temp = Replace(Temp, vbCrLf, "")
Temp = Replace(Temp, Chr(13), "")
RemoveBrackets = Trim(Temp)
End Function

Private Function TableExists(TableName As String, Optional FeildName As String) As Boolean          'Used to make sure that no duplicate tables are created
Dim i As Integer
Dim j As Integer
Dim TTable() As String
Dim Columns As Integer
Dim TColumns() As String
TableExists = False
For i = 1 To TotalTables
    If TableName = Tables(i, 0) Then
        TableExists = True
        If FeildName <> "" Then
            TableExists = False                                             'Resets it if its searching for a feild aswell
            TTable = Split(Tables(i, 4), Chr(212) & Chr(232) & Chr(212))    'Splits the table info up
            Columns = CountSubStrings(TTable(0), vbTab)                     'Counts the number of column headers
            TColumns = Split(TTable(0), vbTab)                              'Splits the columns headers up
            For j = 0 To Columns Step 8                                     'Loops through to find the if the Key column exists in table
                If TColumns(j) = FeildName Then
                    TableExists = True
                    Exit Function
                End If
            Next j
        End If
    End If
Next i
End Function

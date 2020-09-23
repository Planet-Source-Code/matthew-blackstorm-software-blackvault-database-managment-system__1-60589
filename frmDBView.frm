VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDBView 
   Caption         =   "Black Vault -"
   ClientHeight    =   10080
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   ScaleHeight     =   10080
   ScaleWidth      =   12825
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   3480
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwDB 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5530
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAdd 
         Caption         =   "Add Record"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCT 
         Caption         =   "Close Table"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileXML 
         Caption         =   "Export To XML"
      End
      Begin VB.Menu mnuFileHTML 
         Caption         =   "Export To HMTL"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileET 
         Caption         =   "Export Table"
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTE 
      Caption         =   "TableEdit"
      Visible         =   0   'False
      Begin VB.Menu mnuTEAdd 
         Caption         =   "&Add New Record"
      End
      Begin VB.Menu mnuTEBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTEDR 
         Caption         =   "&Delete Record"
      End
   End
End
Attribute VB_Name = "frmDBView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SavedTable As String
Dim ColumnHeaders As String

Public TableName As String

Dim TotalColumns As Integer

'Column Information
Dim CItem() As String


Dim XPos As Integer
Dim YPos As Integer
Dim RClicked As Boolean

Private Sub Form_Load()
lvwDB.Width = frmDBView.Width - 100
lvwDB.Height = frmDBView.Height - 700
RClicked = False
End Sub

Private Sub Form_Resize()
If frmDBView.WindowState <> vbMinimized Then
    lvwDB.Width = frmDBView.Width - 100
    lvwDB.Height = frmDBView.Height - 700
End If
End Sub


Public Function CreateColumn(CName As String, DRequired As Boolean, DType As Integer, CType As Integer, SQL As String, VRules As String, Default As String, MaxLen As Integer) As Boolean
Dim i As Integer
Dim j As Integer
Dim TCItem() As String                  'Temporary Variable for Column Names
TotalColumns = TotalColumns + 1         'Add New Column to List

ReDim TCItem(TotalColumns, 7)           'Redimension All Temporary Column Variables

If TotalColumns <> 1 Then
    For i = 1 To TotalColumns - 1       'Loop to Collect All Column Data
        For j = 0 To 7
            TCItem(i, j) = CItem(i, j)
        Next j
    Next i
    
    ReDim CItem(TotalColumns, 7)        'Redimension All Column Variables
    
    For i = 1 To TotalColumns           'Loop to Replace All Data Into Column Variables
        For j = 0 To 7
            CItem(i, j) = TCItem(i, j)
        Next j
    Next i
    
    CItem(TotalColumns, 0) = CName      'Put All New Column Information in to Column Variables
    CItem(TotalColumns, 1) = DRequired
    CItem(TotalColumns, 2) = DType
    CItem(TotalColumns, 3) = CType
    CItem(TotalColumns, 4) = SQL
    CItem(TotalColumns, 5) = VRules
    CItem(TotalColumns, 6) = Default
    CItem(TotalColumns, 7) = MaxLen
    lvwDB.ColumnHeaders.Add , , CName
Else
    ReDim CItem(1, 7)
    CItem(1, 0) = CName
    CItem(1, 1) = DRequired
    CItem(1, 2) = DType
    CItem(1, 3) = CType
    CItem(1, 4) = SQL
    CItem(1, 5) = VRules
    CItem(1, 6) = Default
    CItem(1, 7) = MaxLen
    lvwDB.ColumnHeaders.Add , , CName
End If
End Function

Public Function SaveTable() As String
Dim Total As String
Dim i As Integer
Dim j As Integer

Dim DReq As Integer

For i = 1 To lvwDB.ColumnHeaders.Count
    If i <> lvwDB.ColumnHeaders.Count Then
        Total = Total & CItem(i, 0) & vbTab & CItem(i, 1) & vbTab & CItem(i, 2) & vbTab & CItem(i, 3) & vbTab & CItem(i, 4) & vbTab & CItem(i, 5) & vbTab & CItem(i, 6) & vbTab & CItem(i, 7) & vbTab
    Else
        Total = Total & CItem(i, 0) & vbTab & CItem(i, 1) & vbTab & CItem(i, 2) & vbTab & CItem(i, 3) & vbTab & CItem(i, 4) & vbTab & CItem(i, 5) & vbTab & CItem(i, 6) & vbTab & CItem(i, 7)
    End If
Next i

Total = Total & Chr(212) & Chr(232) & Chr(212)

For i = 1 To lvwDB.ListItems.Count
    For j = 1 To lvwDB.ColumnHeaders.Count
        If i = lvwDB.ListItems.Count And j = lvwDB.ColumnHeaders.Count Then
            If j = 1 Then
                Total = Total & lvwDB.ListItems(i).Text
            Else
                Total = Total & lvwDB.ListItems(i).ListSubItems(j - 1).Text
            End If
        Else
            If j = 1 Then
                Total = Total & lvwDB.ListItems(i).Text & vbTab
            Else
                Total = Total & lvwDB.ListItems(i).ListSubItems(j - 1).Text & vbTab
            End If
        End If
    Next j
Next i
SaveTable = Total
End Function

Public Function LoadTable(TName As String, Table As String) As Boolean
Dim CInfo As String
Dim TInfo As String
Dim CTSplit() As String
Dim SCInfo() As String
Dim CInfoI As Integer
Dim STInfo() As String
Dim TInfoI As Integer

Dim i As Integer
Dim j As Integer

Dim CName As String
Dim DRequired As Boolean
Dim DaType As Integer
Dim CType As Integer
Dim CSQL As String
Dim VRules As String
Dim CDefault As String
Dim CLen As String
Dim Doing As Integer
Dim ColumnUpTo As Integer

Dim Delimiter As Integer

UnloadCurrentTable

TableName = TName

Delimiter = FindPosition(1, Chr(212) & Chr(232) & Chr(212), Table) + 4


CTSplit = Split(Table, Chr(212) & Chr(232) & Chr(212))
CInfo = CTSplit(0)
If Delimiter < Len(Table) And Delimiter - 4 > 0 Then
    TInfo = CTSplit(1)
End If

ColumnHeaders = CInfo
CInfoI = CountSubStrings(CInfo, vbTab)
SCInfo = Split(CInfo, vbTab)
TInfoI = CountSubStrings(TInfo, vbTab)
STInfo = Split(TInfo, vbTab)
Doing = 1
For i = 0 To CInfoI Step 8
    CName = SCInfo(i)
    If Val(SCInfo(i + 1)) = 0 Then
        DRequired = False
    Else
        DRequired = True
    End If
    DaType = Val(SCInfo(i + 2))
    CType = Val(SCInfo(i + 3))
    CSQL = SCInfo(i + 4)
    VRules = SCInfo(i + 5)
    CDefault = SCInfo(i + 6)
    CLen = SCInfo(i + 7)
    CreateColumn CName, DRequired, DaType, CType, CSQL, VRules, CDefault, Val(CLen)
Next i

Doing = 1

If Len(Trim(TInfo)) = 0 Then
    Exit Function
End If

For i = 0 To TInfoI
    If Doing = 1 Then
        Set LItem = lvwDB.ListItems.Add(, , STInfo(i))
    End If
    If Doing <> 1 Then
        LItem.ListSubItems.Add , , STInfo(i)
    End If
    If Doing = lvwDB.ColumnHeaders.Count Then
        Doing = 1
    Else
        Doing = Doing + 1
    End If
Next i

End Function

Public Function UnloadCurrentTable()
ReDim CItem(0)
ColumnHeaders = ""
SavedTable = ""
TableName = ""
TotalColumns = 0

lvwDB.ListItems.Clear
lvwDB.ColumnHeaders.Clear
End Function

Private Sub Form_Unload(Cancel As Integer)
UnloadCurrentTable
End Sub

Private Sub lvwDB_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub lvwDB_ItemClick(ByVal Item As MSComctlLib.ListItem)
If RClicked = True Then
    PopupMenu mnuTE, , XPos, YPos
    RClicked = False
End If
End Sub

Private Sub lvwDB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    XPos = x
    YPos = y
    RClicked = True
End If
End Sub

Private Sub mnuFileAdd_Click()
frmAddRecord.LoadColumns (ColumnHeaders)
frmAddRecord.Show vbModal
End Sub

Private Sub mnuFileExit_Click()
UnloadCurrentTable
frmDBView.Hide
End Sub

Private Sub mnuFileXML_Click()
Dim FileName As String
cdlFile.Filter = "XML File (*.XML)|*.xml"
cdlFile.FileName = ""
cdlFile.ShowSave
FileName = cdlFile.FileName
If FileName = "" Then
    Exit Sub
Else
    ExportToXML TableName, FileName, False
End If
End Sub


Private Sub mnuTEAdd_Click()
frmAddRecord.LoadColumns (ColumnHeaders)
frmAddRecord.Show vbModal
End Sub

Private Sub mnuTEDR_Click()                         'Deletes selected records
Dim Selected As Integer
Dim Records As Integer
Dim PSplit() As String
Dim i As Integer
Dim LItems As Integer

LItems = lvwDB.ListItems.Count


For i = 1 To LItems                                 'Loop removes all selected itesm
    If i > LItems Then                              'If some are removed loop out of wack
        Exit For                                    'so statment quits loop if it trys to
    End If                                          'one outside of total count
        
    If lvwDB.ListItems(i).Selected = True Then
        lvwDB.ListItems.Remove (i)
        LItems = lvwDB.ListItems.Count
    End If
Next i

For i = 1 To TotalTables                            'Finds the table and saves it to the records
    If frmDBView.TableName = Tables(i, 0) Then      'and updates the count of records
        Tables(i, 4) = SaveTable
        PSplit = Split(Tables(i, 3), vbTab)
        Records = CInt(PSplit(0)) - 1
        PSplit(3) = Format(Date, "dd/MM/yy")
        Tables(i, 3) = Records & vbTab & PSplit(1) & vbTab & PSplit(2) & vbTab & PSplit(3)
        Exit For
    End If
Next i

If SaveToFile(CFile) = False Then                    'Saves it to the file and updates the list
    MsgBox "There was an error updating the database with the new record", vbOKOnly
Else
    frmAddRecord.Hide
    frmMain.LoadTablesToList
End If
RClicked = False
End Sub

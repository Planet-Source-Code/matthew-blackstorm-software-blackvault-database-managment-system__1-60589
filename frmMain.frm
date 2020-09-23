VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Black Vault - "
   ClientHeight    =   8640
   ClientLeft      =   165
   ClientTop       =   465
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   11835
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wskWeb 
      Left            =   360
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskServer 
      Left            =   360
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdMacros 
      Caption         =   "View Macros"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   5640
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Black Vault Database Files (*.BVD)|*.bvd"
   End
   Begin VB.CommandButton cmdQuerries 
      Caption         =   "View Querries"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdTables 
      Caption         =   "View Tables"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imlDBIcons 
      Left            =   8160
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0338
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2295
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlDBIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "# Of Rows"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date Created"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date Modified"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save Database As"
      End
      Begin VB.Menu mnuFileLoad 
         Caption         =   "Load Database"
      End
      Begin VB.Menu mnuFileNew 
         Caption         =   "New Database"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileNT 
         Caption         =   "New Table"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "Import Table"
      End
      Begin VB.Menu mnuFileRel 
         Caption         =   "Edit Relationships"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Begin VB.Menu mnuServerStart 
         Caption         =   "Start Server"
      End
      Begin VB.Menu mnuServerStop 
         Caption         =   "Stop Server"
      End
   End
   Begin VB.Menu mnuPref 
      Caption         =   "&Preferences"
      Begin VB.Menu mnuPrefOptions 
         Caption         =   "Database Options"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
   End
   Begin VB.Menu mnuTo 
      Caption         =   "TableOption"
      Visible         =   0   'False
      Begin VB.Menu mnuTOOpen 
         Caption         =   "Open Table"
      End
      Begin VB.Menu mnuTODel 
         Caption         =   "Delete Table"
      End
      Begin VB.Menu mnuTOBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTOXML 
         Caption         =   "Export As XML"
      End
      Begin VB.Menu mnuTOHTML 
         Caption         =   "Export As HTML"
      End
      Begin VB.Menu mnuTOBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTOProperties 
         Caption         =   "Table Properties"
      End
   End
   Begin VB.Menu mnuQueries 
      Caption         =   "&Queries"
      Begin VB.Menu mnuQueriesQQ 
         Caption         =   "Quick Query"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAboutBV 
         Caption         =   "About BlackVault"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelItem As Integer

Dim RClicked As Boolean

Dim XPos As Integer
Dim YPos As Integer

Private Sub cmdTables_Click()
LoadTablesToList
End Sub

Private Sub Form_Load()
lvwItems.Width = frmMain.Width - 1540
lvwItems.Height = frmMain.Height - 700
RClicked = False
End Sub

Private Sub Form_Resize()                       'Resizes the listview when the main window is resized
If frmMain.WindowState <> vbMinimized Then
    lvwItems.Width = frmMain.Width - 1540
    lvwItems.Height = frmMain.Height - 700
End If
End Sub

Public Function LoadTablesToList()              'Loads all of the tables to the main listview
Dim i As Integer
Dim TSplit() As String
lvwItems.ListItems.Clear

For i = 1 To TotalTables
    Set LItem = lvwItems.ListItems.Add(, , Tables(i, 0), , 1)
    TSplit = Split(Tables(i, 3), vbTab)
    LItem.ListSubItems.Add , , TSplit(0)
    LItem.ListSubItems.Add , , TSplit(1)
    LItem.ListSubItems.Add , , TSplit(2)
    LItem.ListSubItems.Add , , TSplit(3)
Next i
End Function

Private Sub lvwItems_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub lvwItems_DblClick()
If SelItem <> 0 Then
    OpenTable (SelItem)
End If
SelItem = 0
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
SelItem = Item.Index
If RClicked = True Then
    PopupMenu mnuTo, , XPos, YPos
    RClicked = False
End If
End Sub

Private Sub lvwItems_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    RClicked = True
    XPos = x + 1440
    YPos = y
End If
End Sub

Private Function OpenTable(SelItem As Integer)
Dim TableName As String
Dim i As Integer
TableName = lvwItems.ListItems(SelItem).Text
For i = 1 To TotalTables
    If Tables(i, 0) = TableName Then
        frmDBView.Caption = "BlackVault Database - " & TableName
        frmDBView.Show
        If Tables(i, 1) = 0 Then
            frmDBView.mnuFileAdd.Enabled = False
            frmDBView.mnuTEAdd.Enabled = False
            frmDBView.mnuTEDR.Enabled = False
        Else
            frmDBView.mnuFileAdd.Enabled = True
            frmDBView.mnuTEAdd.Enabled = True
            frmDBView.mnuTEDR.Enabled = True
        End If
        frmDBView.LoadTable TableName, Tables(i, 4)
        Exit Function
    End If
Next i
End Function

Private Sub mnuAboutBV_Click()
frmAbout.DisplayMe
End Sub

Private Sub mnuFileLoad_Click()
Dim FileName As String
cdlFile.FileName = ""
cdlFile.Filter = "Black Vault Database Files (*.BVD)|*.bvd"
cdlFile.ShowOpen
FileName = cdlFile.FileName
If FileName = "" Then
    Exit Sub
Else
    CFile = FileName
    LoadFromFile (FileName)
    LoadTablesToList
    frmMain.Caption = "BlackVault Database - " & Left(cdlFile.FileTitle, Len(cdlFile.FileTitle) - 4)
End If
End Sub

Private Sub mnuFileNT_Click()
frmATable.Show vbModal
End Sub

Private Sub mnuFileRel_Click()
frmRelationships.LoadAllRelationships
End Sub

Private Sub mnuFileSaveAs_Click()
Dim SaveTo As String
cdlFile.ShowSave
SaveTo = cdlFile.FileName
If SaveTo <> "" Then
    If SaveToFile(SaveTo) = False Then
        MsgBox "Error Saving DB To Location"
        Exit Sub
    Else
        CFile = SaveTo
    End If
End If
End Sub

Private Sub mnuQueriesQQ_Click()
frmQQuery.Show vbModal
End Sub

Private Sub mnuTOOpen_Click()
OpenTable (SelItem)
End Sub

Private Sub mnuTOProperties_Click()
frmProperties.DisplayTableProp (SelItem)
End Sub

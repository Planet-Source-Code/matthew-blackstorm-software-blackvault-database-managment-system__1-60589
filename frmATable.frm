VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmATable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create New Table"
   ClientHeight    =   7335
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   9360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlFile 
      Left            =   7680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Black Vault Database Files (*.BVD)|*.bvd"
   End
   Begin VB.TextBox txtTName 
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   28
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtNotes 
      Height          =   2415
      Left            =   4560
      MaxLength       =   400
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   4800
      Width           =   4695
   End
   Begin VB.Frame famGeneral 
      Height          =   2415
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   4215
      Begin VB.ComboBox cboDefault 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmATable.frx":0000
         Left            =   1320
         List            =   "frmATable.frx":000A
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox txtValidation 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2040
         Width           =   2400
      End
      Begin VB.CommandButton cmdValidation 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3795
         TabIndex        =   9
         Top             =   2040
         Width           =   300
      End
      Begin VB.TextBox txtDefault 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox txtMax 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "0"
         ToolTipText     =   "Max lenght of feild (0 is unlimited)"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton optRNo 
         Caption         =   "No"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   960
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optRYes 
         Caption         =   "Yes"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cboType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmATable.frx":001B
         Left            =   1320
         List            =   "frmATable.frx":0043
         TabIndex        =   2
         Text            =   "String"
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblValidation 
         Caption         =   "Validation Rules:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblDefault 
         Caption         =   "Default Value:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblMax 
         Caption         =   "Max Lenght:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblReq 
         Caption         =   "Required:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblType 
         Caption         =   "Feild Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame famLookup 
      Height          =   2415
      Left            =   120
      TabIndex        =   22
      Top             =   4800
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdSQL 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   255
         Left            =   3800
         TabIndex        =   12
         Top             =   600
         Width           =   300
      End
      Begin VB.ComboBox cboLType 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmATable.frx":00A4
         Left            =   1320
         List            =   "frmATable.frx":00B1
         TabIndex        =   10
         Text            =   "Text"
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtRSource 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   600
         Width           =   2400
      End
      Begin VB.Label lblLType 
         Caption         =   "Feild Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblRSource 
         Caption         =   "Record Source:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdDown 
      Enabled         =   0   'False
      Height          =   300
      Left            =   25
      Picture         =   "frmATable.frx":00CF
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2520
      Width           =   300
   End
   Begin VB.CommandButton cmdUp 
      Enabled         =   0   'False
      Height          =   315
      Left            =   25
      Picture         =   "frmATable.frx":0299
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1800
      Width           =   300
   End
   Begin MSComctlLib.ImageList imgTable 
      Left            =   0
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmATable.frx":0463
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwColumns 
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   6800
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgTable"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Column Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Column Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbsProperties 
      Height          =   2895
      Left            =   0
      TabIndex        =   13
      Top             =   4440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Look Up"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNotes 
      Caption         =   "Table Notes:"
      Height          =   255
      Left            =   4560
      TabIndex        =   29
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label lblTName 
      Caption         =   "Table Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   27
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAdd 
         Caption         =   "Add New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCreate 
         Caption         =   "Create Table"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuChange 
      Caption         =   "Change"
      Visible         =   0   'False
      Begin VB.Menu mnuChangeSetPrimary 
         Caption         =   "Set As Primary Key"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeDel 
         Caption         =   "Delete Column"
      End
   End
End
Attribute VB_Name = "frmATable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelTab As Integer                                               'Stores the selected tab

Dim SelItem As Integer                                              'Stores the selected list item
    
Dim TableName As String                                             'Stores the table name
Dim TabProp() As String                                             'Array holds the table properties
Dim TotalColumns As Integer                                         'Stores the total number of columns
Dim TNotes As String                                                'Stores the table notes / descrip

Dim XPos As Integer                                                 'Used for position of the cursor on the listview
Dim YPos As Integer
Dim RClicked As Boolean

Dim PKey As String                                                  'Stores the name of the primary key

Private Sub cboDefault_Click()                                      'Sets or resets the default value of a boolean field
If cboDefault.Text <> "True" And cboDefault.Text <> "False" Then
    cboDefault.Text = "True"
    TabProp(SelItem, 6) = 1
End If
If cboDefault.Text = "True" Then
    TabProp(SelItem, 6) = 1
End If
If cboDefault.Text = "False" Then
    TabProp(SelItem, 6) = 0
End If
End Sub

Private Sub cboDefault_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 38 And KeyCode <> 40 Then
    KeyCode = 0
End If
End Sub

Private Sub cboDefault_KeyPress(KeyAscii As Integer)
If KeyAscii <> 38 And KeyAscii <> 40 Then
    KeyAscii = 0
End If
End Sub

Private Sub cboLType_Click()                    'Sets the List type
If cboLType.Text <> "Text" And cboLType.Text <> "Value List" And cboLType.Text <> "Querry" Then
    cboLType.Text = "Text"
    TabProp(SelItem, 3) = 0
End If
If cboLType.Text = "Text" Then
    TabProp(SelItem, 3) = 0
    cmdSQL.Enabled = False
    txtRSource.Enabled = False
End If
If cboLType.Text = "Value List" Then
    TabProp(SelItem, 3) = 1
    cmdSQL.Enabled = False
    txtRSource.Enabled = True
End If
If cboLType.Text = "Querry" Then
    TabProp(SelItem, 3) = 2
    cmdSQL.Enabled = True
    txtRSource.Enabled = True
End If
End Sub

Private Sub cboLType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 38 And KeyCode <> 40 Then
    KeyCode = 0
End If
End Sub

Private Sub cboLType_KeyPress(KeyAscii As Integer)
If KeyAscii <> 38 And KeyAscii <> 40 Then
    KeyAscii = 0
End If
End Sub

Private Sub cboType_Click()
Dim i As Integer
Dim Found As Boolean

Found = False

lvwColumns.ListItems(SelItem).ListSubItems(1).Text = cboType.Text

For i = 0 To 11
    If cboType.Text = cboType.List(i) Then
        TabProp(SelItem, 2) = i
        Found = True
    End If
Next i

If cboType.Text <> "Bit" Then
        cboDefault.Visible = False
        txtDefault.Visible = True
    Else
        cboDefault.Visible = True
        txtDefault.Visible = False
End If

If cboType.Text <> "Bit" Then
    cmdValidation.Enabled = True
Else
    cmdValidation.Enabled = False
End If

If Found = False Then
    cboType.Text = "String"
    TabProp(SelItem, 2) = 1
End If
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode <> 38 And KeyCode <> 40 Then
    KeyCode = 0
End If
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
If KeyAscii <> 38 And KeyAscii <> 40 Then
    KeyAscii = 0
End If
End Sub

Private Sub cmdDown_Click()                                 'Moves the selected item down
MoveItemDown
End Sub

Private Sub cmdUp_Click()                                   'Moves the selected item up
MoveItemUp
End Sub

Private Sub cmdValidation_Click()                           'Used to edit the validation rules
Dim SelType As Integer
SelType = Val(TabProp(SelItem, 2))
frmVRules.SetForType SelType, TabProp(SelItem, 5)
End Sub

Public Sub SetVRules(Rules As String)                       'Displays the rules in the textbox
txtValidation.Text = Rules
TabProp(SelItem, 5) = Rules
End Sub

Private Sub Form_Load()
SelTab = 1
End Sub

Private Sub lvwColumns_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub lvwColumns_ItemClick(ByVal Item As MSComctlLib.ListItem)    'Sets the properties of the new selected item
SelItem = Item.Index
SetAllItems
If RClicked = True Then
    If Trim(lvwColumns.ListItems(SelItem).Text) = "" Then
        mnuChangeSetPrimary.Enabled = False
    Else
        mnuChangeSetPrimary.Enabled = True
    End If
    PopupMenu mnuChange, , XPos, YPos
    RClicked = False
End If
End Sub

Private Sub lvwColumns_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Or KeyCode = 46 Then
    DeleteSelected
End If
End Sub

Private Sub lvwColumns_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    XPos = x + 360                              'Compensates for the position of the listview on form
    YPos = y + 480
    RClicked = True
End If
End Sub

Private Sub SetAllItems()                       'Used to refresh all the table properties
Dim Name As String                              'of the new selected item in the listview
Dim FType As String
Dim ColumnType As Integer
Dim SQL As String
Dim MaxLen As Integer
Dim DefaultValue As String
Dim ValidationRules As String

Name = TabProp(SelItem, 0)
FType = TabProp(SelItem, 2)                     'Field Type
ColumnType = Val(TabProp(SelItem, 3))           'Column Type
SQL = TabProp(SelItem, 4)                       'Record Source
ValidationRules = TabProp(SelItem, 5)           'Validation Rules
DefaultValue = TabProp(SelItem, 6)              'Default Value
MaxLen = Val(TabProp(SelItem, 7))               'Max Len

txtName.Text = TabProp(SelItem, 0)              'Column Name
If Name = PKey And PKey <> "" Then              'If its a primary key then it must be required
    optRYes.Enabled = False
    optRYes.Value = True
    optRNo.Enabled = False
Else
    optRYes.Enabled = True
    optRNo.Enabled = True
End If

If Val(TabProp(SelItem, 1)) = 0 Then            'Sets the required settings
    optRNo.Value = True
    optRYes.Value = False
Else
    optRNo.Value = False
    optRYes.Value = True
End If

cboType.Text = cboType.List(FType)

If ColumnType = 1 Then
    txtDefault.Visible = False
    cboDefault.Visible = True
Else
    txtDefault.Visible = True
    cboDefault.Visible = False
End If

txtMax.Text = Val(TabProp(SelItem, 7))

txtDefault.Text = DefaultValue
txtValidation.Text = ValidationRules
If SelItem = 1 Then
    cmdUp.Enabled = False
    If TotalColumns > 1 Then
        cmdDown.Enabled = True
    Else
        cmdDown.Enabled = False
    End If
Else
    cmdUp.Enabled = True
    If SelItem = TotalColumns Then
        cmdDown.Enabled = False
    Else
        cmdDown.Enabled = True
    End If
End If

If SelItem = TotalColumns And TotalColumns <> 1 Then
    cmdUp.Enabled = True
    cmdDown.Enabled = False
End If

End Sub

Private Sub mnuChangeSetPrimary_Click()                     'Changes the selected item to the primary key
Dim NimgLst As ImageList                                    'Changes the listview icon position
Set lvwColumns.SmallIcons = NimgLst
lvwColumns.SmallIcons = imgTable
lvwColumns.ListItems(SelItem).SmallIcon = 1
PKey = lvwColumns.ListItems(SelItem).Text
TabProp(SelItem, 1) = 1                                     'Makes the data required
optRYes.Value = True
optRYes.Enabled = False
optRNo.Enabled = False
End Sub

Private Sub mnuFileAdd_Click()                              'Adds a new column
NewColumn
End Sub

Private Sub mnuFileCreate_Click()                           'Starts the creation of the table
CreateTable
End Sub

Private Sub mnuFileExit_Click()
frmATable.Hide
End Sub

Private Sub optRNo_Click()                                  'Sets the item to not required
TabProp(SelItem, 1) = 0
End Sub

Private Sub optRYes_Click()                                 'Sets the item to required
TabProp(SelItem, 1) = 1
End Sub

Private Sub tbsProperties_Click()
If SelTab <> tbsProperties.SelectedItem.Index Then
    SelTab = tbsProperties.SelectedItem.Index
    Select Case SelTab
        Case 1
            famGeneral.Visible = True
            famLookup.Visible = False
        Case 2
            famGeneral.Visible = False
            famLookup.Visible = True
    End Select
End If
End Sub

Private Sub txtDefault_Change()
'Put mdlDVald call in here
TabProp(SelItem, 6) = txtDefault.Text
End Sub

Private Sub txtMax_Change()
If txtMax.Text <> "" Then               'Checks to see if anything was entered, and if so it is numeric
    If IsNumeric(txtMax.Text) = False Or Val(txtMax.Text) < 0 Then
        MsgBox "Please enter a whole integer", vbCritical, "Error: Please enter a real number"
        txtMax.Text = "0"
        TabProp(SelItem, 7) = 0
    Else
        TabProp(SelItem, 7) = Val(txtMax.Text)
    End If
End If
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 And KeyAscii <> 8 Then            'Makes sure that its only a number that is pressed
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    Else
        TabProp(SelItem, 7) = Val(txtMax.Text)
    End If
Else
    TabProp(SelItem, 7) = Val(txtMax.Text)
End If
End Sub

Private Sub NewColumn()
Dim TTP() As String
Dim i As Integer
Dim j As Integer

Properties True

TotalColumns = TotalColumns + 1
If TotalColumns > 2 Then
    cmdUp.Enabled = True
    cmdDown.Enabled = True
End If

ReDim TTP(TotalColumns, 7) As String

If TotalColumns > 0 Then
    For i = 1 To TotalColumns - 1
        For j = 0 To 7
            TTP(i, j) = TabProp(i, j)
        Next j
    Next i
    ReDim TabProp(TotalColumns, 7)
Else
    ReDim TabProp(TotalColumns, 7)
End If

For i = 1 To TotalColumns
    For j = 0 To 7
        TabProp(i, j) = TTP(i, j)
    Next j
Next i

TabProp(TotalColumns, 0) = ""           'Column Name
TabProp(TotalColumns, 1) = "0"          'Required
TabProp(TotalColumns, 2) = "0"          'Field Type
TabProp(TotalColumns, 3) = "0"          'Column Type
TabProp(TotalColumns, 4) = "<none>"     'Record Source
TabProp(TotalColumns, 5) = "<none>"     'Validation Rules
TabProp(TotalColumns, 6) = ""           'Default Value
TabProp(TotalColumns, 7) = "0"          'Max Len

Set LItem = lvwColumns.ListItems.Add(, , " ")
LItem.ListSubItems.Add , , "String"

SelItem = TotalColumns
Set lvwColumns.SelectedItem = lvwColumns.ListItems(SelItem)
SetAllItems
txtName.SetFocus
End Sub

Private Sub DeleteSelected()
Dim i As Integer
Dim ItemDoing As Integer
Dim TTabProp() As String

ItemDoing = 1
ReDim TTabProp(TotalColumns - 1, 7)
If SelItem > 0 Then
    If lvwColumns.ListItems(SelItem).Text = PKey Then
        PKey = ""
    End If
    lvwColumns.ListItems.Remove (SelItem)
    If lvwColumns.ListItems.Count > SelItem Then
        Set lvwColumns.SelectedItem = lvwColumns.ListItems(SelItem)
        SetAllItems
    End If
    If lvwColumns.ListItems.Count < SelItem And lvwColumns.ListItems.Count > 0 Then
        SelItem = SelItem - 1
        Set lvwColumns.SelectedItem = lvwColumns.ListItems(SelItem)
    End If
    If lvwColumns.ListItems.Count = 0 Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
        Properties False
        SelItem = 0
    End If
    If lvwColumns.ListItems.Count = 1 Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    End If
    For i = 1 To TotalColumns - 1               'Puts the column headers into the temp array
        If i = SelItem Then                     'This IF skips over the deleted column
            ItemDoing = ItemDoing + 1
        End If
        TTabProp(i, 0) = TabProp(ItemDoing, 0)
        TTabProp(i, 1) = TabProp(ItemDoing, 1)
        TTabProp(i, 2) = TabProp(ItemDoing, 2)
        TTabProp(i, 3) = TabProp(ItemDoing, 3)
        TTabProp(i, 4) = TabProp(ItemDoing, 4)
        TTabProp(i, 5) = TabProp(ItemDoing, 5)
        TTabProp(i, 6) = TabProp(ItemDoing, 6)
        TTabProp(i, 7) = TabProp(ItemDoing, 7)
        ItemDoing = ItemDoing + 1
    Next i
    
    TotalColumns = TotalColumns - 1
    ReDim TabProp(TotalColumns, 7)              'Redimensions the column headers array
    
    For i = 1 To TotalColumns                   'Puts all the column headers back into the array
        TabProp(i, 0) = TTabProp(i, 0)
        TabProp(i, 1) = TTabProp(i, 1)
        TabProp(i, 2) = TTabProp(i, 2)
        TabProp(i, 3) = TTabProp(i, 3)
        TabProp(i, 4) = TTabProp(i, 4)
        TabProp(i, 5) = TTabProp(i, 5)
        TabProp(i, 6) = TTabProp(i, 6)
        TabProp(i, 7) = TTabProp(i, 7)
    Next i
End If
End Sub

Private Sub txtName_Change()
Dim i As Integer                    'Checks to see if that column name already exists
For i = 1 To TotalColumns
    If LCase(lvwColumns.ListItems(i).Text) = LCase(txtName.Text) And SelItem <> i Then
        MsgBox "Sorry but this column name already exists", vbCritical, "Error: Column exists"
    End If
Next i
If TabProp(SelItem, 0) = PKey And TabProp(SelItem, 0) <> "" Then
    PKey = txtName
End If
TabProp(SelItem, 0) = txtName.Text
lvwColumns.ListItems(SelItem).Text = txtName.Text
End Sub

Private Sub MoveItemUp()                                    'Moves a column header up in the array
Dim Key As Integer                                          'The tables primary key
Dim TTP(7) As String                                        'Temporary table property array
Dim TCIO(1) As String                                       'Temporary column information first
Dim TCIT(1) As String                                       'Temporary column information second
Dim NimgLst As ImageList

Key = 0

Properties True                                             'Resets all the properties

If lvwColumns.ListItems(SelItem).SmallIcon = 1 Then         'Checks to see if either of them are keys
    Key = 1
End If
If lvwColumns.ListItems(SelItem - 1).SmallIcon = 1 Then
    Key = 2
End If

TTP(0) = TabProp(SelItem, 0)                    'Sets all temporary column info
TTP(1) = TabProp(SelItem, 1)
TTP(2) = TabProp(SelItem, 2)
TTP(3) = TabProp(SelItem, 3)
TTP(4) = TabProp(SelItem, 4)
TTP(5) = TabProp(SelItem, 5)
TTP(6) = TabProp(SelItem, 6)
TTP(7) = TabProp(SelItem, 7)

TabProp(SelItem, 0) = TabProp((SelItem - 1), 0)         'Moves selected one up one
TabProp(SelItem, 1) = TabProp((SelItem - 1), 1)
TabProp(SelItem, 2) = TabProp((SelItem - 1), 2)
TabProp(SelItem, 3) = TabProp((SelItem - 1), 3)
TabProp(SelItem, 4) = TabProp((SelItem - 1), 4)
TabProp(SelItem, 5) = TabProp((SelItem - 1), 5)
TabProp(SelItem, 6) = TabProp((SelItem - 1), 6)
TabProp(SelItem, 7) = TabProp((SelItem - 1), 7)

TCIO(0) = lvwColumns.ListItems(SelItem - 1).Text                    'Puts the column data to be swapped in temporary arrays
TCIO(1) = lvwColumns.ListItems(SelItem - 1).ListSubItems(1).Text

TCIT(0) = lvwColumns.ListItems(SelItem).Text
TCIT(1) = lvwColumns.ListItems(SelItem).ListSubItems(1).Text

SelItem = SelItem - 1

TabProp(SelItem, 0) = TTP(0)            'Puts temp into lower space in table properties
TabProp(SelItem, 1) = TTP(1)
TabProp(SelItem, 2) = TTP(2)
TabProp(SelItem, 3) = TTP(3)
TabProp(SelItem, 4) = TTP(4)
TabProp(SelItem, 5) = TTP(5)
TabProp(SelItem, 6) = TTP(6)
TabProp(SelItem, 7) = TTP(7)

lvwColumns.ListItems(SelItem + 1).Text = TCIO(0)                    'Swaps the temporary column data around
lvwColumns.ListItems(SelItem + 1).ListSubItems(1).Text = TCIO(1)

lvwColumns.ListItems(SelItem).Text = TCIT(0)
lvwColumns.ListItems(SelItem).ListSubItems(1).Text = TCIT(1)

If Key = 1 Then                         'If either of them are Key Feils then it will move them
    Set lvwColumns.SmallIcons = NimgLst
    lvwColumns.SmallIcons = imgTable
    lvwColumns.ListItems(SelItem).SmallIcon = 1
End If
If Key = 2 Then
    Set lvwColumns.SmallIcons = NimgLst
    lvwColumns.SmallIcons = imgTable
    lvwColumns.ListItems(SelItem + 1).SmallIcon = 1
End If

SetAllItems
Set lvwColumns.SelectedItem = lvwColumns.ListItems(SelItem)
End Sub

Private Sub MoveItemDown()
Dim Key As Integer
Dim TTP(7) As String
Dim TCIO(1) As String
Dim TCIT(1) As String
Dim NimgLst As ImageList        'A blank image list

Key = 0

If lvwColumns.ListItems(SelItem).SmallIcon = 1 Then     'Checks to see if either of them are the Keys
    Key = 1
End If
If lvwColumns.ListItems(SelItem + 1).SmallIcon = 1 Then
    Key = 2
End If

TTP(0) = TabProp(SelItem, 0)            'Puts top item in temporary array
TTP(1) = TabProp(SelItem, 1)
TTP(2) = TabProp(SelItem, 2)
TTP(3) = TabProp(SelItem, 3)
TTP(4) = TabProp(SelItem, 4)
TTP(5) = TabProp(SelItem, 5)
TTP(6) = TabProp(SelItem, 6)
TTP(7) = TabProp(SelItem, 7)

TabProp(SelItem, 0) = TabProp((SelItem + 1), 0)         'Moves selected item down one
TabProp(SelItem, 1) = TabProp((SelItem + 1), 1)
TabProp(SelItem, 2) = TabProp((SelItem + 1), 2)
TabProp(SelItem, 3) = TabProp((SelItem + 1), 3)
TabProp(SelItem, 4) = TabProp((SelItem + 1), 4)
TabProp(SelItem, 5) = TabProp((SelItem + 1), 5)
TabProp(SelItem, 6) = TabProp((SelItem + 1), 6)
TabProp(SelItem, 7) = TabProp((SelItem + 1), 7)

TCIO(0) = lvwColumns.ListItems(SelItem).Text
TCIO(1) = lvwColumns.ListItems(SelItem).ListSubItems(1).Text

TCIT(0) = lvwColumns.ListItems(SelItem + 1).Text
TCIT(1) = lvwColumns.ListItems(SelItem + 1).ListSubItems(1).Text

SelItem = SelItem + 1

TabProp(SelItem, 0) = TTP(0)                'Moves original up one
TabProp(SelItem, 1) = TTP(1)
TabProp(SelItem, 2) = TTP(2)
TabProp(SelItem, 3) = TTP(3)
TabProp(SelItem, 4) = TTP(4)
TabProp(SelItem, 5) = TTP(5)
TabProp(SelItem, 6) = TTP(6)
TabProp(SelItem, 7) = TTP(7)

lvwColumns.ListItems(SelItem).Text = TCIO(0)                    'Puts the column data in temporary arrays
lvwColumns.ListItems(SelItem).ListSubItems(1).Text = TCIO(1)

lvwColumns.ListItems(SelItem - 1).Text = TCIT(0)
lvwColumns.ListItems(SelItem - 1).ListSubItems(1).Text = TCIT(1)

If Key = 1 Then
    Set lvwColumns.SmallIcons = NimgLst     'Set the icons to a blank image list
    lvwColumns.SmallIcons = imgTable        'Resets the icons to image list
    lvwColumns.ListItems(SelItem).SmallIcon = 1      'Changes small icon to the key
End If
If Key = 2 Then
    Set lvwColumns.SmallIcons = NimgLst     'Set the icons to a blank image list
    lvwColumns.SmallIcons = imgTable        'Resets the icons to image list
    lvwColumns.ListItems(SelItem - 1).SmallIcon = 1    'Changes small icon to the key
End If

SetAllItems
Set lvwColumns.SelectedItem = lvwColumns.ListItems(SelItem)
End Sub

Private Sub Properties(Enable As Boolean)
If Enable = False Then                      'If there is no columns then this is false
    txtName.Enabled = False
    cboType.Enabled = False
    optRYes.Enabled = False
    optRNo.Enabled = False
    txtMax.Enabled = False
    txtDefault.Enabled = False
    cboDefault.Enabled = False
    cmdValidation.Enabled = False
    cboLType.Enabled = False
    cmdSQL.Enabled = False
    mnuFileCreate.Enabled = False
Else                                        'If there is columns then it is set so that everything can be used
    txtName.Enabled = True
    cboType.Enabled = True
    optRYes.Enabled = True
    optRNo.Enabled = True
    txtMax.Enabled = True
    txtDefault.Enabled = True
    cboDefault.Enabled = True
    cmdValidation.Enabled = True
    cboLType.Enabled = True
    cmdSQL.Enabled = False
    mnuFileCreate.Enabled = True
End If
End Sub

Private Sub txtNotes_Change()
TNotes = txtNotes.Text
End Sub

Private Sub txtRSource_Change()
TabProp(SelItem, 4) = txtRSource.Text
End Sub

Private Sub CreateTable()
Dim Whole As String                     'The process of actually making the table in text format
Dim i As Integer

Dim Key As String

Dim Name As String
Dim Required As Integer
Dim FType As String
Dim ColumnType As Integer
Dim SQL As String
Dim MaxLen As Integer
Dim DefaultValue As String
Dim ValidationRules As String

Dim Splitter As String

Splitter = Chr(222) & Chr(232) & Chr(222)               'The delimiter of the table

For i = 1 To TotalColumns                               'Finds which column is the key
    If lvwColumns.ListItems(i).SmallIcon = 1 Then
        Key = lvwColumns.ListItems(i).Text
        Exit For
    End If
Next i

If Key = "" Then
    If MsgBox("This table does not contian a Key feild, do you wish to proced without one", vbYesNo) = vbNo Then
        Exit Sub
    End If
End If
    

Whole = TableName & Splitter & "1" & Splitter & Key & Splitter & "0" & vbTab & TNotes & vbTab & Format(Date, "dd/MM/yy") & vbTab & Format(Date, "dd/MM/yy") & Splitter

For i = 1 To TotalColumns
    Name = TabProp(i, 0)            'Column Name
    Required = TabProp(i, 1)        'Required
    FType = TabProp(i, 2)           'Field Type
    ColumnType = TabProp(i, 3)      'Column Type
    SQL = TabProp(i, 4)             'Record Source
    ValidationRules = TabProp(i, 5) 'Validation Rules
    DefaultValue = TabProp(i, 6)    'Default Value
    MaxLen = TabProp(i, 7)          'Max Len
    If ColumnType = 2 Then
        ColumnType = 1
    End If
    If i = 1 Then
        Whole = Whole & Name & vbTab & Required & vbTab & FType & vbTab & ColumnType & vbTab _
        & SQL & vbTab & ValidationRules & vbTab & DefaultValue & vbTab & MaxLen
    Else
        Whole = Whole & vbTab & Name & vbTab & Required & vbTab & FType & vbTab & ColumnType & vbTab _
        & SQL & vbTab & ValidationRules & vbTab & DefaultValue & vbTab & MaxLen
    End If
Next i
AddTable Whole
frmATable.Hide
End Sub

Private Sub AddTable(TTable As String)
Dim TTables() As String
Dim i As Integer
Dim j As Integer

Dim TempTable() As String

ReDim TTables(TotalTables, 4)

If CFile = "" Then
    If MsgBox("There is no database file selected, do you wish to create a new one?", vbYesNo, "Create a new database?") = vbYes Then
        cdlFile.ShowSave
        If cdlFile.FileName <> "" Then
            CFile = cdlFile.FileName
        End If
    End If
End If


TempTable = Split(TTable, Chr(222) & Chr(232) & Chr(222))

For i = 1 To TotalTables
    For j = 0 To 4
        TTables(i, j) = Tables(i, j)
    Next j
Next i

TotalTables = TotalTables + 1

ReDim Tables(TotalTables, 4)

For i = 1 To TotalTables - 1
    For j = 0 To 4
        Tables(i, j) = TTables(i, j)
    Next j
Next i

Tables(TotalTables, 0) = TempTable(0)
Tables(TotalTables, 1) = TempTable(1)
Tables(TotalTables, 2) = TempTable(2)
Tables(TotalTables, 3) = TempTable(3)
Tables(TotalTables, 4) = TempTable(4)

SaveToFile (CFile)
frmMain.LoadTablesToList
End Sub

Private Sub txtTName_Change()
TableName = txtTName.Text
End Sub

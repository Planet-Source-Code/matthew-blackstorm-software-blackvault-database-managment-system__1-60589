VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRelationships 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add or Edit Relationships"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddRel 
      Caption         =   "Add Relationship"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox cboMany 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.ComboBox cboOne 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin MSComctlLib.ListView lvwRel 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Key Table"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Key Feild"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Many Table"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblMany 
      Caption         =   "M"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   390
      Width           =   375
   End
   Begin VB.Label lblOne 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   390
      Width           =   375
   End
   Begin VB.Line linOtM 
      X1              =   2880
      X2              =   3240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblNonKey 
      Caption         =   "Table With Non Key Feild:"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblKey 
      Caption         =   "Table With Key Feild:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu mnuDel 
      Caption         =   "mnuDel"
      Visible         =   0   'False
      Begin VB.Menu mnuDelRel 
         Caption         =   "Delete Relationship"
      End
   End
End
Attribute VB_Name = "frmRelationships"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim KeyFeild As String
Dim SelItem As Integer

Private Sub cboMany_Click()
If TableExists(False) = True Then               'If the table exists then the relationship can be added
    cmdAddRel.Enabled = True
Else
    cmdAddRel.Enabled = False
End If
End Sub

Private Sub cboOne_Click()
Dim i As Integer
Dim j As Integer
If cboOne.Text = "" Then                                                        'If there is no text in the feild then it will make it all disabled
    cboMany.Clear
    cboMany.Enabled = False
    cmdAddRel.Enabled = False
    Exit Sub
ElseIf TableExists(True) = True Then
    cboMany.Clear
    cmdAddRel.Enabled = False
    For i = 1 To TotalTables                                                    'Loops through to find the Key table
        If Tables(i, 0) = cboOne.Text Then                                      'When found it starts the look for the possible Non-Key tables
            KeyFeild = Tables(i, 1)
            For j = 1 To TotalTables
                If Tables(i, 0) <> Tables(j, 0) Then
                    If KeyExistInMany(Tables(j, 0), Tables(i, 1)) = True Then   'When found they are added to the posibilites
                        cboMany.AddItem Tables(j, 0)
                    End If
                End If
            Next j
        End If
    Next i
    If cboMany.ListCount > 0 Then                   'If there is any Non-Key tables found then it is enabled
        cboMany.Enabled = True
    Else
        cboMany.Enabled = False
    End If
Else
    cboMany.Clear                                   'If there is no feilds found then its all disalbed
    cboMany.Enabled = False
    cmdAddRel.Enabled = False
End If
End Sub

Private Sub cmdAddRel_Click()
Dim i As Integer
If TableExists(True) = True And TableExists(False) = True Then              'Check that both the Key and Non-Key tables exist
    For i = 1 To lvwRel.ListItems.Count
        If lvwRel.ListItems(i).Text = cboOne.Text Then
            MsgBox "Relationship already exists for the '" & cboOne.Text & "' table", vbCritical, "Error: Relationship Already Exists"
            Exit Sub
        End If
    Next i
    Set LItem = lvwRel.ListItems.Add(, , cboOne.Text)                       'Adds the new relationship
    LItem.ListSubItems.Add , , KeyFeild
    LItem.ListSubItems.Add , , cboMany.Text
    SaveAllRelationships                                                    'Saves the new relationship
End If
End Sub

Private Sub cmdCancel_Click()
frmRelationships.Hide
End Sub

Private Sub cmdOk_Click()
frmRelationships.Hide
End Sub

Private Sub lvwRel_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub lvwRel_ItemClick(ByVal Item As MSComctlLib.ListItem)
SelItem = Item.Index
End Sub

Private Sub lvwRel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Or KeyCode = 46 And lvwRel.ListItems.Count > 0 Then
    If MsgBox("Do you wish to delete the relationship between the two tables;" & vbNewLine & lvwRel.ListItems(SelItem).Text & " and " & lvwRel.ListItems(SelItem).ListSubItems(2), vbYesNo + vbInformation, "Delete This Relationship?") = vbYes Then
        lvwRel.ListItems.Remove (SelItem)
        SaveAllRelationships
    End If
End If
End Sub

Private Function KeyExistInMany(TableName As String, FeildName As String) As Boolean
Dim TTable() As String
Dim TColumns() As String
Dim Columns As Integer
Dim i As Integer
Dim j As Integer
KeyExistInMany = False
For i = 1 To TotalTables                                                'Loops through to find possible Non-Key table
    If TableName = Tables(i, 0) Then
        TTable = Split(Tables(i, 4), Chr(212) & Chr(232) & Chr(212))    'Splits the table info up
        Columns = CountSubStrings(TTable(0), vbTab)                     'Counts the number of column headers
        TColumns = Split(TTable(0), vbTab)                              'Splits the columns headers up
        For j = 0 To Columns Step 8                                     'Loops through to find the if the Key column exists in table
            If TColumns(j) = FeildName Then
                KeyExistInMany = True
                Exit Function
            End If
        Next j
    End If
Next i
End Function

Public Function LoadAllRelationships()
Dim i As Integer
Dim CRel() As String
Dim TotalRel As Integer
Dim TRel() As String
cboOne.Clear

For i = 1 To TotalTables                            'Loads all current Key tables to combo box
    If Tables(i, 1) <> "" Then
        cboOne.AddItem Tables(i, 0)
    End If
Next i

If TotalTables = 0 Then
    cboMany.Enabled = False
End If

CRel = Split(DBRelationships, Chr(212) & Chr(232) & Chr(212))               'Splits the total relationships text up into each relationship
TotalRel = CountSubStrings(DBRelationships, Chr(212) & Chr(232) & Chr(212)) 'Find the total number of realationships

For i = 0 To TotalRel                               'Loads all of the relationships to the listview
    TRel = Split(CRel(i), vbTab)
    Set LItem = lvwRel.ListItems.Add(, , TRel(0))
    LItem.ListSubItems.Add , , TRel(1)
    LItem.ListSubItems.Add , , TRel(2)
Next i

frmRelationships.Show vbModal
End Function

Private Sub SaveAllRelationships()
Dim i As Integer
Dim Total As String
For i = 1 To lvwRel.ListItems.Count
    If i = 1 Then
        Total = lvwRel.ListItems(i).Text & vbTab & lvwRel.ListItems(i).ListSubItems(1) & vbTab & lvwRel.ListItems(i).ListSubItems(2)
    Else
        Total = Total & Chr(212) & Chr(232) & Chr(212) & lvwRel.ListItems(i).Text & vbTab & lvwRel.ListItems(i).ListSubItems(1) & vbTab & lvwRel.ListItems(i).ListSubItems(2)
    End If
Next i
DBRelationships = Total
If SaveToFile(CFile) = False Then
    MsgBox "There was an error updating the database with the new relationship", vbOKOnly
End If
End Sub

Private Function TableExists(OneFeild As Boolean) As Boolean        'Used to check if the table entered exists
Dim i As Integer
TableExists = False
If OneFeild = True Then
    For i = 0 To cboOne.ListCount
        If cboOne.Text = cboOne.List(i) Then
            TableExists = True
            Exit Function
        End If
    Next i
Else
    For i = 0 To cboMany.ListCount
        If cboMany.Text = cboMany.List(i) Then
            TableExists = True
            Exit Function
        End If
    Next i
End If
End Function


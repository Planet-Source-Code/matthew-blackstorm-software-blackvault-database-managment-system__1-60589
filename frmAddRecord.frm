VERSION 5.00
Begin VB.Form frmAddRecord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Record"
   ClientHeight    =   990
   ClientLeft      =   4500
   ClientTop       =   5070
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox cboLItem 
      Height          =   315
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtLItem 
      Height          =   315
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblName 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAddRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim YPos As Integer                                 'Used for tracking the vertical position of items

Dim Columns() As String                             'Column headers array

Dim LIUpto As Integer                               'Tracks the total number of list items loaded

Public Sub LoadColumns(ColumnInfo As String)        'Loads the text and combo boxes and lables
Dim SCInfo() As String                              'Array that contains the split column headers
Dim CInfoI As Integer                               'Total number of column header items
Dim CDoing As Integer                               'Column header position currently being done
Dim TotalColumns As Integer                         'Total number of columns being loaded

Dim i As Integer


YPos = 0                                            'Resets the YPos variable to top
LIUpto = 0                                          'Resets the list item upto

CInfoI = CountSubStrings(ColumnInfo, vbTab)         'Counts the number of column headers
SCInfo = Split(ColumnInfo, vbTab)                   'Splits the column header info

TotalColumns = (CInfoI + 1) / 8                     'Works out the total column headers

ReDim Columns(TotalColumns, 7) As String            'Redimensions the column headers array

CDoing = 1

For i = 0 To CInfoI Step 8                          'Loads each of the 8 components of a column into the array
    Columns(CDoing, 0) = SCInfo(i)                  'Done in a step of 8 to speed up the loop
    Columns(CDoing, 1) = SCInfo(i + 1)
    Columns(CDoing, 2) = SCInfo(i + 2)
    Columns(CDoing, 3) = SCInfo(i + 3)
    Columns(CDoing, 4) = SCInfo(i + 4)
    Columns(CDoing, 5) = SCInfo(i + 5)
    Columns(CDoing, 6) = SCInfo(i + 6)
    Columns(CDoing, 7) = SCInfo(i + 7)
    CreateRow (CDoing)
    CDoing = CDoing + 1
Next i

End Sub

Private Sub CreateRow(Index As Integer)             'Creates a row of the specified type
LIUpto = LIUpto + 1
Load lblName(LIUpto)                                'Sets the properties of the new lable
lblName(LIUpto).Caption = Columns(Index, 0)
lblName(LIUpto).Left = 120
lblName(LIUpto).Top = 120 + ((LIUpto - 1) * 360)
lblName(LIUpto).Visible = True

Select Case Columns(Index, 2)                       'Used for working out what type of feild is to be loaded
    Case 0
        If Columns(Index, 4) = "<none>" Then        'Checks for any SQL statment
            CreateText (Index)
            txtLItem(Index).TabIndex = Index
        Else
            CreateCombo (Index)
            cboLItem(Index).TabIndex = Index
        End If
    Case 1
        CreateCombo Index, True                     'Creates a Bit field
        cboLItem(Index).TabIndex = Index
    Case 11
        CreateText Index, True                      'Creates A password Feild
        txtLItem(Index).TabIndex = Index
    Case Else
        CreateText (Index)
        txtLItem(Index).TabIndex = Index
End Select

cmdAdd.Top = 500 + ((LIUpto - 1) * 360)             'Positions the buttons
cmdCancel.Top = 500 + ((LIUpto - 1) * 360)
frmAddRecord.Height = 1380 + ((LIUpto - 1) * 360)
End Sub

Private Sub CreateText(Index As Integer, Optional IsPword As Boolean)       'Used to create a text box
Load txtLItem(LIUpto)
txtLItem(LIUpto).Left = 2400
txtLItem(LIUpto).Top = 120 + ((LIUpto - 1) * 360)
txtLItem(LIUpto).MaxLength = Columns(Index, 7)
txtLItem(LIUpto).Visible = True
If IsPword = True Then                                                      'If it is of a password type it sets the password character
    txtLItem(LIUpto).PasswordChar = "*"
End If
End Sub

Private Sub CreateCombo(Index As Integer, Optional IsBit As Boolean)        'Used to create a combo box
Load cboLItem(LIUpto)                                                       'Loads the new combo box and sets its properties
cboLItem(LIUpto).Left = 2400
cboLItem(LIUpto).Top = 120 + ((LIUpto - 1) * 360)
cboLItem(LIUpto).Visible = True
If IsBit = True Then                                                        'If its a bit feild it adds the T or F options
    SetList "True::-::False", Index
    Exit Sub
End If
If Left(Columns(Index, 4), 4) = "LST:" Then
    SetList Columns(Index, 4), LIUpto
Else
    'Insert Code Here To Preform SQL query
End If
End Sub

Private Sub SetList(TList As String, Index As Integer)                      'Adds the listitems to the combo box
Dim Whole As String
Dim TSLst() As String
Dim TLItems As Integer
Dim i As Integer
Whole = Right(TList, Len(TList) - 4)                                        'Splits up the list items
TLItems = CountSubStrings(Whole, "::-::")
TSLst = Split(Whole, "::-::")
For i = 0 To TLItems                                                        'Loops through adding the items
    cboLItem(Index).AddItem TSLst(i)
Next i
End Sub

Private Sub cmdAdd_Click()
Dim i As Integer
Dim Item As String
Dim PSplit() As String                              'For editing the table properties
Dim TNEditing As Integer                            'The index of the table being edited in the Tables array
Dim Records As Integer                              'Total number of records
Dim Key As String                                   'Name of the key feild
Dim KeyCol As Integer                               'Index of the key feild
Dim SelText As String                               'Used for checking for duplicates
Dim NewText As String
Dim Temp As String
For i = 1 To TotalTables                            'Find the table to edit
    If frmDBView.TableName = Tables(i, 0) Then
        TNEditing = i
        Key = Tables(i, 2)                          'Sets the key of the table
        Exit For
    End If
Next i

For i = 1 To LIUpto                                             'Validates all of the feilds
    If Columns(i, 2) = 1 Or Columns(i, 4) <> "<none>" Then
        Temp = cboLItem(i).Text
    Else
        Temp = txtLItem(i).Text
    End If
    If Validate(Int(Columns(i, 2)), Columns(i, 5), Temp) = False Then
        MsgBox "Record cannot be added as the feild '" & Columns(i, 0) & "' does not fit the validation rules", vbCritical, "Error: Validation Rules Not Met"
        Exit Sub
    End If
Next i

If Key <> "" Then                                   'Routine to check if the table has a
    For i = 1 To LIUpto                             'Key feild, and if so it invokes
        If Key = Columns(i, 0) Then                 'loop to check for duplicates
            KeyCol = i                              'Sets the column number for the key column
            Exit For
        End If
    Next i
                                                    'Find the text in the key field
    If CInt(Columns(KeyCol, 2)) = 1 Or Columns(KeyCol, 4) <> "<none>" Then
        NewText = cboLItem(KeyCol).Text
    Else
        NewText = txtLItem(KeyCol).Text
    End If

    For i = 1 To frmDBView.lvwDB.ListItems.Count     'Loops through the list looking for duplicates
        If KeyCol = 1 Then
            SelText = frmDBView.lvwDB.ListItems(i).Text
        Else                                                'If the key is not in the first column then uses this
            Set LItem = frmDBView.lvwDB.ListItems(i)
            SelText = LItem.ListSubItems(KeyCol - 1).Text
        End If
        If NewText = SelText Then
            MsgBox "This table cannot contain duplicate records containing the same '" & Key & "'", vbCritical, "Error: Duplicates Not Allowed"
            Exit Sub
        End If
    Next i
End If

                                                    'Finds the first item to add to list
If CInt(Columns(1, 2)) = 1 Or Columns(1, 4) <> "<none>" Then
    Item = cboLItem(1).Text
Else
    If Int(Columns(1, 2)) = 11 Then
        Item = CreatePWDHash(txtLItem(1).Text)
    Else
        Item = txtLItem(1).Text
    End If
End If
    
Set LItem = frmDBView.lvwDB.ListItems.Add(, , Item)

For i = 2 To LIUpto                                 'Loops through adding data to all columns
    If Columns(i, 2) = 1 Or Columns(i, 4) <> "<none>" Then
        Item = cboLItem(i).Text
    Else
        If Int(Columns(i, 2)) = 11 Then
            Item = CreatePWDHash(txtLItem(i).Text)
        Else
            Item = txtLItem(i).Text
        End If
    End If
    LItem.ListSubItems.Add , , Item
Next i

                                                    'Edits the table properties updating the
Tables(TNEditing, 4) = frmDBView.SaveTable

PSplit = Split(Tables(TNEditing, 3), vbTab)
Records = CInt(PSplit(0)) + 1
PSplit(3) = Format(Date, "dd/MM/yy")
Tables(TNEditing, 3) = Records & vbTab & PSplit(1) & vbTab & PSplit(2) & vbTab & PSplit(3)


If SaveToFile(CFile) = False Then
    MsgBox "There was an error updating the database with the new record", vbOKOnly
Else
    frmAddRecord.Hide
    UnloadAll
    frmMain.LoadTablesToList
End If
End Sub

Private Sub cmdCancel_Click()
frmAddRecord.Hide
UnloadAll
End Sub

Private Sub Form_Unload(Cancel As Integer)
UnloadAll
End Sub

Private Sub UnloadAll()                             'Unloads all of the text and combo boxes and lables
Dim i As Integer
For i = 1 To LIUpto                                 'Loops through all the items
    Unload lblName(i)
    Select Case Columns(i, 2)
    Case 0
        If Columns(i, 4) = "<none>" Then
            Unload txtLItem(i)
        Else
            Unload cboLItem(i)
        End If
    Case 1
        Unload txtLItem(i)
    Case 11
        Unload txtLItem(i)
    Case Else
        Unload txtLItem(i)
    End Select
Next i
ReDim Columns(1)                                    'Redimensions the column header array
LIUpto = 0                                          'Resets the list item count

End Sub

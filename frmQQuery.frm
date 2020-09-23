VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmQQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quick Query"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwResults 
      Height          =   1815
      Left            =   0
      TabIndex        =   4
      Top             =   3720
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   3201
      View            =   3
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
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtSQL 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   9255
   End
   Begin VB.Label lblResults 
      Caption         =   "Results are displayed below:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   5415
   End
   Begin VB.Label lblSQLInfo 
      Caption         =   "Please enter the SQL you wish to execute:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmQQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TotalColumns As Integer

Private Sub cmdExecute_Click()
Dim SQLEng As clsSQL
Dim SQL As String
Dim ERROR As String
Dim WasERROR As Boolean
Dim Temp As String
WasERROR = False
If txtSQL.Text <> "" Then
    SQL = Trim(txtSQL.Text)
    Set SQLEng = New clsSQL
    Temp = SQLEng.PraseSQL(SQL, WasERROR, ERROR)
    If WasERROR = True Then
        MsgBox ERROR
    Else
        LoadResults Temp
    End If
    Set SQLEng = Nothing
End If
End Sub

Private Function UnloadResultsTable()
TotalColumns = 0

lvwResults.ListItems.Clear
lvwResults.ColumnHeaders.Clear
End Function

Public Function CreateColumn(CName As String)
TotalColumns = TotalColumns + 1         'Add New Column to List
lvwResults.ColumnHeaders.Add , , CName
End Function

Private Function LoadResults(Table As String)
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
Dim Doing As Integer
Dim ColumnUpTo As Integer

Dim Delimiter As Integer

UnloadResultsTable

Delimiter = FindPosition(1, Chr(212) & Chr(232) & Chr(212), Table)
CTSplit = Split(Table, Chr(212) & Chr(232) & Chr(212))
CInfo = CTSplit(0)
If Delimiter < Len(Table) And Delimiter - 4 > 0 Then
    TInfo = CTSplit(1)
End If

CInfoI = CountSubStrings(CInfo, vbTab)
SCInfo = Split(CInfo, vbTab)
TInfoI = CountSubStrings(TInfo, vbTab)
STInfo = Split(TInfo, vbTab)
Doing = 1
For i = 0 To CInfoI Step 8
    CName = SCInfo(i)
    CreateColumn CName
Next i

Doing = 1
If Len(Trim(TInfo)) = 0 Then
    Exit Function
End If

For i = 0 To TInfoI
    If Doing = 1 Then
        Set LItem = lvwResults.ListItems.Add(, , STInfo(i))
    End If
    If Doing <> 1 Then
        LItem.ListSubItems.Add , , STInfo(i)
    End If
    If Doing = lvwResults.ColumnHeaders.Count Then
        Doing = 1
    Else
        Doing = Doing + 1
    End If
Next i

End Function

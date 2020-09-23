VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table Properties - "
   ClientHeight    =   5805
   ClientLeft      =   6720
   ClientTop       =   4365
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.Frame famTable 
      Height          =   5295
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      Begin VB.CheckBox chkReadOnly 
         Caption         =   "Read Only"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   4680
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Image Image1 
         Height          =   960
         Left            =   120
         Picture         =   "frmProperties.frx":0000
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblTDescrip 
         Caption         =   "Description:"
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4680
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label lblTSpace 
         Caption         =   "Total Size:"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label lblTName 
         Caption         =   "Table Name:"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lblTCreated 
         Caption         =   "Date Created:"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblTEdited 
         Caption         =   "Last Edited:"
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label lblTRecords 
         Caption         =   "Total Records:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   5400
      Width           =   1335
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PType As Integer                                'Stores the properties type ie. table or querry
Dim PPos As Integer                                 'Stores the position of the selected item in the array

Public Sub DisplayTableProp(TabPosition As Integer)
Dim PSplit() As String
Dim TSize As Long
PType = 1
PPos = TabPosition
frmProperties.Caption = "Table Properties - " & Tables(TabPosition, 0)

If Tables(TabPosition, 1) = 0 Then
    chkReadOnly.Value = vbChecked
Else
    chkReadOnly.Value = vbUnchecked
End If

PSplit = Split(Tables(TabPosition, 3), vbTab)
lblTName.Caption = "Table Name: " & Tables(TabPosition, 0)
lblTCreated.Caption = "Table Created: " & PSplit(2)
lblTEdited.Caption = "Table Edited: " & PSplit(3)
lblTDescrip.Caption = "Table Description:" & vbNewLine & PSplit(1)

lblTRecords.Caption = "Total Records: " & PSplit(0)
TSize = Len(Tables(TabPosition, 0)) + Len(Tables(TabPosition, 1)) + Len(Tables(TabPosition, 2)) + Len(Tables(TabPosition, 3)) + Len(Tables(TabPosition, 4))
lblTSpace.Caption = "Total Size: " & ModSize(TSize)

frmProperties.Show vbModal
End Sub


Private Function ModSize(Size As Long) As String
If Size <= 1024 Then
    ModSize = Size & " Bytes"
ElseIf Size > 1024 And Size < 1048576 Then
    ModSize = GetKBs(Size) & " KBytes"
ElseIf Size > 1048576 Then
    ModSize = GetMegs(Size) & " MBytes"
End If
End Function

Private Function GetKBs(Size As Long) As Double
Dim KBS As Double
KBS = Size / 1024
GetKBs = Round(KBS * 100) / 100
End Function

Private Function GetMegs(Size As Long) As Double
Dim MBS As Double
MBS = Size / 1048576
GetMegs = Round(MBS * 100) / 100
End Function

Private Sub chkReadOnly_Click()
Select Case PType
    Case 1
        If chkReadOnly.Value = vbChecked Then
            Tables(PPos, 1) = 0
        Else
            Tables(PPos, 1) = 1
        End If
        SaveToFile (CFile)
    Case 2
        
    Case 3
        
End Select

End Sub

Private Sub cmdOk_Click()
frmProperties.Hide
End Sub

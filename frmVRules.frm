VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVRules 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Validation Rules"
   ClientHeight    =   6390
   ClientLeft      =   4140
   ClientTop       =   2790
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame famARules 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   4850
      Begin VB.CommandButton cmdAdd 
         Cancel          =   -1  'True
         Caption         =   "Add"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtValue 
         Height          =   285
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   1815
      End
      Begin VB.OptionButton optS 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   3855
      End
      Begin VB.OptionButton optS 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   3855
      End
      Begin VB.OptionButton optS 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   3855
      End
      Begin VB.OptionButton optS 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   3855
      End
      Begin VB.Label lblInfo 
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   6000
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvwValidation 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   5953
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Constraints"
         Object.Width           =   6174
      EndProperty
   End
End
Attribute VB_Name = "frmVRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelItem As Integer

Dim SelectedType As Integer

Private Sub cmdCancel_Click()
frmVRules.Hide
End Sub

Private Sub cmdOk_Click()
Dim i As Integer
Dim Whole As String
Dim VType As String
Dim VRule As String
For i = 1 To lvwValidation.ListItems.Count
    VType = lvwValidation.ListItems(i).Text
    VRule = lvwValidation.ListItems(i).ListSubItems(1).Text
    Whole = Whole & FindValidation(VType, VRule)
Next i
If lvwValidation.ListItems.Count = 0 Then
    Whole = "<none>"
End If
frmATable.SetVRules Trim(Whole)
frmVRules.Hide
End Sub

Private Function FindValidation(VType As String, VRule As String) As String


VRule = Replace(VRule, " ", "<<space>>")
VRule = Replace(VRule, ";", "<<semi>>")
VRule = Replace(VRule, vbTab, "<<tab>>")

Select Case SelectedType
    Case 0
        Select Case VType
            Case "Contains"
                FindValidation = "CT[" & VRule & "]; "
            Case "Does Not Contain"
                FindValidation = "DC[" & VRule & "]; "
            Case "Prefix Of"
                FindValidation = "PF[" & VRule & "]; "
            Case "Suffix Of"
                FindValidation = "SF[" & VRule & "]; "
        End Select
    Case 2, 3, 4, 5
        Select Case VType
            Case "Less Than"
                FindValidation = "LS[" & VRule & "]; "
            Case "Greater Than"
                FindValidation = "GR[" & VRule & "]; "
            Case "Equal To"
                FindValidation = "EQ[" & VRule & "]; "
            Case "Not Equal To"
                FindValidation = "NE[" & VRule & "]; "
        End Select
    Case 6, 7, 8
        Select Case VType
            Case "Before"
                FindValidation = "BF[" & VRule & "]; "
            Case "After"
                FindValidation = "AF[" & VRule & "]; "
            Case "Equal To"
                FindValidation = "EQ[" & VRule & "];"
            Case "Not Equal To"
                FindValidation = "NE[" & VRule & "];"
        End Select
    Case 9
        
    Case 10
        Select Case VType
            Case "Of Domain Type"
                FindValidation = "OT[" & VRule & "]; "
            Case "Not Of Domain Type"
                FindValidation = "NT[" & VRule & "]; "
            Case "Of Domain Name"
                FindValidation = "OD[" & VRule & "]; "
            Case "Not Of Domain Name"
                FindValidation = "ND[" & VRule & "]; "
        End Select
    Case 11
        Select Case VType
            Case "Contains"
                FindValidation = "CT[" & VRule & "]; "
            Case "Does Not Contain"
                FindValidation = "DC[" & VRule & "]; "
            Case "Min Length"
                FindValidation = "MN[" & VRule & "]; "
            Case "Max Length"
                FindValidation = "MX[" & VRule & "]; "
        End Select
End Select
End Function

Private Sub LoadAllRules(Rules As String)
Dim VRules() As String
Dim NVRules As Integer
Dim i As Integer
NVRules = CountSubStrings(Rules, ";") - 1
VRules = Split(Rules, ";")
For i = 0 To NVRules
    FindRules (Trim(VRules(i)))
Next i
End Sub


Private Function FindRules(Validation As String)
Dim VRName As String
Dim VRule As String

Select Case Left(Validation, 2)
    Case "CT"
        VRName = "Contains"
    Case "DC"
        VRName = "Does Not Contain"
    Case "PF"
        VRName = "Prefix Of"
    Case "SF"
        VRName = "Suffix Of"
    Case "LS"
        VRName = "Less Than"
    Case "GR"
        VRName = "Greater Than"
    Case "EQ"
        VRName = "Equal To"
    Case "NE"
        VRName = "Not Equal To"
    Case "BF"
        VRName = "Before"
    Case "AF"
        VRName = "After"
    Case "OT"
        VRName = "Of Domain Type"
    Case "NT"
        VRName = "Not Of Domain Type"
    Case "OD"
        VRName = "Of Domain Name"
    Case "ND"
        VRName = "Not Of Domain Name"
    Case "MN"
        VRName = "Min Length"
    Case "MX"
        VRName = "Max Length"
End Select

VRule = Mid(Validation, 4, Len(Validation) - 4)

VRule = Replace(VRule, "<<space>>", " ")
VRule = Replace(VRule, "<<semi>>", ";")
VRule = Replace(VRule, "<<tab>>", vbTab)

Set LItem = lvwValidation.ListItems.Add(, , VRName)
LItem.ListSubItems.Add , , VRule
End Function

Private Sub cmdAdd_Click()
Dim Temp As String

Temp = txtValue.Text
Temp = Replace(Temp, " ", "<<space>>")
Temp = Replace(Temp, ";", "<<semi>>")
Temp = Replace(Temp, vbTab, "<<tab>>")

Select Case SelectedType
    Case 0
        If optS(0).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Contains")
            LItem.ListSubItems.Add , , Temp
        End If
        If optS(1).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Does Not Contain")
            LItem.ListSubItems.Add , , Temp
        End If
        If optS(2).Value = True Then
            If ContainsText("]and[", LCase(Temp)) Or ContainsText("]not[", LCase(Temp)) Then
                MsgBox "Cannot use operators in Prefix or Suffix rules", vbCritical, "Error: Operators Cannot Be Used"
            Else
                Set LItem = lvwValidation.ListItems.Add(, , "Prefix Of")
                LItem.ListSubItems.Add , , Temp
            End If
        End If
        If optS(3).Value = True Then
            If ContainsText("]and[", LCase(Temp)) Or ContainsText("]not[", LCase(Temp)) Then
                MsgBox "Cannot use operators in Prefix or Suffix rules", vbCritical, "Error: Operators Cannot Be Used"
            Else
                Set LItem = lvwValidation.ListItems.Add(, , "Suffix Of")
                LItem.ListSubItems.Add , , Temp
            End If
        End If
    Case 2, 3, 4, 5
        If IsNumeric(Temp) Then
            If IsNumType(CDbl(Temp), SelectedType) = True Then
                If optS(0).Value = True Then
                    Set LItem = lvwValidation.ListItems.Add(, , "Less Than")
                    LItem.ListSubItems.Add , , Temp
                End If
                If optS(1).Value = True Then
                    Set LItem = lvwValidation.ListItems.Add(, , "Greater Than")
                    LItem.ListSubItems.Add , , Temp
                End If
                If optS(2).Value = True Then
                    Set LItem = lvwValidation.ListItems.Add(, , "Equal To")
                    LItem.ListSubItems.Add , , Temp
                End If
                If optS(3).Value = True Then
                    Set LItem = lvwValidation.ListItems.Add(, , "Not Equal To")
                    LItem.ListSubItems.Add , , Temp
                End If
            Else
                MsgBox "Please enter a number of the correct type", vbCritical, "Error: Incorrect Number Type"
            End If
        Else
            MsgBox "Please enter a number of the correct type", vbCritical, "Error: Non Numberic Input"
        End If
    Case 6, 7, 8
        If optS(0).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Before")
            LItem.ListSubItems.Add , , Temp
        End If
        If optS(1).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "After")
            LItem.ListSubItems.Add , , Temp
        End If
        If optS(2).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Equal To")
            LItem.ListSubItems.Add , , Temp
        End If
        If optS(3).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Not Equal To")
            LItem.ListSubItems.Add , , Temp
        End If
    Case 9
        
    Case 10
        If optS(0).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Of Domain Type")
            LItem.ListSubItems.Add , , Temp
        End If
        If optS(1).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Not Of Domain Type")
            LItem.ListSubItems.Add , , Temp
        End If
        If optS(2).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Of Domain Name")
            LItem.ListSubItems.Add , , Temp
        End If
        If optS(3).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Not Of Domain Name")
            LItem.ListSubItems.Add , , Temp
        End If
    Case 11
        If optS(0).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Contains")
            LItem.ListSubItems.Add , , Temp
        End If
        If optS(1).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Does Not Contain")
            LItem.ListSubItems.Add , , Temp
        End If
        If optS(2).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Min Length")
            LItem.ListSubItems.Add , , Temp
        End If
        If optS(3).Value = True Then
            Set LItem = lvwValidation.ListItems.Add(, , "Max Length")
            LItem.ListSubItems.Add , , Temp
        End If
End Select
txtValue.Text = ""
End Sub

Private Sub lvwValidation_BeforeLabelEdit(Cancel As Integer)
Cancel = 1
End Sub

Private Sub lvwValidation_ItemClick(ByVal Item As MSComctlLib.ListItem)
SelItem = Item.Index
End Sub

Private Sub lvwValidation_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 Or KeyCode = 46 Then
    If SelItem > 0 Then
        lvwValidation.ListItems.Remove (SelItem)
        If lvwValidation.ListItems.Count > SelItem Then
            Set lvwValidation.SelectedItem = lvwValidation.ListItems(SelItem)
        End If
        If lvwValidation.ListItems.Count < SelItem And lvwValidation.ListItems.Count > 0 Then
            SelItem = SelItem - 1
            Set lvwValidation.SelectedItem = lvwValidation.ListItems(SelItem)
        End If
        If lvwValidation.ListItems.Count = 0 Then
            SelItem = 0
        End If
    End If
End If
End Sub

Public Sub SetForType(SelType As Integer, VRules As String)
SelectedType = SelType
lvwValidation.ListItems.Clear
txtValue.Text = ""

Select Case SelType
    Case 0                                  'Text validation rules
        optS(0).Caption = "Contains"
        optS(0).Visible = True
        optS(1).Caption = "Does Not Contain"
        optS(1).Visible = True
        optS(2).Caption = "Prefix Of"
        optS(2).Visible = True
        optS(3).Caption = "Suffix Of"
        optS(3).Visible = True
        lblInfo.Caption = ""
    Case 2, 3, 4, 5                         'Byte, Integer, Long Integer and Double valadation rules
        optS(0).Caption = "Less Than"       'This code is reapeated because the rules are the same
        optS(0).Visible = True              'For each of the 4 different data types
        optS(1).Caption = "Greater Than"
        optS(1).Visible = True
        optS(2).Caption = "Equal To"
        optS(2).Visible = True
        optS(3).Caption = "Not Equal To"
        optS(3).Visible = True
        lblInfo.Caption = ""
    Case 6                                  'Time validation rules
        optS(0).Caption = "Before"
        optS(0).Visible = True
        optS(1).Caption = "After"
        optS(1).Visible = True
        optS(2).Caption = "Equal To"
        optS(2).Visible = True
        optS(3).Caption = "Not Equal To"
        optS(3).Visible = True
        lblInfo.Caption = "eg: 01:24, 23:41:05"
    Case 7
        optS(0).Caption = "Before"
        optS(0).Visible = True
        optS(1).Caption = "After"
        optS(1).Visible = True
        optS(2).Caption = "Equal To"
        optS(2).Visible = True
        optS(3).Caption = "Not Equal To"
        optS(3).Visible = True
        lblInfo.Caption = "eg: Monday, Tuesday"
    Case 8                                  'Date validation rules
        optS(0).Caption = "Before"
        optS(0).Visible = True
        optS(1).Caption = "After"
        optS(1).Visible = True
        optS(2).Caption = "Equal To"
        optS(2).Visible = True
        optS(3).Caption = "Not Equal To"
        optS(3).Visible = True
        lblInfo.Caption = "dd/mm/yyyy"
    Case 9
        
    Case 10                                 'URL validation rules
        optS(0).Caption = "Of Domain Type (eg .com, .net)"
        optS(0).Visible = True
        optS(1).Caption = "Not Of Domain Type (eg .com, .net)"
        optS(1).Visible = True
        optS(2).Caption = "Of Domain Name (eg. yourhost.com)"
        optS(2).Visible = True
        optS(3).Caption = "Not Of Domain Name (eg. yourhost.com)"
        optS(3).Visible = True
        lblInfo.Caption = ""
    Case 11
        optS(0).Caption = "Contains"
        optS(0).Visible = True
        optS(1).Caption = "Does Not Contain"
        optS(1).Visible = True
        optS(2).Caption = "Min Length"
        optS(2).Visible = True
        optS(3).Caption = "Max Length"
        optS(3).Visible = True
        lblInfo.Caption = ""
End Select
If VRules <> "<none>" Then
    LoadAllRules (VRules)
End If
frmVRules.Show vbModal
End Sub

Private Sub txtValue_Change()
If Len(txtValue.Text) = 0 Then
    cmdAdd.Enabled = False
Else
    cmdAdd.Enabled = True
End If
End Sub

VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3608
      TabIndex        =   1
      Top             =   9000
      Width           =   1935
   End
   Begin VB.TextBox txtBV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ToSay As String

Public Function DisplayMe()
ToAdd "Well thanks for downloading BlackVault, it is a program that I have been working on"
ToAdd "for a fair while now, originaly it was just going to be a poof of concept type"
ToAdd "program but soon became a bit more. The program once had an array of textboxes"
ToAdd "that were used to enter user data in but it did not take long to realise that this was"
ToAdd "memory intensive so I changed to listview which i had used many times before."
ToAdd "Currently the program is not even in beta stage, i needs a lot of work, the SQL class"
ToAdd "is nowhere near finished and there are many holes all over the program that need to"
ToAdd "be filled, but what i want is for feedback"
ToAdd ""
ToAdd ""
ToAdd "Well currently what you can do is create tables (not delete them, sorry), set the"
ToAdd "properties of each column in the table, just like a proper DBMS, run quick queries"
ToAdd "and export tables into XML format, there is also the functionality of having a MD5"
ToAdd "hashing class as when I set up the server and user system for the program, it will"
ToAdd "use this for added security, but this can also be used in any tables created"
ToAdd ""
ToAdd ""
ToAdd "What I plan to do in the future is finish all of the little gaps, create the server part,"
ToAdd "set up a HTML export function, get realational queries working (I never thought"
ToAdd "how complex it would be until I started thinking about it)"
txtBV.Text = ToSay
frmAbout.Show vbModal
End Function


Private Sub ToAdd(NewLine As String)
ToSay = ToSay & NewLine & vbCrLf
End Sub

Private Sub cmdOk_Click()
frmAbout.Hide
End Sub

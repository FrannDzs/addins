VERSION 5.00
Begin VB.Form frmNewGroup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Group"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3525
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   96
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Group name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmNewGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Function ShowEx(Optional InitText As String) As String
   Text1.Text = InitText
   If Len(InitText) Then
      Me.Caption = "Rename Group"
      Text1.SelStart = 0: Text1.SelLength = Len(InitText)
   Else
      Me.Caption = "New Group"
   End If
   Me.Show vbModal
   ShowEx = Text1.Text
   Unload Me
End Function
Private Sub cmdCancel_Click()
   Text1.Text = vbNullString
   Me.Hide
End Sub
Private Sub cmdOK_Click()
   If Len(Text1) = 0 Then MsgBox "Can't have a blank Group Name!", vbExclamation, "Come on now!"
   Me.Hide
End Sub


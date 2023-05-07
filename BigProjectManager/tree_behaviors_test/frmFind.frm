VERSION 5.00
Begin VB.Form frmFind 
   Caption         =   "Find"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5970
   LinkTopic       =   "Form2"
   ScaleHeight     =   6300
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin Project1.ucFilterList lv 
      Height          =   5535
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   9763
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_tv As TreeView

Sub init(tv As TreeView)

    On Error Resume Next
    Dim n As Node, fn As String, li As ListItem
    
    lv.Clear
    lv.SetIcons tv.ImageList
    
    For Each n In tv.Nodes
        If n.Image <> "folder" And n.Image <> "proj" Then
            Set li = lv.AddItem(n.Text)
            Set li.Tag = n
            li.SmallIcon = n.Image
        End If
    Next
    
    Me.Visible = True
    
End Sub

Private Sub Form_Load()
    lv.SetColumnHeaders "file*"
    lv.SetFont "Courier", 12
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lv.Width = Me.ScaleWidth - 200
    lv.Height = Me.ScaleHeight - 200
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Dim n As Node
    Set n = Item.Tag
    n.EnsureVisible
    n.selected = True
End Sub

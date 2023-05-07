VERSION 5.00
Begin VB.Form frmFind 
   Caption         =   "Find"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4230
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5715
   ScaleWidth      =   4230
   Begin BigProjectManager.ucFilterList lv 
      Height          =   5595
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   9869
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this could be another dockable user document.. its a toss up..

Dim selNode As Node

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -1

Sub SetWindowTopMost(f As Form)
   SetWindowPos f.hwnd, HWND_TOPMOST, f.Left / 15, _
        f.Top / 15, f.Width / 15, _
        f.Height / 15, Empty
End Sub


Sub init(tv As TreeView, Optional img As ImageList)

    On Error Resume Next
    Dim n As Node, fn As String, li As ListItem, c As CVBComponent
    
    lv.Clear
    lv.GridLines = False
    
    If Not img Is Nothing Then lv.SetIcons img
 
    For Each n In tv.Nodes
        If n.Image <> "folder" And n.Image <> "proj" Then
            Set li = lv.AddItem(n.Text)
            Set li.tag = n
            Set c = n.tag
            li.subItems(1) = n.Parent.Text 'each on their own line in case of error..
            li.subItems(2) = fso.FileNameFromPath(c.path)
            li.ToolTipText = c.path 'does not work with topmost
            If Not img Is Nothing Then li.SmallIcon = n.Image
        End If
    Next
    
    Me.Show
    
End Sub

Private Sub Form_Load()
    lv.SetColumnHeaders "Name*,Group,File", "2085,1365,4060"
    lv.SetFont "Courier", 12
    lv.Move 0, 0
    FormPos Me, True
    'SetWindowTopMost Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lv.Width = Me.ScaleWidth - 200
    lv.Height = Me.ScaleHeight - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    FormPos Me, True, True
End Sub

Private Sub lv_DblClick()
    On Error Resume Next
    If Not selNode Is Nothing Then handleCmd "show:" & selNode.Text
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Dim n As Node
    Set n = Item.tag
    n.EnsureVisible
    n.selected = True
    Set selNode = n
End Sub

'allow drag drop from listview to treeview in user document
Private Sub lv_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    On Error Resume Next
    Dim c As CVBComponent, n As Node
    AllowedEffects = vbDropEffectMove
    Set c = lv.SelectedItem.tag
    Set n = c.n
    Data.SetData n.key
    blnDragging = True
End Sub

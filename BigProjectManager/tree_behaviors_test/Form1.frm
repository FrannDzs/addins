VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   4200
      TabIndex        =   2
      Top             =   2580
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   5820
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   2235
   End
   Begin MSComctlLib.TreeView tvProject 
      Height          =   4035
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   7117
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "img1"
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   4740
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":059A
            Key             =   "form"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":115E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14B0
            Key             =   "mdichild"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1804
            Key             =   "bas"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D9E
            Key             =   "cls"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2338
            Key             =   "prop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":28D2
            Key             =   "ctl"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2E6C
            Key             =   "property"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2FC6
            Key             =   "func"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3318
            Key             =   "dob"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":38B2
            Key             =   "connect"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3E4C
            Key             =   "proj"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuMove 
         Caption         =   "Move"
         Begin VB.Menu mnuMoveUp 
            Caption         =   "Up"
         End
         Begin VB.Menu mnuMoveDown 
            Caption         =   "Down"
         End
         Begin VB.Menu mnuSpacer3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMoveTop 
            Caption         =   "Top"
         End
         Begin VB.Menu mnuMoveBottom 
            Caption         =   "Bottom"
         End
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
         Begin VB.Menu mnuAddGroup 
            Caption         =   "Group"
         End
         Begin VB.Menu mnuAddFolder 
            Caption         =   "Folder"
         End
         Begin VB.Menu mnuAddFile 
            Caption         =   "File"
         End
      End
      Begin VB.Menu mnuRemoveItem 
         Caption         =   "Remove Item"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'https://www.developerfusion.com/article/77/treeview-control/8/

'// variable that tells us if
'// we are dragging (ie the user is dragging a node from this treeview control
'// or not (ie the user is trying to drag an object from another
'// control and/or program)
Private blnDragging As Boolean
Private selNode As Node
Dim dlg As New CCmnDlg
Dim fso As New CFileSystem2

Private Sub Command1_Click()
    Form_Load
End Sub

Private Sub Form_Load()
    '// fill the control with some dummy nodes
    Dim n As Node
    
    tvProject.Nodes.Clear
    
    With tvProject.Nodes
        .Add , , "Root", "Root Item", "proj"
        '// add some child folders
        .Add "Root", tvwChild, "ChildFolder1", "Child Folder 1", "folder"
        .Add "Root", tvwChild, "ChildFolder2", "Child Folder 2", "folder"
        .Add "Root", tvwChild, "ChildFolder3", "Child Folder 3", "folder"
        '// add some children to the folders
        Set n = .Add("ChildFolder1", tvwChild, "c:\file1.bas", "file1.bas", "bas")
        .Add "ChildFolder1", tvwChild, "c:\file2.cls", "file2.cls", "cls"
        .Add "ChildFolder1", tvwChild, "c:\file3.frm", "file3.frm", "form"
        .Add "ChildFolder1", tvwChild, "c:\file4.frm", "file4.frm", "form"
        .Add "ChildFolder1", tvwChild, "c:\file5.frm", "file5.frm", "form"
        .Add "ChildFolder1", tvwChild, "c:\file6.frm", "file6.frm", "form"
        
        .Add "ChildFolder2", tvwChild, "c:\file7.cls", "file7.cls", "cls"
        .Add "ChildFolder2", tvwChild, "c:\file8.frm", "file8.frm", "form"
        .Add "ChildFolder2", tvwChild, "c:\file9.frm", "file9.frm", "form"
        .Add "ChildFolder2", tvwChild, "c:\file10.frm", "file10.frm", "form"
        .Add "ChildFolder2", tvwChild, "c:\file11.frm", "file11.frm", "form"
        
        
    End With
    

    For Each n In tvProject.Nodes
        n.Expanded = True
    Next
    
End Sub

Private Sub mnuAddFolder_Click()
        
        On Error Resume Next
        Dim f As String, p As Node, fn As String
        
        If selNode Is Nothing Then selNode = tvProject.Nodes(1)
        
        f = dlg.FolderDialog2()
        If Len(f) = 0 Then Exit Sub
        
        fn = fso.FolderName(f)
        Set p = tvProject.Nodes.Add(selNode, tvwChild, f, fn, "folder")
        AddFolderToTree f, p
            
End Sub

Function AddFolderToTree(f As String, p As Node, Optional recursive As Boolean = True)
    
        Dim ff() As String, x, n As Node, bn As String, pp As Node, fn As String, icon As String
        
        If Not fso.FolderExists(f) Then Exit Function
        
        fn = fso.FolderName(f)
        ff = fso.GetFolderFiles(f)
        p.Expanded = True
        
        For Each x In ff
            bn = fso.FileNameFromPath(CStr(x))
            If fileTypeOk(x, icon) Then
                Set n = tvProject.Nodes.Add(p, tvwChild, x, bn, icon)
                If Err.Number <> 0 Then
                    Debug.Print "failed to add " & bn & " " & Err.Description
                    Err.Clear
                Else
                    n.Tag = x
                End If
            End If
        Next
        
        If recursive Then
            ff = fso.GetSubFolders(f)
            If Not AryIsEmpty(ff) Then
                For Each x In ff
                    bn = fso.FolderName(CStr(x))
                    Set pp = tvProject.Nodes.Add(p, tvwChild, x, bn, "folder")
                    pp.Expanded = True
                    AddFolderToTree CStr(x), pp
                    
                Next
            End If
        End If
        
End Function



Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function fileTypeOk(f, ByRef icon As String) As Boolean
    
    Dim ext As String
    icon = Empty
    ext = LCase(fso.GetExtension(f))
    
    Select Case ext
        Case ".frm": icon = "form"
        Case ".bas": icon = "bas"
        Case ".cls": icon = "cls"
        Case ".dob": icon = "dob"
        Case ".ctl": icon = "ctl"
    End Select
    
    If Len(icon) > 0 Then fileTypeOk = True
    
End Function

Function Children(ByVal n As Node, c As Collection) As Long
    
    Dim nn As Node
    
    Set c = New Collection
    If n Is Nothing Then Exit Function
    If n.Children = 0 Then Exit Function
    
    Set nn = n.Child
    'Debug.Print nn.Text
    c.Add nn
    
    For i = 1 To n.Children - 1
        Set nn = nn.Next
        c.Add nn
        'Debug.Print nn.Text
    Next
    
    Children = c.Count
    
End Function

Sub AllNodesUnder(ByVal n As Node, c As Collection)
    
    Dim nn As Node
    If c Is Nothing Then Set c = New Collection
    c.Add n
    
    For Each nn In tvProject.Nodes
        If Not nn.Parent Is Nothing Then
            If nn.Parent = n Then
                c.Add nn
                If nn.Children > 0 Then AllNodesUnder nn, c
            End If
        End If
    Next
    
End Sub
    
Private Sub mnuAddGroup_Click()
    On Error Resume Next
    Dim nn As String, f As String
    
    If selNode Is Nothing Then Set selNode = tvProject.Nodes(1)
    
    nn = selNode.Text
    
    f = InputBox("Enter name of new folder to add under " & nn)
    If Len(f) = 0 Then Exit Sub
    
    tvProject.Nodes.Add selNode, tvwChild, f, f, "folder"
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuFind_Click()
    frmFind.init tvProject
End Sub

Private Sub mnuMoveBottom_Click()
    'On Error Resume Next
    If selNode Is Nothing Then Exit Sub
    Dim n As Node, p As Node, c As Collection
    Set p = selNode.Parent
    
    If Children(p, c) > 1 Then
        For i = c.Count To 1 Step -1
            Set n = c(i)
            If n <> selNode Then
                Set n.Parent = p
            End If
        Next
    End If
    

End Sub

Private Sub mnuMoveDown_Click()

    If selNode Is Nothing Then Exit Sub
    Dim n As Node, p As Node, c As Collection, targetAt As Long
    
    Set p = selNode.Parent
    
    If Children(p, c) > 1 Then
    
        For i = 1 To c.Count
            Set n = c(i)
            If ObjPtr(n) = ObjPtr(selNode) Then
                targetAt = i
                Exit For
            End If
        Next
        
        If targetAt <> 0 And targetAt <> c.Count Then
            Set selNode.Parent = p
            For i = targetAt + 1 To 1 Step -1
                If i <> targetAt Then
                    Set n = c(i)
                    Set n.Parent = p
                    Debug.Print "moving " & n.Text
                End If
            Next
        End If
        
    End If
    
End Sub

Private Sub mnuMoveTop_Click()
'On Error Resume Next
    If selNode Is Nothing Then Exit Sub
    Dim p As Node
    Set p = selNode.Parent
    If Not p Is Nothing Then Set selNode.Parent = p
End Sub

Private Sub mnuMoveUp_Click()

    If selNode Is Nothing Then Exit Sub
    Dim n As Node, p As Node, c As Collection, targetAt As Long, nn As Node

    Set p = selNode.Parent
    
    If Children(p, c) > 1 Then
    
        For i = 1 To c.Count
            Set n = c(i)
            If ObjPtr(n) = ObjPtr(selNode) Then
                targetAt = i
                Exit For
            End If
        Next
        
        If targetAt > 1 Then 'not 0 not 1
            Set selNode.Parent = p
            For i = targetAt - 2 To 1 Step -1
                If i <> targetAt Then
                    Set n = c(i)
                    Set n.Parent = p
                    Debug.Print "moving " & n.Text
                End If
            Next
        End If
        
    End If
    
End Sub

Private Sub mnuRemoveItem_Click()

    Dim n As Node
    Dim c As New Collection
    
    If selNode Is Nothing Then Exit Sub
    
    'MsgBox selNode.Image
   
    If selNode.Children > 0 Then
        If MsgBox("Are you sure you want to delete " & selNode.Children & " nodes?", vbYesNo) = vbNo Then Exit Sub
        AllNodesUnder selNode, c
        For i = c.Count To 1 Step -1
            Set n = c(i)
            tvProject.Nodes.Remove n.Key
        Next
        Exit Sub
    End If

    tvProject.Nodes.Remove selNode.Key
    Set selNode = Nothing
    
    ' If selNode.Image = "folder" Then
    
    
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Text1 = Data.Files(1)
End Sub

Private Sub tvProject_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
    Dim nodNode As Node
    '// get the node we are over
    Set nodNode = tvProject.HitTest(x, y)
    If nodNode Is Nothing Then Exit Sub '// no node
    '// ensure node is actually selected, just incase we start dragging.
    nodNode.selected = True
End Sub

Private Sub tvProject_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub tvProject_NodeClick(ByVal Node As MSComctlLib.Node)
    Set selNode = Node
End Sub

'// occurs when the user starts dragging
'// this is where you assign the effect and the data.
Private Sub tvProject_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    
    List1.AddItem "start drag"
    
    '// Set the effect to move
    AllowedEffects = vbDropEffectMove
    '// Assign the selected item's key to the DataObject
    Data.SetData tvProject.SelectedItem.Key
    '// we are dragging from this control
    blnDragging = True
End Sub

'note if your running as admin, you cant drop files from the desktop as thats not admin..

'// occurs when the object is dragged over the control.
'// this is where you check to see if the mouse is over
'// a valid drop object
Private Sub tvProject_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, shift As Integer, x As Single, y As Single, State As Integer)
    Dim nodNode As Node
    '// set the effect
    'List1.AddItem "drag over"
    
    
    Effect = vbDropEffectMove
    '// get the node that the object is being dragged over
    Set nodNode = tvProject.HitTest(x, y)
    If nodNode Is Nothing Or blnDragging = False Then
        '// the dragged object is not over a node, invalid drop target
        '// or the object is not from this control.
        
        If Not Data.GetFormat(vbCFFiles) Then
            Effect = vbDropEffectNone 'setting this will block the transfer further..
        End If
        
    End If
End Sub

Function NodeExists(fPath As String) As Boolean
    On Error Resume Next
    Dim n As Node
    Set n = tvProject.Nodes(fPath)
    NodeExists = (Err.Number = 0)
End Function

'Text = 1 (vbCFText)
'Bitmap = 2 (vbCFBitmap)
'Metafile = 3
'Emetafile = 14
'DIB = 8
'Palette = 9
'Files = 15 (vbCFFiles)
'RTF = -16639

Sub MoveNodeBelow(n As Node, target As Node)

    Dim nn As Node, p As Node, c As Collection, targetAt As Long
    
    If n.Parent <> target.Parent Then Exit Sub
    
    Set p = n.Parent
    
    If Children(p, c) > 1 Then
    
        For i = 1 To c.Count
            Set nn = c(i)
            If ObjPtr(nn) = ObjPtr(target) Then
                targetAt = i
                Exit For
            End If
        Next
        
        If targetAt = c.Count Then
            Set selNode = n
            mnuMoveBottom_Click
            Exit Sub
        End If
        
        Debug.Print "targetAt: " & targetAt
        
        If targetAt > 1 Then
            Set n.Parent = p
            For i = targetAt To 1 Step -1
                Set nn = c(i)
                If ObjPtr(nn) <> ObjPtr(n) Then
                    Set nn.Parent = p
                    Debug.Print "moving " & nn.Text
                End If
            Next
        End If
        
    End If

End Sub


'// occurs when the user drops the object
'// this is where you move the node and its children.
'// this will not occur if Effect = vbDropEffectNone
Private Sub tvProject_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, shift As Integer, x As Single, y As Single)
    Dim strSourceKey As String
    Dim nodTarget    As Node
    Dim f As String, fn As String, p As Node, icon As String
    Dim positionNode As Node
    
    'List1.AddItem "drag drop"
    
    '// get the target node
    Set nodTarget = tvProject.HitTest(x, y)
    If nodTarget Is Nothing Then Set nodTarget = tvProject.Nodes(1)
    
    '// if the target node is not a folder or the root item
    '// then get it's parent (that is a folder or the root item)
    If nodTarget.Image <> "proj" And nodTarget.Image <> "folder" Then
        Set positionNode = nodTarget
        Set nodTarget = nodTarget.Parent
    End If
    
    If Data.GetFormat(vbCFText) Then
        strSourceKey = Data.GetData(vbCFText) '// get the carried data
        Set tvProject.Nodes(strSourceKey).Parent = nodTarget '// move the source node to the target node
        If Not positionNode Is Nothing Then
            MoveNodeBelow tvProject.SelectedItem, positionNode
        End If
    ElseIf Data.GetFormat(vbCFFiles) Then
        
        f = Data.Files(1)
        If NodeExists(f) Then
            MsgBox "This path already exists in tree"
            Exit Sub
        End If
        
        If fso.FolderExists(f) Then
            fn = fso.FolderName(f)
            Set p = tvProject.Nodes.Add(nodTarget, tvwChild, f, fn, "folder")
            AddFolderToTree f, p
        Else
            fn = fso.FileNameFromPath(CStr(f))
            If fileTypeOk(f, icon) Then
                Set p = tvProject.Nodes.Add(nodTarget, tvwChild, f, fn, icon)
                p.Tag = f
            End If
        End If
    End If
        
    '// NOTE: You will also need to update the key to reflect the changes
    '// if you are using it
    '// we are not dragging from this control any more
    blnDragging = False
    Effect = 0 '// cancel effect so that VB doesn't muck up your transfer
    
End Sub

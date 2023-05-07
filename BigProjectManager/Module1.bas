Attribute VB_Name = "Module1"
Option Explicit

Global fso As New CFileSystem2
Global dlg As New CCmnDlg

Global mToolWindow As VBIDE.Window
Global mUserDoc As UserDoc
Global VBInstance As VBIDE.VBE

Global unloading As Boolean
Global mainProjPath As String
Global mProjectLoaded As Boolean
Global mProjectTreeHwnd As Long
Global FormDisplayed As Boolean
Global blnDragging As Boolean  'used in userdoc and frmFind for cross form drag drops

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

'left over from IPC version...
Function handleCmd(m)

    On Error Resume Next
    Dim c As VBComponent, i As Long, tmp As String, p As VBProject, fn As String, j As Long
    Dim x() As String
    
    If Left(m, 8) = "addfile:" Then
        m = Mid(m, 9)
        If FileExists(CStr(m)) Then
            VBInstance.ActiveVBProject.VBComponents.AddFile CStr(m)
        End If
    End If
    
    If Left(m, 12) = "showrelated:" Then
        m = LCase(Mid(m, 13))
        ShellExecute 0, "open", m, vbNullString, vbNullString, SW_SHOWNORMAL
        Exit Function
    End If
    
    If Left(m, 5) = "show:" Then
        m = LCase(Mid(m, 6))
        For Each c In VBInstance.ActiveVBProject.VBComponents
            If LCase(c.name) = m Then
                If c.Type = vbext_ct_RelatedDocument Then
                    ShellExecute 0, "open", c.FileNames(1), vbNullString, vbNullString, SW_SHOWNORMAL
                Else
                    c.CodeModule.CodePane.Show
                End If
                Exit For
            End If
        Next
        Exit Function
    End If
    
    If Left(m, 7) = "remove:" Then
        m = LCase(Mid(m, 8))
        For Each c In VBInstance.ActiveVBProject.VBComponents
            If c.Type = vbext_ct_RelatedDocument Then
                If fso.FileNameFromPath(c.FileNames(1)) = m Then
                    VBInstance.ActiveVBProject.VBComponents.Remove c
                    Exit For
                End If
            ElseIf LCase(c.name) = m Then
                VBInstance.ActiveVBProject.VBComponents.Remove c
                Exit For
            End If
        Next
        Exit Function
    End If
    
    If Left(m, 9) = "designer:" Then
        m = LCase(Mid(m, 10))
        For Each c In VBInstance.ActiveVBProject.VBComponents
            If LCase(c.name) = m Then
                c.DesignerWindow.Visible = True
                c.DesignerWindow.WindowState = vbext_ws_Normal
                Exit For
            End If
        Next
        Exit Function
    End If
    
    If m = "list" Then
        For Each c In VBInstance.ActiveVBProject.VBComponents
            'name may be different than file name, file name may not exist if not yet saved...
            tmp = c.Type & "|" & c.name & "|" & c.FileNames(1)
            push x, tmp
        Next
        tmp = Join(x, vbCrLf)
        'memfile.WriteFile tmp, , True 'maybe > than our 2048 ipc send buffer...
        'ipc.Send Len(tmp)
        handleCmd = tmp
        Exit Function
    End If
    
    If m = "projects" Then
        tmp = VBInstance.VBProjects.Count & "|"
        For Each p In VBInstance.VBProjects
            tmp = tmp & p.name & "|"
        Next
        tmp = Mid(tmp, 1, Len(tmp) - 1)
        'ipc.Send tmp
        handleCmd = tmp
        Exit Function
    End If
    
    If Err.Number <> 0 Then
        Debug.Print "Error in handleCmd: " & Err.Description
    End If
    
End Function

Function Children(ByVal n As Node, c As Collection) As Long
    
    Dim nn As Node, i As Long
    
    Set c = New Collection
    If n Is Nothing Then Exit Function
    If n.Children = 0 Then Exit Function
    
    Set nn = n.Child
    c.Add nn
    'Debug.Print nn.Text
    
    For i = 1 To n.Children - 1
        Set nn = nn.Next
        c.Add nn
        'Debug.Print nn.Text
    Next
    
    Children = c.Count
    
End Function

Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function


Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub


Function PreloadComponentName(fpath As String) As String
    'Attribute VB_Name = "Connect"
    On Error GoTo hell

    Dim x As String, i As Long
    Const marker = "Attribute VB_Name = "
    
    If Not fso.fOpen(fpath, otreading) Then Exit Function
    
    Do While Not fso.EndOfFile
        x = fso.ReadLine
        If InStr(x, marker) > 0 Then
            x = Replace(x, marker, Empty)
            x = Replace(x, vbCr, Empty)
            x = Replace(x, vbLf, Empty)
            x = Trim(Replace(x, """", Empty))
            PreloadComponentName = x
            Exit Do
        End If
    Loop

hell:
    fso.fClose
           
End Function

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function
 

Sub AllNodesUnder(tv As TreeView, ByVal n As Node, c As Collection)
    
    Dim nn As Node
    c.Add n
    
    For Each nn In tv.Nodes
        If Not nn.Parent Is Nothing Then
            If nn.Parent = n Then
                c.Add nn
                If nn.Children > 0 Then AllNodesUnder tv, nn, c
            End If
        End If
    Next
    
End Sub



Sub FormPos(fform As Form, Optional andSize As Boolean = False, Optional save_mode As Boolean = False)
    
    On Error Resume Next
    
    Dim f, sz, i, ff, def
    f = Split(",Left,Top,Height,Width", ",")
    
    If fform.WindowState = vbMinimized Then Exit Sub
    If andSize = False Then sz = 2 Else sz = 4
    
    For i = 1 To sz
        If save_mode Then
            ff = CallByName(fform, f(i), VbGet)
            SaveSetting App.EXEName, fform.name & ".FormPos", f(i), ff
        Else
            def = CallByName(fform, f(i), VbGet)
            ff = GetSetting(App.EXEName, fform.name & ".FormPos", f(i), def)
            CallByName fform, f(i), VbLet, ff
        End If
    Next
    
End Sub

Function keyExistsInCollection(key As String, c As Collection, Optional isObj As Boolean = True) As Boolean
    On Error Resume Next
    Dim o As Object, x
    If isObj Then
        Set o = c(key)
    Else
        x = c(key)
    End If
    keyExistsInCollection = (Err.Number = 0)
End Function

Function ComponentExists(tv As TreeView, name As String, Optional ByRef c As CVBComponent) As Boolean
    On Error Resume Next
    Dim n As Node
    Set c = Nothing
    If NodeExists(tv, name, n) Then
        Set c = n.tag
        ComponentExists = (Not c Is Nothing)
    End If
End Function

Function NodeExists(tv As TreeView, key As String, Optional ByRef n As Node) As Boolean
    On Error Resume Next
    
    'For Each n In tv.Nodes: Debug.Print n.Text & ", " & n.key: Next
    
    Set n = tv.Nodes(key)
    NodeExists = (Err.Number = 0)
End Function

Function HandleComponentEvent(tv As TreeView, e As CComponentEvent, Optional createMissing As Boolean = True) As CVBComponent
    
    On Error GoTo hell
    
    Dim c As CVBComponent
    Dim n As String
    Dim p As Node, nn As Node
    
    If tv.Nodes.Count = 0 Then Exit Function
    
    n = e.ComponentName
    
    If e.ComponentType = &HA Then 'related document
        n = fso.FileNameFromPath(e.filename)
    End If
    
    If e.EventType = ec_Rename Then n = e.OldName '(cant rename a related document)
    
    If ComponentExists(tv, n, c) Then
    
        Set HandleComponentEvent = c
        
        If e.EventType = ec_Remove Then
            tv.Nodes.Remove c.n.key
            Set c = Nothing
            Set HandleComponentEvent = Nothing
            Exit Function
        End If
            
        If e.EventType = ec_Rename Then
            c.name = e.ComponentName
            If Not c.n Is Nothing Then
                Set p = c.n.Parent
                tv.Nodes.Remove c.n.key 'we need to reset its key and new name text
                Set c.n = Nothing
                Set nn = tv.Nodes.Add(p, tvwChild, c.name, c.name, c.icon)
                Set c.n = nn
                Set nn.tag = c
            End If
            Exit Function
        End If
        
    Else
        If createMissing Then
            Set c = New CVBComponent
            c.loadFromEvent e
            
            If Not NodeExists(tv, c.defFolder, p) Then
                Set p = tv.Nodes.Add(tv.Nodes(1), tvwChild, c.defFolder, c.defFolder, "folder")
            End If
            
            If NodeExists(tv, c.name) Then
                 'List1.AddItem "HandleComponentEvent Node exists: " & e.raw
            Else
                Set c.n = tv.Nodes.Add(p, tvwChild, c.name, c.name, c.icon)
                Set c.n.tag = c
            End If
        End If
    End If
    
Exit Function
hell:
    Debug.Print "Err in HandleComponentEvent:" & Err.Description & " " & e.raw

End Function


'Public Enum vbext_ComponentType
'    vbext_ct_StdModule = 1
'    vbext_ct_ClassModule = 2
'    vbext_ct_MSForm = 3
'    vbext_ct_ResFile = 4
'    vbext_ct_VBForm = 5
'    vbext_ct_VBMDIForm = 6
'    vbext_ct_PropPage = 7
'    vbext_ct_UserControl = 8
'    vbext_ct_DocObject = 9
'    vbext_ct_RelatedDocument = &HA
'    vbext_ct_ActiveXDesigner = &HB
'End Enum

Function typeHasDesigner(t As Long) As Boolean
    typeHasDesigner = True
    Select Case t
        Case 1, 2, 4, &HA: typeHasDesigner = False
    End Select
End Function

Function typeFromPath(fpath As String) As Long

    On Error Resume Next
    Dim ext As String, i As Long
    
    ext = LCase(fso.GetExtension(fpath))
    If Left(ext, 1) = "." Then ext = Mid(ext, 2)
    
    Select Case ext
        Case "bas": i = 1
        Case "cls": i = 2
        Case "frm": i = 3
        Case "res": i = 4
        Case "frm": i = 5
        Case "mdi": i = 6
        Case "pag": i = 7
        Case "ctl": i = 8
        Case "dob": i = 9
        Case "txt": i = 10
        Case "dsr": i = 11
    End Select
      
    typeFromPath = i
    
End Function

Function DefaultFolderForType(t As Long, Optional ByRef icon As String) As String
    On Error Resume Next
    Dim tn  As String
    
    Select Case t
        Case 1: tn = "Modules"
                icon = "bas"
                
        Case 2: tn = "Classes"
                icon = "cls"

        Case 3: tn = "Forms"
                icon = "frm"

        Case 4: tn = "Resources"
                icon = "res"

        Case 5: tn = "Forms"
                icon = "frm"

        Case 6: tn = "Forms"
                icon = "mdi"

        Case 7: tn = "Property Pages"
                icon = "pag"

        Case 8: tn = "User Controls"
                icon = "ctl"

        Case 9: tn = "User Documents"
                icon = "dob"

        Case 10: tn = "Related Documents"
                icon = "txt"

        Case 11: tn = "Designers"
                icon = "dsr"

        Case Default:
                tn = "Unknown"
                icon = "unk"
                
    End Select
      
    DefaultFolderForType = tn

End Function

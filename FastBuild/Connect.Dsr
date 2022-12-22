VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   12630
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   17685
   _ExtentX        =   31194
   _ExtentY        =   22278
   _Version        =   393216
   Description     =   "Streamline Build Process "
   DisplayName     =   "Fast Build "
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'this used to use the API Hooking code in Module1.bas, until i learned
'vb actually already exposed the necessary events as part of the addin model..
'oops there goes a solid days labor..
'
'It would be neat to hook createprocessA like compiler control, and give build window output
'like compiling xxx.obj, linking, running post build command..
'code is already there from chamberlin, but really just extra noise..except for the postbuild output..
'
'also would be nice to be able to write postbuild command output to debug window..but havent found way yet..

'Clipboard.SetData LoadPicture("c:\MyPic.bmp") 'or picBox.Picture or an img from a resource file
'    .PasteFace errs sometimes..
'You need to use an IPictureDisp picture object type in order to set the .Picture property.
'CommandBarButton class instead of CommandBarControl, and it had the .Picture
'I dont believe that the .Picture property is available in 2000, only 2002 and 2003.

Private FormDisplayed As Boolean
Dim mfrmAddIn As frmAddIn

Dim mcbFastBuildUI As Office.CommandBarControl
Private WithEvents mnuFastBuildUI As CommandBarEvents
Attribute mnuFastBuildUI.VB_VarHelpID = -1

Dim mcbFastBuild As Office.CommandBarControl
Private WithEvents mnuFastBuild As CommandBarEvents
Attribute mnuFastBuild.VB_VarHelpID = -1

Private WithEvents FileEvents As VBIDE.FileControlEvents
Attribute FileEvents.VB_VarHelpID = -1

Dim mcbExecute As Office.CommandBarControl
Private WithEvents mnuExecute As CommandBarEvents
Attribute mnuExecute.VB_VarHelpID = -1

Dim mcbDebugWindow As Office.CommandBarControl
Private WithEvents mnuDebugWindow As CommandBarEvents
Attribute mnuDebugWindow.VB_VarHelpID = -1

Dim mcbAddref As Office.CommandBarControl
Private WithEvents mnuAddref As CommandBarEvents
Attribute mnuAddref.VB_VarHelpID = -1

Dim mcbImmediate As Office.CommandBarControl
Private WithEvents mnuImmediate As CommandBarEvents
Attribute mnuImmediate.VB_VarHelpID = -1

Dim mcbAddFiles As Office.CommandBarControl
Private WithEvents mnuAddFiles As CommandBarEvents
Attribute mnuAddFiles.VB_VarHelpID = -1

Dim mcbMemWindow As Office.CommandBarControl
Private WithEvents mnuMemWindow As CommandBarEvents
Attribute mnuMemWindow.VB_VarHelpID = -1

Dim mcbCodeDB As Office.CommandBarControl
Private WithEvents mnuCodeDB As CommandBarEvents
Attribute mnuCodeDB.VB_VarHelpID = -1

Dim mcbApiAddin As Office.CommandBarControl
Private WithEvents mnuApiAddin As CommandBarEvents
Attribute mnuApiAddin.VB_VarHelpID = -1

Dim mcbOpenHomeDir As Office.CommandBarControl
Private WithEvents mnuOpenHomeDir As CommandBarEvents
Attribute mnuOpenHomeDir.VB_VarHelpID = -1

Dim mcbRealMakeMenu As Office.CommandBarControl

'vb6 ide bug..if you hold a reference to an existing button in an addin..it will disable the button
'when you enter the run state as if it was owned by the addin..just use f5 or runstart button then..
'maybe I can switch to using a hooklib to hook it at a lower level so this bug manifest. I have a 20yr habit of
'pressing that stupid run arrow rather than reaching for the f5 key everytime.. no one will like this..
'Public mcbRealStartButton As Office.CommandBarControl
'Public WithEvents mnuRealRun As CommandBarEvents 'hook into existing controls events


Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmAddIn.Hide
   
End Sub

Sub Show()
  
    On Error Resume Next
    Dim needsRefresh As Boolean
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmAddIn
    Else
        needsRefresh = True
    End If
    
    If Not VBInstance.ActiveVBProject Is Nothing Then
        Debug.Print "OnConnect Project: " & VBInstance.ActiveVBProject.FileName
    Else
        Debug.Print "VBInstance.ActiveVBProject is nothing "
    End If
    
    FormDisplayed = True
    mfrmAddIn.Show
    
    'If needsRefresh Then mfrmAddIn.cmdRefresh_Click
    
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    On Error Resume Next
    
    'save the vb instance
    If VBInstance Is Nothing Then Set VBInstance = Application
    'If Connect Is Nothing Then Set Module2.Connect = Me
        
    'this is a good place to set a breakpoint and
    'test various addin objects, properties and methods
    'Debug.Print "FullName: " & VBInstance.FullName
     
    If ConnectMode = ext_cm_External Then
        'Used by the wizard toolbar to start this wizard
        Me.Show
    Else
    
        If GetSetting("FastBuild", "Settings", "DisplayAsHex", 0) = "1" Then
            LoadHexToolTipsDll
        End If
        
        MemWindowExe = App.path & "\MemoryWindow\standalone.exe"
        CodeDBExe = App.path & "\CodeDB\CodeDB.exe"
        APIAddInExe = App.path & "\API_AddIn\API_AddIn.exe"
        ExternalDebugWindow = App.path & "\DbgWindow\DbgWindow.exe"
        
        ClearImmediateOnStart = GetSetting("fastbuild", "settings", "ClearImmediateOnStart", 0)
        ShowPostBuildOutput = GetSetting("fastbuild", "settings", "ShowPostBuildOutput", 1)
        
        Set mcbFastBuildUI = AddButton("Fast Build", 101)
        If Not mcbFastBuildUI Is Nothing Then
            Set mnuFastBuildUI = VBInstance.Events.CommandBarEvents(mcbFastBuildUI)
        End If
        
        Set mcbExecute = AddButton("Execute", 102)
        If Not mcbExecute Is Nothing Then
            Set mnuExecute = VBInstance.Events.CommandBarEvents(mcbExecute)
        End If
        
        Set mcbAddref = AddToAddInCommandBar("Quick AddRef")
        If Not mcbAddref Is Nothing Then
            Set mnuAddref = VBInstance.Events.CommandBarEvents(mcbAddref)
        End If

        Set mcbAddFiles = AddToAddInCommandBar("Add Multiple Files")
        If Not mcbAddFiles Is Nothing Then
            Set mnuAddFiles = VBInstance.Events.CommandBarEvents(mcbAddFiles)
        End If

        Set mcbOpenHomeDir = AddToAddInCommandBar("Open Project Directory")
        If Not mcbOpenHomeDir Is Nothing Then
            Set mnuOpenHomeDir = VBInstance.Events.CommandBarEvents(mcbOpenHomeDir)
        End If

'       if we add these to the Projects menu..vb freaks out on disconnect and saves weird and legit ones go missing..fuck it
'        Set mcbAddref = AddrefMenu("Quick AddRef")
'        If Not mcbAddref Is Nothing Then
'            Set mnuAddref = VBInstance.Events.CommandBarEvents(mcbAddref)
'        End If
'
'        Set mcbAddFiles = AddrefMenu("Add Multiple Files", , "")
'        If Not mcbAddFiles Is Nothing Then
'            Set mnuAddFiles = VBInstance.Events.CommandBarEvents(mcbAddFiles)
'        End If
'
'        Set mcbOpenHomeDir = AddrefMenu("Open Project Directory", , "")
'        If Not mcbOpenHomeDir Is Nothing Then
'            Set mnuOpenHomeDir = VBInstance.Events.CommandBarEvents(mcbOpenHomeDir)
'        End If
        
        'external utilities
        '-----------------------------------------------------------------------------
        If FileExists(MemWindowExe) Then
            Set mcbMemWindow = AddButton("Memory Window", 105)
            If Not mcbMemWindow Is Nothing Then
                Set mnuMemWindow = VBInstance.Events.CommandBarEvents(mcbMemWindow)
            End If
        End If
        
        If FileExists(CodeDBExe) Then
            Set mcbCodeDB = AddrefMenu("Code-DB", "&Add-Ins", "")
            If Not mcbCodeDB Is Nothing Then
                Set mnuCodeDB = VBInstance.Events.CommandBarEvents(mcbCodeDB)
            End If
        End If

        If FileExists(APIAddInExe) Then
            Set mcbApiAddin = AddrefMenu("Api-Viewer++", "&Add-Ins", "")
            If Not mcbApiAddin Is Nothing Then
                Set mnuApiAddin = VBInstance.Events.CommandBarEvents(mcbApiAddin)
            End If
        End If
        
        If FileExists(ExternalDebugWindow) Then
            Set mcbDebugWindow = AddButton("External Debug Window", 106)
            If Not mcbDebugWindow Is Nothing Then
                Set mnuDebugWindow = VBInstance.Events.CommandBarEvents(mcbDebugWindow)
            Else
                'MsgBox "add debug window button failed"
            End If
        Else
            'MsgBox "no ExternalDebugWindow:" & ExternalDebugWindow
        End If
        '-----------------------------------------------------------------------------
        
        Set FileEvents = Application.Events.FileControlEvents(Nothing)
        
        'hook the run button events (side effect it will disbale during debugging so you have to use f5..sucky!
        'we do this to clear immediate window on start which annoyed me too..
'        If ClearImmediateOnStart = 1 Then
'            Set mcbRealStartButton = FindRunButton()
'            If Not mcbRealStartButton Is Nothing Then
'                Set mnuRealRun = VBInstance.Events.CommandBarEvents(mcbRealStartButton)
'            End If
'        End If

        Set mcbRealMakeMenu = FindMakeMenu()
        
        Set mcbFastBuild = AddButton("Compile", 103)
        If Not mcbFastBuild Is Nothing Then
                Set mnuFastBuild = VBInstance.Events.CommandBarEvents(mcbFastBuild)
        End If
        
        Set mcbImmediate = AddButton("Start and Clear Immediate", 104)
        If Not mcbImmediate Is Nothing Then
            Set mnuImmediate = VBInstance.Events.CommandBarEvents(mcbImmediate)
        End If
                
        Load frmIPC
        SaveSetting "fastbuild", "ipc", "hIpc", frmIPC.txtIPCServer.hwnd
        'frmIPC.Visible = True
        'MsgBox Hex(frmIPC.txtIPCServer.hwnd)
        
    End If

    Exit Sub
    
error_handler:
    
    MsgBox "FastBuild.AddinInstance_OnConnection: " & Err.Description
    
End Sub

Function showErr(X)
    If Err.Number <> 0 Then
        Debug.Print X & " Err: " & Err.Description
        Err.Clear
    End If
End Function
'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    
    On Error Resume Next
    'you can get an error in here if the user resets the toolbars on you you will still have a ref but delete fails..
    
    Dim f As Form
    
    If Not FileEvents Is Nothing Then Set FileEvents = Nothing
    If Not mcbRealMakeMenu Is Nothing Then Set mcbRealMakeMenu = Nothing
    
    'If Not mnuRealRun Is Nothing Then Set mnuRealRun = Nothing
    'If Not mcbRealStartButton Is Nothing Then Set mcbRealStartButton = Nothing
    
    Err.Clear
    
    'If Not mcbFastBuild Is Nothing Then
        mcbFastBuild.Delete
        showErr "mcbFastBuild.Delete"
    '    Set mcbFastBuild = Nothing
    '    Set mnuFastBuild = Nothing
    'End If
    
    'If Not mcbFastBuildUI Is Nothing Then
        mcbFastBuildUI.Delete
        showErr "mcbFastBuildUI.Delete"
    '    Set mcbFastBuildUI = Nothing
    '    Set mnuFastBuildUI = Nothing
    'End If
    
    'If Not mcbExecute Is Nothing Then
         mcbExecute.Delete
         showErr "mcbExecute.Delete"
    '     Set mcbExecute = Nothing
    '     Set mnuExecute = Nothing
    'End If
    
    'If Not mcbAddref Is Nothing Then
         mcbAddref.Delete
         showErr "mcbAddref.Delete"
    '     Set mcbAddref = Nothing
    '     Set mnuAddref = Nothing
    'End If
    
    'If Not mcbImmediate Is Nothing Then
         mcbImmediate.Delete
         showErr "mcImmed.Delete"
     '    Set mcbImmediate = Nothing
     '    Set mnuImmediate = Nothing
    'End If
    
    'If Not mcbApiAddin Is Nothing Then
         mcbApiAddin.Delete
         showErr "mcbApiAddin.Delete"
    '     Set mcbApiAddin = Nothing
    '     Set mnuApiAddin = Nothing
    'End If
    
    'If Not mcbMemWindow Is Nothing Then
         mcbMemWindow.Delete
         showErr "mcbMemWindow.Delete"
    '     Set mcbMemWindow = Nothing
    '     Set mnuMemWindow = Nothing
    'End If
    
    'If Not mcbCodeDB Is Nothing Then
         mcbCodeDB.Delete
         showErr "mcbCodeDB.Delete"
    '     Set mcbCodeDB = Nothing
    '     Set mnuCodeDB = Nothing
    'End If
    
    'If Not mcbAddFiles Is Nothing Then
         mcbAddFiles.Delete
         showErr "mcbAddFiles.Delete"
    '     Set mcbAddFiles = Nothing
    '     Set mnuAddFiles = Nothing
    'End If

    'If Not mcbOpenHomeDir Is Nothing Then
         mcbOpenHomeDir.Delete
         showErr "mcbOpenHome.Delete"
    '     Set mcbOpenHomeDir = Nothing
    '     Set mnuOpenHomeDir = Nothing
    'End If
    
    mcbDebugWindow.Delete
    showErr "mcbDebugWindow.Delete"
        
    If Not mfrmAddIn Is Nothing Then Set mfrmAddIn = Nothing
    
    For Each f In Forms
        Unload f
    Next
    
    'release all references so object can shut down and remove itself..
    'otherwise you wont be able to unload and compile, you will have to restart ide
    Set VBInstance = Nothing

Exit Sub
hell:
    
    'MsgBox "FastBuild.AddinInstance_OnDisconnection error: " & Err.Description
    
End Sub

Private Sub FileEvents_AfterWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal result As Integer)
    
    Dim postbuild As String
    Dim buildOutput As String
    Dim tmp As String, t2 As String
    
    If FileType <> vbext_ft_Exe Then Exit Sub
           
    If Not isBuildPathSet() Then
        tmp = GetParentFolder(VBInstance.ActiveVBProject.FileName)
        t2 = GetParentFolder(FileName)
        If InStr(1, tmp, t2) > 0 Then
            tmp = Replace(FileName, tmp, "%ap%\")
            tmp = Replace(tmp, "\\", "\")
        Else
            tmp = FileName
        End If
        VBInstance.ActiveVBProject.WriteProperty "fastBuild", "fullPath", tmp
    End If
    
    LastCommandOutput = Empty
    postbuild = GetPostBuildCommand()
    
    If Len(postbuild) > 0 Then
        SetHomeDir
        postbuild = ExpandVars(postbuild, FileName)
        LastCommandOutput = GetCommandOutput("cmd /c " & postbuild, True, True)
    End If
    
    postbuild = ConsoleAppCommand() 'do we have to change sub system? to console (also supports vblink.exe from linktool addin)
    If Len(postbuild) > 0 Then
        SetHomeDir
        postbuild = ExpandVars(postbuild, FileName)
        LastCommandOutput = LastCommandOutput & vbCrLf & GetCommandOutput("cmd /c " & postbuild, True, True)
    End If
    
    If ShowPostBuildOutput = 1 Then
        'MsgBox LastCommandOutput
        buildOutput = GetFileReport(FileName)
        If Len(LastCommandOutput) > 0 Then
            buildOutput = buildOutput & vbCrLf & vbCrLf & "Post Build Command Output: " & vbCrLf & String(50, "-") & vbCrLf & LastCommandOutput
        End If
        SetImmediateText Replace(buildOutput, Chr(0), Empty)
    End If
    
    
End Sub


Private Sub FileEvents_DoGetNewFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, NewName As String, ByVal OldName As String, CancelDefault As Boolean)
    Dim fastBuildPath As String
    Dim pf As String
    
    If FileType <> vbext_ft_Exe Then
        'MsgBox "Filetype: " & FileType
        Exit Sub
    End If
    
    If Not isBuildPathSet() Then
        'MsgBox "Build path not set"
        Exit Sub
    End If
     
    fastBuildPath = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
    If Len(fastBuildPath) = 0 Then
        'MsgBox "fast build path empty?"
        Exit Sub
    End If
    
    On Error Resume Next
    
    fastBuildPath = Replace(fastBuildPath, "%AP%", GetParentFolder(VBInstance.ActiveVBProject.FileName), , , vbTextCompare)
    pf = GetParentFolder(fastBuildPath)
    
    If FolderExists(pf) Then
        'MsgBox "overriding path! " & NewName & " old: " & OldName
        NewName = fastBuildPath
        OldName = fastBuildPath
        CancelDefault = True
    Else
        'msgbox "path is out of date project must have moved..resetting.."
        VBInstance.ActiveVBProject.WriteProperty "fastBuild", "fullPath", ""
    End If
 
End Sub

Private Sub mnuAddFiles_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    
    Dim f() As String
    Dim ff
    Dim e() As String
    
    On Error Resume Next
    
    f() = ShowOpenMultiSelect()
    If AryIsEmpty(f) Then Exit Sub
    
    For Each ff In f
        Err.Clear
        VBInstance.ActiveVBProject.VBComponents.AddFile ff
        If Err.Number <> 0 Then
            'push e, FileNameFromPath(ff) & ": " & Err.Description
            push e, Err.Description  'seems to already contain file names
        End If
    Next
    
    If AryIsEmpty(e) Then
        'MsgBox "All files imported no errors."
    Else
        MsgBox Join(e, vbCrLf), vbExclamation
    End If
    
End Sub

Private Sub mnuAddref_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    frmAddRefs.Show
End Sub

Private Sub mnuApiAddin_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If FileExists(APIAddInExe) Then
        Shell APIAddInExe, vbNormalFocus
    Else
        MsgBox "File not found: " & APIAddInExe
    End If
End Sub

Private Sub mnuCodeDB_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim cmdLine As String
    
    If FileExists(CodeDBExe) Then
        If Not VBInstance.ActiveVBProject Is Nothing Then
            cmdLine = " """ & VBInstance.ActiveVBProject.FileName & """"
        End If
        cmdLine = cmdLine & " hwnd=" & frmIPC.txtIPCServer.hwnd
        Shell CodeDBExe & cmdLine, vbNormalFocus
    Else
        MsgBox "File not found: " & CodeDBExe
    End If
End Sub

Private Sub mnuDebugWindow_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    
    If FileExists(ExternalDebugWindow) Then
        Shell ExternalDebugWindow
    End If
        
    If Err.Number <> 0 Then
        MsgBox "ExternalDebugWindow not found? " & ExternalDebugWindow & vbCrLf & vbCrLf & _
                "You can download it here: https://github.com/dzzie/libs/tree/master/_elroy/PersistentDebugPrint_mod1"
    End If
        
End Sub

Private Sub mnuFastBuildUI_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Private Sub mnuFastBuild_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next

    mcbRealMakeMenu.Execute
    
'I am removing this method..it has bugs in how MakeCompiledFile is implemented..
'if the path you specify in BuildFileName is not valid, then it will fail without error
'I could work around it, but its better to manually add a Build tool bar button from the command bar editor.
'
'    If isBuildPathSet() Then
'        VBInstance.ActiveVBProject.BuildFileName = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
'    End If
'
'    'if you want to readd this..first test that buildfilename path is valid for the system, then after compile test that
'    'the exe file was created and is different from what was there..
'
'    'apparently calling this method manually like this just uses the default and skips DoGetNewFileName hooks..
'    VBInstance.ActiveVBProject.MakeCompiledFile

End Sub

Private Sub mnuExecute_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim fastBuildPath As String
    Dim cmdLine As String
    
    If Not isBuildPathSet() Then
        MsgBox "Can not launch the executable, path not yet set", vbInformation
        Exit Sub
    End If
    
    fastBuildPath = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
    fastBuildPath = Replace(fastBuildPath, "%AP%", GetParentFolder(VBInstance.ActiveVBProject.FileName), , , vbTextCompare)
    
    If Not FileExists(fastBuildPath) Then
        MsgBox "File not found: " & fastBuildPath, vbInformation
        Exit Sub
    End If
    
    cmdLine = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "ExecBtnCmdLine")
    If Len(cmdLine) > 0 And Left(cmdLine, 1) <> " " Then cmdLine = " " & cmdLine
    Err.Clear
    
    Shell fastBuildPath & cmdLine, vbNormalFocus
    
    If Err.Number <> 0 Then
        MsgBox "Menu Execute Error: " & Err.Description, vbExclamation
    End If
    
End Sub


'Private Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
'    Dim cbMenuCommandBar As Office.CommandBarControl
'    Dim cbMenu As Object
'
'    On Error GoTo hell
'
'    Set cbMenu = VBInstance.CommandBars("Add-Ins")
'    If cbMenu Is Nothing Then Exit Function
'
'    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
'    cbMenuCommandBar.caption = sCaption
'    Set AddToAddInCommandBar = cbMenuCommandBar
'
'    Exit Function
'hell:
'End Function

Private Function AddButton(caption As String, resImg As Long) As Office.CommandBarControl
    Dim cbMenu As CommandBarButton
    Dim orgData As String
    Dim ipict As IPictureDisp
    
    On Error Resume Next
    
1    If VBInstance.CommandBars.count = 0 Then VBInstance.CommandBars.Add temporary:=True

    orgData = Clipboard.GetText
    Clipboard.Clear
    
2    VBInstance.CommandBars(1).Visible = True
3    Set cbMenu = VBInstance.CommandBars(1).Controls.Add(1, , , , temporary:=True) ', , , VBInstance.CommandBars(2).Controls.Count)
4    cbMenu.caption = caption
5    Set ipict = LoadResPicture(resImg, 0)

     If ipict Is Nothing Then
6        MsgBox "Failed to load res picture: " & resImg
     Else
7       Clipboard.SetData ipict
8       cbMenu.PasteFace
     End If
9    Set AddButton = cbMenu
    
    Clipboard.Clear
    If Len(orgData) > 0 Then Clipboard.SetText orgData
    
    Exit Function
hell: 'this can barf with typename(cbmenu) = nothing
    MsgBox "FastBuild.AddButton: " & caption & " Err: " & Err.Description & " line: " & Erl & " " & TypeName(cbMenu)
    
End Function

'Private Function FindRunButton() As Office.CommandBarControl
'
'    Dim cbToolbar As Office.CommandBar
'    Dim cbSubMenu As Office.CommandBarControl
'
'    For Each cbToolbar In VBInstance.CommandBars
'        'Debug.Print "Toolbar: " & cbToolbar.Index
'        'If cbToolbar.Index = 17 Then Stop
'        For Each cbSubMenu In cbToolbar.Controls
'            'Debug.Print vbTab & cbSubMenu.caption
'            If cbSubMenu.caption = "&Start" Then
'                Set FindRunButton = cbSubMenu
'                Exit Function
'            End If
'        Next
'    Next
'
'End Function

Private Function AddToAddInCommandBar(ByRef sCaption As String, Optional topMenuName As String = "Add-Ins") As Office.CommandBarControl
                 
    Dim cbMenuCommandBar As Office.CommandBarControl
    Dim cbMenu           As Object
  
    On Error GoTo error_handler
    
    Set cbMenu = VBInstance.CommandBars(topMenuName)
    
    If cbMenu Is Nothing Then
        Debug.Print "Could not find top menu " & topMenuName
        Exit Function
    End If
    
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.caption = sCaption
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
error_handler:
    
    Debug.Print "Connect::AddToAddInCommandBar"
    
End Function

Private Function AddrefMenu(caption As String, Optional menuName As String = "&Project", Optional afterItem = "Refere&nces...") As CommandBarControl

    Dim cbc As CommandBarControl
    'Dim cb As CommandBarButton
    Dim o As Object
    Dim i As Long, j As Long
    
    On Error GoTo hell
    'On Error Resume Next

    'this accounts for if the menuName is not present..
    For i = 1 To VBInstance.CommandBars(1).Controls.count
          If VBInstance.CommandBars(1).Controls(i).caption = menuName Then
                Exit For
          End If
    Next

    If i > VBInstance.CommandBars(1).Controls.count Then
        Debug.Print "AddrefMenu not found: " & menuName
        Exit Function
    End If
    
    
'    For i = 1 To VBInstance.CommandBars(1).Controls.Count
'          If VBInstance.CommandBars(1).Controls(i).caption = "Fast Build2" Then
'                Set cbc = VBInstance.CommandBars(1).Controls(i)
'                Exit For
'          Else
'            Debug.Print i & " " & VBInstance.CommandBars(1).Controls(i).caption
'          End If
'    Next
'
'    If cbc Is Nothing Then
'        Set cb = VBInstance.CommandBars(1)(.Add(1) 'typename commandbarbutton
'        cb.caption = "Fast Build2"
'    End If
'
'    Set o = cbc.Controls.Add()
    
    If Len(afterItem) > 0 Then
    
        For j = 1 To VBInstance.CommandBars(1).Controls(i).Controls.count
            If VBInstance.CommandBars(1).Controls(i).Controls(j).caption = afterItem Then Exit For
            'Debug.Print VBInstance.CommandBars(1).Controls(i).Controls(j).caption
        Next

        If j > VBInstance.CommandBars(1).Controls(i).Controls.count Then
            Debug.Print "AddrefMenu subitem not found: " & afterItem & " (adding to end)"
            'j = VBInstance.CommandBars(1).Controls(i).Controls.Count - 3
            ' Set AddrefMenu = VBInstance.CommandBars(1).Controls(i).Controls.Add()
            Exit Function
        Else
            Set AddrefMenu = VBInstance.CommandBars(1).Controls(i).Controls.Add(, , , j + 2, temporary:=True)   'add the menu before the References ... menu
        End If
        
    Else
        Dim cb As CommandBarPopup
        Set cb = VBInstance.CommandBars(1).Controls(i)
        Set AddrefMenu = cb.Controls.Add(, , , 1, temporary:=True)
    End If

    
    AddrefMenu.caption = caption

Exit Function

hell:
    MsgBox "FastBuild.AddrefMenu(" & caption & "): " & Err.Description
    
End Function

Private Function FindMakeMenu() As Office.CommandBarControl

    Dim cbSubMenu As Office.CommandBarControl
    Dim i As Long
    
    'On Error GoTo hell
    On Error Resume Next

    For Each cbSubMenu In VBInstance.CommandBars(1).Controls("File").Controls
        i = i + 1
        'Debug.Print cbSubMenu.caption
        If InStr(cbSubMenu.caption, "Ma&ke") > 0 Or cbSubMenu.caption = "Make..." Then
            Set FindMakeMenu = cbSubMenu
            Exit Function
        End If
    Next

Exit Function
hell:
    MsgBox "FastBuild.FindMakeMenu: " & Err.Description
    
End Function

Private Sub mnuImmediate_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    ClearImmediateWindow
    SendKeys "{F5}", True
End Sub

Sub ClearImmediateWindow()
    On Error Resume Next
    Dim oWindow As VBIDE.Window
    Set oWindow = VBInstance.ActiveWindow
    VBInstance.Windows("Immediate").SetFocus
    SendKeys "^{Home}", True     'win10 permission denied
    SendKeys "^+{End}", True
    SendKeys "{Del}", True
    oWindow.SetFocus
End Sub

'doesnt work in win10? worked in XP
Sub SetImmediateText(text As String)
    On Error Resume Next
    Dim oWindow As VBIDE.Window
    Dim saved As String
    Dim s As Date
    
    If Len(text) = 0 Then Exit Sub
    
    ClearImmediateWindow
    'saved = Clipboard.GetText
    Clipboard.Clear
    Clipboard.SetText text
    'MsgBox "saved: " & Clipboard.GetText
    
    Set oWindow = VBInstance.ActiveWindow
    VBInstance.Windows("Immediate").SetFocus
    SendKeys "^v", False 'True
   
    's = Now apparently win10 has some timing issues..we cant restore the old clipboard not gonna fight with it..
    'While DateDiff("s", s, Now) < 2
    '    DoEvents
    'Wend
    'Clipboard.Clear
    'If Len(saved) > 0 Then Clipboard.SetText saved
    'MsgBox text
End Sub


Private Sub mnuMemWindow_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    If FileExists(MemWindowExe) Then
        Shell MemWindowExe & " /pid:" & GetCurrentProcessId(), vbNormalFocus
    Else
        MsgBox "File not found: " & MemWindowExe
    End If
End Sub

'Private Sub mnuRealRun_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
'    If ClearImmediateOnStart = 1 Then
'        mnuImmediate_Click CommandBarControl, handled, CancelDefault
'    End If
'End Sub

Private Sub mnuOpenHomeDir_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim path As String
    path = VBInstance.ActiveVBProject.FileName
    path = GetParentFolder(path)
    Shell "explorer " & path, vbNormalFocus
End Sub

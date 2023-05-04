VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9825
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   19530
   _ExtentX        =   34449
   _ExtentY        =   17330
   _Version        =   393216
   Description     =   "Big Project Manager"
   DisplayName     =   "Big Project Manager"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 98 (ver 6.0)"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'cribbed a lot of notes from ColinE66 custom drawn project explorer
'   https://www.vbforums.com/showthread.php?890617-Add-In-Large-Project-Organiser-(alternative-Project-Explorer)-No-sub-classing!&highlight=
    
Private mcbMenuCommandBar     As Office.CommandBarControl
Public WithEvents MenuHandler As CommandBarEvents
Attribute MenuHandler.VB_VarHelpID = -1
Private WithEvents mFileEvents As FileControlEvents
Attribute mFileEvents.VB_VarHelpID = -1
Private WithEvents mProjectEvents As VBProjectsEvents
Attribute mProjectEvents.VB_VarHelpID = -1
Private WithEvents mComponentEvents As VBComponentsEvents
Attribute mComponentEvents.VB_VarHelpID = -1

Const PROJEXGUID As String = "{D7123946-C446-4CCC-A97D-5ABBF19147FF}"

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private multiProjWarned As Boolean

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)
    On Error Resume Next
    
    If Not mToolWindow Is Nothing Then
        VBInstance.ActiveVBProject.WriteProperty "BigProjectManager", "showAtStartup", mToolWindow.Visible
        If mToolWindow.Visible Then mUserDoc.SaveTreeToFile
    End If
    
End Sub

Private Sub AutoShowForProjectIfRequired()
    On Error Resume Next
    Dim showIt As Boolean
    showIt = VBInstance.ActiveVBProject.ReadProperty("BigProjectManager", "showAtStartup") 'can throw errors...
    If showIt Then Me.Show
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    On Error Resume Next
    mProjectLoaded = True
    mainProjPath = VBInstance.ActiveVBProject.filename
    mUserDoc.ipc_Message "AddinInstance_OnStartupComplete:" & mainProjPath
    AutoShowForProjectIfRequired
End Sub

Function warnMultipleProjects() As Boolean

     If VBInstance.VBProjects.Count > 1 Then
        If Not multiProjWarned Then
            MsgBox "Tree Explorer is only designed to work with a single project for now, groups not yet supported.", vbInformation
        End If
        multiProjWarned = True
        warnMultipleProjects = True
    End If
   
End Function

Sub Hide()
   On Error Resume Next
   mToolWindow.Visible = False
   'FormDisplayed = False
End Sub

Sub Show()
   On Error GoTo EH
   
   'If FormDisplayed Then Exit Sub
   If warnMultipleProjects() Then Exit Sub
  
   mUserDoc.cmdStartup_Click
   mToolWindow.Visible = True
   'FormDisplayed = True
   Exit Sub
EH:
   MsgBox "Show " & Err.Description
End Sub

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub

Private Sub mComponentEvents_ItemAdded(ByVal c As VBIDE.VBComponent)

    If Not mProjectLoaded Then Exit Sub
    If warnMultipleProjects() Then Exit Sub
    
    'If c.Type = vbext_ct_RelatedDocument Then Stop
    
    mUserDoc.ipc_Message "Component|Added|" & c.Collection.Parent.name & "|" & c.Type & "|" & c.name & "|" & c.FileNames(1)
    
End Sub

Private Sub mComponentEvents_ItemRemoved(ByVal c As VBIDE.VBComponent)

    If Not mProjectLoaded Then Exit Sub
    If warnMultipleProjects() Then Exit Sub
    
    mUserDoc.ipc_Message "Component|Removed|" & c.Collection.Parent.name & "|" & c.Type & "|" & c.name & "|" & c.FileNames(1)

End Sub

Private Sub mComponentEvents_ItemRenamed(ByVal c As VBIDE.VBComponent, ByVal OldName As String)

    If Not mProjectLoaded Then Exit Sub
    If warnMultipleProjects() Then Exit Sub
    
    mUserDoc.ipc_Message "Component|Renamed|" & c.Collection.Parent.name & "|" & c.Type & "|" & c.name & "|" & c.FileNames(1) & "|" & OldName

End Sub

Private Sub mComponentEvents_ItemSelected(ByVal c As VBIDE.VBComponent)

    If Not mProjectLoaded Then Exit Sub
    If warnMultipleProjects() Then Exit Sub
    
    mUserDoc.ipc_Message "Component|Selected|" & c.Collection.Parent.name & "|" & c.Type & "|" & c.name & "|" & c.FileNames(1)

End Sub


''these are not triggered just from doing Add Form or remove Form..
'Private Sub mFileEvents_AfterAddFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal filename As String)
'    'ipc.Send "FileEvents_AddFile: " & FileName
'End Sub
'
''they did a save as, changing form name in properties does not trigger this...
'Private Sub mFileEvents_AfterChangeFileName(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal newName As String, ByVal OldName As String)
'    'ipc.Send "FileEvents_ChangeFileName: " & OldName & "|" & NewName
'End Sub
'
'Private Sub mFileEvents_AfterRemoveFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal filename As String)
'   ' ipc.Send "FileEvents_RemoveFile: " & FileName
'End Sub

Private Sub mProjectEvents_ItemAdded(ByVal VBProject As VBIDE.VBProject) '... and this marks the end of a project loading
    If mToolWindow.Visible Then
        If warnMultipleProjects() Then mToolWindow.Visible = False
    End If
End Sub

'Private Sub mProjectEvents_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
'
'End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    
    On Error GoTo error_handler
    
    Dim hwnd As Long
     
 
    'If ConnectMode = ext_cm_Startup Then 'does not trigger if loaded manually after IDE startup
         
    If VBInstance Is Nothing Then
        Set VBInstance = Application
  
        If mToolWindow Is Nothing Then
           Set mToolWindow = VBInstance.Windows.CreateToolWindow(AddInInst, "BigProjectManager.UserDoc", "Big Project Manager", PROJEXGUID, mUserDoc)
        End If
 
        Set mcbMenuCommandBar = AddToAddInCommandBar("Big Project Manager")
        Set Me.MenuHandler = VBInstance.Events.CommandBarEvents(mcbMenuCommandBar)

        hwnd = FindWindowEx(VBInstance.MainWindow.hwnd, 0, "PROJECT", vbNullString)
        mProjectTreeHwnd = FindWindowEx(hwnd, 0, "SysTreeView32", vbNullString)
        
        Set mFileEvents = VBInstance.Events.FileControlEvents(Nothing)
        Set mComponentEvents = VBInstance.Events.VBComponentsEvents(Nothing)
        Set mProjectEvents = VBInstance.Events.VBProjectsEvents
        
     End If
  
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub



Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
      
    unloading = True
    mcbMenuCommandBar.Delete
  
    Set mToolWindow = Nothing
    Set VBInstance = Nothing
    Set mFileEvents = Nothing
    Set mComponentEvents = Nothing
    Set mProjectEvents = Nothing
    
End Sub



Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    Set cbMenu = VBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        Exit Function
    End If
    
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function





VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   12345
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   19155
   _ExtentX        =   33787
   _ExtentY        =   21775
   _Version        =   393216
   Description     =   "Alternative to the built-in Project Explorer"
   DisplayName     =   "Project Explorer II"
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
Attribute VB_Description = "Project Explorer II"
Option Explicit

Private mcbMenuCommandBar     As Office.CommandBarControl
Private mToolWindow           As VBIDE.Window

Const PROJEXGUID As String = "{D7123946-C446-4CCC-A97D-5ABBF19147B2}"
Private mUserDoc As UserDoc
Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1
Sub Hide()
   On Error Resume Next
   mToolWindow.Visible = False
   FormDisplayed = False
End Sub
Sub Show()
   On Error GoTo EH
   
   If FormDisplayed Then Exit Sub
   
   mUserDoc.Init
   mToolWindow.Visible = True
   FormDisplayed = True
   Exit Sub
EH:
   MsgBox "Show " & Err.Description
End Sub
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
   On Error GoTo error_handler
    
   Set gVBInstance = Application
   If mToolWindow Is Nothing Then
      Set mToolWindow = gVBInstance.Windows.CreateToolWindow(AddInInst, "ProjEx.UserDoc", "Project Explorer", PROJEXGUID, mUserDoc)
   End If
    
   If ConnectMode = ext_cm_External Then
      Me.Show
   Else 'ext_cm_Startup=1 ; ext_cm_AfterStartup=0
      Set mcbMenuCommandBar = AddToAddInCommandBar("Project Explorer II")
      Set Me.MenuHandler = gVBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
      Me.Show
  End If

   If ConnectMode = ext_cm_AfterStartup Then
      If GetSetting(App.Title, "Settings", "DisplayOnConnect", "1") = "1" Then
         Me.Show
      End If
   End If
  
   Exit Sub
    
error_handler:
    
    MsgBox "AddinInstance_OnConnection " & Err.Description
    
End Sub
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
   On Error Resume Next
    
   mcbMenuCommandBar.Delete
   Set mToolWindow = Nothing
   Set gVBInstance = Nothing
    
   SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
   FormDisplayed = False
    
End Sub
Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
   If GetSetting(App.Title, "Settings", "DisplayOnConnect", "1") = "1" Then
      Me.Show
   End If
End Sub
Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    If Not FormDisplayed Then Me.Show
End Sub
Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
    Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
    Dim cbMenu As Object
  
    On Error GoTo AddToAddInCommandBarErr
    
    Set cbMenu = gVBInstance.CommandBars("Add-Ins")
    If cbMenu Is Nothing Then
        Exit Function
    End If
    
    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
    cbMenuCommandBar.Caption = sCaption
    
    Set AddToAddInCommandBar = cbMenuCommandBar
    
    Exit Function
    
AddToAddInCommandBarErr:

End Function


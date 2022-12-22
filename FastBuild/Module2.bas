Attribute VB_Name = "Module2"
Public LastCommandOutput As String
Public VBInstance As VBIDE.VBE
'Public Connect As Connect
Public ClearImmediateOnStart As Long
Public ShowPostBuildOutput As Long

Public dbgIntercept As New CDebugIntercept

Public MemWindowExe As String
Public CodeDBExe As String
Public APIAddInExe As String
Public ExternalDebugWindow As String

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

    Private Const OFN_ALLOWMULTISELECT = &H200
    Private Const OFN_CREATEPROMPT = &H2000
    Private Const OFN_ENABLEHOOK = &H20
    Private Const OFN_ENABLETEMPLATE = &H40
    Private Const OFN_ENABLETEMPLATEHANDLE = &H80
    Private Const OFN_EXPLORER = &H80000                         '  new look commdlg
    Private Const OFN_EXTENSIONDIFFERENT = &H400
    Private Const OFN_FILEMUSTEXIST = &H1000
    Private Const OFN_HIDEREADONLY = &H4
    Private Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules
    Private Const OFN_NOCHANGEDIR = &H8
    Private Const OFN_NODEREFERENCELINKS = &H100000
    Private Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
    Private Const OFN_NONETWORKBUTTON = &H20000
    Private Const OFN_NOREADONLYRETURN = &H8000
    Private Const OFN_NOTESTFILECREATE = &H10000
    Private Const OFN_NOVALIDATE = &H100
    Private Const OFN_OVERWRITEPROMPT = &H2
    Private Const OFN_PATHMUSTEXIST = &H800
    Private Const OFN_READONLY = &H1
    Private Const OFN_SHAREAWARE = &H4000
    Private Const OFN_SHAREFALLTHROUGH = 2
    Private Const OFN_SHARENOWARN = 1
    Private Const OFN_SHAREWARN = 0
    Private Const OFN_SHOWHELP = &H10
     
    Private Type OPENFILENAME
            lStructSize As Long
            hwndOwner As Long
            hInstance As Long
            lpstrFilter As String
            lpstrCustomFilter As String
            nMaxCustFilter As Long
            nFilterIndex As Long
            lpstrFile As String
            nMaxFile As Long
            lpstrFileTitle As String
            nMaxFileTitle As Long
            lpstrInitialDir As String
            lpstrTitle As String
            flags As Long
            nFileOffset As Integer
            nFileExtension As Integer
            lpstrDefExt As String
            lCustData As Long
            lpfnHook As Long
            lpTemplateName As String
    End Type
     
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Function LoadHexToolTipsDll() As Boolean

    Dim h As Long
    Const dll = "hexTooltip.dll"
    
    If GetModuleHandle(dll) = 0 Then
        h = LoadLibrary(dll)
        If h = 0 Then h = LoadLibrary(App.path & "\" & dll)
        If h = 0 Then Exit Function
    End If

    LoadHexToolTipsDll = True
    
End Function

Function IPCCommand(msg As String)
    On Error GoTo hell
    
    Dim cmd As String
    Dim a As Long
    
    'MsgBox "in ipc command!"
    
    a = InStr(msg, ":")
    If a < 1 Then Exit Function
    
    cmd = LCase(Mid(msg, 1, a))
    msg = Mid(msg, a + 1)
    
    If cmd = "add:" Then
        If FileExists(msg) Then
            VBInstance.ActiveVBProject.VBComponents.AddFile msg
        Else
            MsgBox "Can not add file, File not found: " & msg, vbInformation
        End If
    End If
    
    'this does not work when IDE is at runtime or startup :_(yet:)
    If cmd = "cls:" Then
        ClearImmediateWindow
        MsgBox "cleared!" & Err.Description
    End If
        
    
   Exit Function
hell:
    MsgBox "Error in IpcCommand: " & Err.Description, vbInformation
End Function

Sub ClearImmediateWindow()
    On Error Resume Next
    Dim oWindow As VBIDE.Window
1    Set oWindow = VBInstance.ActiveWindow
2    VBInstance.Windows("Immediate").SetFocus
3    SendKeys "^{Home}", True
4    SendKeys "^+{End}", True
5    SendKeys "{Del}", True
6    If Not oWindow Is Nothing Then oWindow.SetFocus
    
     If Err.Number <> 0 Then MsgBox Erl & " " & Err.Description
End Sub

Function ShowOpenMultiSelect(Optional hwnd As Long) As String()
    Dim tOPENFILENAME As OPENFILENAME
    Dim lResult As Long
    Dim vFiles As Variant
    Dim lIndex As Long, lStart As Long
    Dim ret() As String, pd As String
    
    
    With tOPENFILENAME
        .flags = OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_FILEMUSTEXIST Or OFN_LONGNAMES
        .hwndOwner = hwnd
        .nMaxFile = 2048
        .lpstrFilter = "All Files" & Chr(0) & "*.*" & Chr(0) & Chr(0)
        .lpstrFile = Space(.nMaxFile - 1) & Chr(0)
        .lStructSize = Len(tOPENFILENAME)
    End With
    
    lResult = GetOpenFileName(tOPENFILENAME)
    
    If lResult > 0 Then
        With tOPENFILENAME
            vFiles = Split(Left(.lpstrFile, InStr(.lpstrFile, Chr(0) & Chr(0)) - 1), Chr(0))
        End With
        
        If UBound(vFiles) = 0 Then
            push ret, vFiles(0)
        Else
            pd = vFiles(0)
            If Right$(pd, 1) <> "\" Then pd = pd & "\"
            For lIndex = 1 To UBound(vFiles)
                push ret, pd & vFiles(lIndex)
            Next
        End If
    End If
    
    ShowOpenMultiSelect = ret()
    
End Function
     
'Private Function AddBS(ByVal sPath As String) As String
'    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
'    AddBS = sPath
'End Function

 

Function ExpandVars(ByVal cmd As String, exeFullPath As String) As String
    Dim appDir As String
    Dim fName As String
    
    cmd = Trim(cmd)
    appDir = GetParentFolder(exeFullPath)
    fName = FileNameFromPath(exeFullPath)
    
    ExpandVars = Replace(cmd, "%1", exeFullPath)
    ExpandVars = Replace(ExpandVars, "%app", appDir, , , vbTextCompare)
    ExpandVars = Replace(ExpandVars, "%ap%", appDir, , , vbTextCompare)
    ExpandVars = Replace(ExpandVars, "%fname", fName, , , vbTextCompare)
    
End Function

Function isBuildPathSet() As Boolean

    On Error Resume Next
    Dim fastBuildPath As String
    
    If VBInstance.ActiveVBProject Is Nothing Then Exit Function
    fastBuildPath = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "fullPath")
    If Len(fastBuildPath) = 0 Then Exit Function
    isBuildPathSet = True
    
End Function

'set the current directory to be parent folder as vbp folder path...
Sub SetHomeDir()
    On Error Resume Next
    Dim homeDir As String
    homeDir = VBInstance.ActiveVBProject.FileName 'path to vbp file
    homeDir = GetParentFolder(homeDir)
    If Len(homeDir) > 0 Then ChDir homeDir
End Sub

Function GetPostBuildCommand() As String
    On Error Resume Next
    If VBInstance.ActiveVBProject Is Nothing Then Exit Function
    GetPostBuildCommand = VBInstance.ActiveVBProject.ReadProperty("fastBuild", "PostBuild")
End Function

Function ConsoleAppCommand() As String
    On Error Resume Next
    Dim i As Long, exe As String
    
    If VBInstance.ActiveVBProject Is Nothing Then Exit Function
    
    i = CInt(VBInstance.ActiveVBProject.ReadProperty("fastBuild", "IsConsoleApp"))
    
    If i <> 0 Then
        exe = GetVB6Path()
        
        If FileExists(exe & "\vblink.exe") Then
            exe = exe & "\vblink.exe" 'if link tool addin is in use this is real vb linker
        Else
            exe = exe & "\link.exe"
        End If
        
        If FileExists(exe) Then
            exe = GetShortName(exe)
            ConsoleAppCommand = exe & " /EDIT /SUBSYSTEM:CONSOLE %1"
        End If
    End If
    
End Function

Public Function GetVB6Path() As String
     Dim h As Long, ret As String
     ret = Space(500)
     h = GetModuleHandle("vb6.exe")
     h = GetModuleFileName(h, ret, 500)
     If h > 0 Then ret = Mid(ret, 1, h)
     GetVB6Path = Replace(ret, "vb6.exe", Empty, , , vbTextCompare)
End Function

'file must exist for this to work which is stupid...
Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 500
    Dim lResult As Long
    
    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)
    GetShortName = Replace(GetShortName, Chr(0), Empty)
    
    'if the api fails, we will revert to a quoted version of the full file name
    '(maybe file doesnt exist, or buf to small)
    If Len(GetShortName) = 0 Then
        GetShortName = """" & sFile & """"
    End If

End Function

Function IsIde() As Boolean
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  If Err.Number <> 0 Then FileExists = False
End Function

Function FolderExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbDirectory) <> "" Then FolderExists = True
  If Err.Number <> 0 Then FolderExists = False
End Function

Function GetParentFolder(path) As String
    On Error Resume Next
    Dim tmp() As String
    Dim ub As String
    If Len(path) = 0 Then Exit Function
    If InStr(path, "\") < 1 Then Exit Function
    If Right(path, 1) = "\" Then path = Mid(path, 1, Len(path) - 1)
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
    If Err.Number <> 0 Then GetParentFolder = Empty
End Function

Function FileNameFromPath(fullpath) As String
    If InStr(fullpath, "\") > 0 Then
        tmp = Split(fullpath, "\")
        FileNameFromPath = CStr(tmp(UBound(tmp)))
    Else
        FileNameFromPath = fullpath
    End If
End Function

Function GetFileReport(fpath As String) As String
    On Error Resume Next
    
    Dim MyStamp As Date
    Dim ret() As String
    
    If Not FileExists(fpath) Then
        GetFileReport = "Build Failed: " & fpath
        Exit Function
    End If
    
    MyStamp = FileDateTime(fpath)
    
    push ret, "Output File: " & fpath & "  (" & FileSize(fpath) & ")"
    push ret, "Last Modified: " & Format(MyStamp, "dddd, mmmm dd, yyyy - h:mm:ss AM/PM")
    
    GetFileReport = Join(ret, vbCrLf)

End Function

Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo Init
    X = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
Init: ReDim ary(0): ary(0) = Value
End Sub

Public Function FileSize(fpath As String) As String
    Dim fsize As Long
    Dim szName As String
    On Error GoTo hell
    
    fsize = FileLen(fpath)
    
    szName = " bytes"
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Kb"
    End If
    
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Mb"
    End If
    
    FileSize = fsize & szName
    
    Exit Function
hell:
    
End Function



Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function


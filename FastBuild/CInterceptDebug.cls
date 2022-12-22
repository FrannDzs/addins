VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDebugIntercept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'this could still use a little more cleanup but you get he idea

Implements IVBPrint

'https://www.vbforums.com/showthread.php?898185-Redirect-where-Debug-Print-goes&p=5586322#post5586322
'https://www.vbforums.com/showthread.php?891891-(VB6)-Implement-the-Print-method-on-any-object&highlight=
'IDL: https://www.vbforums.com/showthread.php?858545-(VB6)-Replace-VB-s-Circle-method-with-API-s

'persistant debug.print by elroy:
'      http://www.vbforums.com/showthread.php?874127-Persistent-Debug-Print-Window

' //
' // Debug redirect
' // by The trick
' //
 
Dim pOriginal As PTR
Private isEnabled As Boolean

Private Enum PTR
    [_]
End Enum

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As PTR) As PTR
Private Declare Sub GetMem4 Lib "msvbvm60" (ByRef pAddr As Any, ByRef pDst As Any)
Private Declare Sub GetMemPtr Lib "msvbvm60" Alias "GetMem4" (ByRef pAddr As Any, ByRef pDst As Any)
Private Declare Sub GetMem8 Lib "msvbvm60" (ByRef pAddr As Any, ByRef pDst As Any)
Private Declare Sub PutMemPtr Lib "msvbvm60" Alias "PutMem4" (ByRef pDst As Any, ByVal pVal As PTR)
                    
'
' This is a Stand-Alone module that can be thrown into any project.
' It works in conjunction with the PersistentDebugPrint program, and that program must be running to use this module.
' The only procedure you should worry about is the DebugPrint procedure.
' Basically, it does what it says, provides a "Debug" window that is persistent across your development IDE exits and starts (even IDE crashes).
'
Private Type COPYDATASTRUCT
    dwData  As Long
    cbData  As Long
    lpData  As Long
End Type
'
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Bytes As Long)
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Dim mhWndTarget As Long
'
Const DoDebugPrint As Boolean = True
Const LANG_US = &H409

Property Get isActive() As Boolean
    isActive = isEnabled
End Property

Function Enable() As Boolean

    On Error GoTo hell
    
    'If Not IsIde() Then Exit Function as an addin we are always in the IDE
    
    If isEnabled Then
        Enable = True
        Exit Function
    End If
    
    pOriginal = ReplaceDebugObject(ObjPtr(Me)) 'can throw error
    Enable = True
    isEnabled = True
    
hell:
    If Err.Number <> 0 Then MsgBox Err.Description
    
End Function
 
Function Disable() As Boolean
    On Error GoTo hell
    If Not isEnabled Then Exit Function
    ReplaceDebugObject pOriginal 'can throw error
    isEnabled = False
    Disable = True
hell:
End Function

Private Property Let IVBPrint_Column(ByVal RHS As Long)
    MsgBox "Let IVBPrint_Column " & RHS
End Property

Private Property Get IVBPrint_Column() As Long
    MsgBox "Get IVBPrint_Column"
End Property

Private Sub IVBPrint_WriteText(ByVal strText As String)
    If Len(strText) = 0 Then Exit Sub
    If strText = vbCrLf Then Exit Sub
    debugPrint strText
End Sub


Private Function ReplaceDebugObject(ByVal pObj As PTR) As PTR

    Static s_pCurObject As PTR
    Dim hVBA        As PTR
    Dim pNTHdr      As PTR
    Dim pStart      As PTR
    Dim pEnd        As PTR
    Dim cSign       As Currency
    Dim lLength     As Long
    Dim lOldProtect As Long
    
    If s_pCurObject = 0 Then
    
        hVBA = GetModuleHandle(StrPtr("vba6"))
        If hVBA = 0 Then Exit Function
    
        GetMem4 ByVal hVBA + &H3C, pNTHdr
        pNTHdr = pNTHdr + hVBA
        
        GetMem4 ByVal pNTHdr + &H12C, pStart
        pStart = pStart + hVBA
        
        GetMem4 ByVal pNTHdr + &H128, lLength
        pEnd = pStart + lLength - 8
        
        Do While pStart <= pEnd
            
            GetMem8 ByVal pStart, cSign
            
            If cSign = 511398171365990.4051@ Then
            
                GetMemPtr ByVal pStart + &H11, pStart
                GetMemPtr ByVal pStart + &H44, pStart
                GetMemPtr ByVal pStart + &H1, s_pCurObject
                Exit Do
                
            End If
            
            pStart = pStart + 1
            
        Loop

    End If
    
    If s_pCurObject = 0 Then
        Err.Raise 51
    End If
    
    GetMemPtr ByVal s_pCurObject, ReplaceDebugObject
    PutMemPtr ByVal s_pCurObject, pObj
    
End Function
 


'Function IsIde() As Boolean
'    On Error GoTo out
'    Debug.Print 1 / 0
'out: IsIde = Err
'End Function

'------------------------------------------------------------



Public Sub debugPrintf(ByVal msg As String, ParamArray values() As Variant)

    Dim i As Long, tmp(), result As String
    
    'paramArray to variant array so we can pass it to printf
    For i = 0 To UBound(values)
        If IsNull(values(i)) Then
            push tmp, "[Null]"
        ElseIf IsObject(values(i)) Then
            push tmp, "[Object:" & TypeName(values(i)) & "]"
        Else
            push tmp, values(i)
        End If
    Next
    
    result = printf(msg, tmp)
    DebugPrint_Internal result
    
End Sub

Public Sub debugDiv()
    DebugPrint_Internal "<div>"
End Sub

Public Sub debugClear()
    DebugPrint_Internal "<cls>"
End Sub

Sub debugPrint(ParamArray vArgs() As Variant)

    Dim v       As Variant
    Dim sMsg    As String
    Dim bNext   As Boolean
    
    For Each v In vArgs
        If bNext Then
            sMsg = sMsg & Space$(8&)
            sMsg = Left$(sMsg, (Len(sMsg) \ 8&) * 8&)
        End If
        bNext = True
        sMsg = sMsg & CStr(v)
    Next
    
    DebugPrint_Internal sMsg
    
End Sub

Private Function canStartServer() As Boolean
    Dim pth As String
    pth = GetSetting("dbgWindow", "settings", "path", "")
    If Not FileExists(pth) Then Exit Function
    Shell pth, vbNormalFocus
    canStartServer = (Err.Number = 0)
    Sleep 400
    FindServerWindow
End Function

Private Sub DebugPrint_Internal(sMsg As String)
    ' Commas are allowed, but not semicolons.
    '
    If Not DoDebugPrint Then Exit Sub
    '
    Static bErrorMessageShown As Boolean
    
    FindServerWindow
    
    If mhWndTarget = 0& Then
        If Not bErrorMessageShown Then
            If Not canStartServer() Then
                MsgBox "The Persistent Debug Print Window could not be found. I can auto start it, but you havent run it yet for it to save its path to the registry.", vbCritical, "Persistent Debug Message"
                bErrorMessageShown = True
                Exit Sub
            End If
        End If
    End If
   
    SendStringToAnotherWindow sMsg
End Sub

Private Function FindServerWindow(Optional forceRefresh As Boolean = False) As Boolean
    
    Const vbIDEClassName = "ThunderFormDC"
    Const vbEXEClassName = "ThunderRT6FormDC"
    Const vbExeClass2 = "ThunderRT6Form"
    Const vbWindowCaption = "Persistent Debug Print Window"

    If forceRefresh Then mhWndTarget = 0

    If IsWindow(mhWndTarget) Then
        FindServerWindow = True
        Exit Function
    End If
    
    mhWndTarget = FindWindow(vbIDEClassName, vbWindowCaption)

    If mhWndTarget = 0 Then
        mhWndTarget = FindWindow(vbEXEClassName, vbWindowCaption)
    End If
    
    If mhWndTarget = 0 Then
        mhWndTarget = FindWindow(vbExeClass2, vbWindowCaption)
    End If
    
    If IsWindow(mhWndTarget) = 0 Then Exit Function
    
    FindServerWindow = True
     
End Function

'Private Sub ValidateTargetHwnd()
'    If IsWindow(mhWndTarget) Then
'        Select Case WindowClass(mhWndTarget)
'        Case "ThunderForm", "ThunderRT6Form"
'            If WindowText(mhWndTarget) = "Persistent Debug Print Window" Then
'                Exit Sub
'            End If
'        End Select
'    End If
'    EnumWindows AddressOf EnumToFindTargetHwnd, 0&
'End Sub
'
'Private Function EnumToFindTargetHwnd(ByVal hwnd As Long, ByVal lParam As Long) As Long
'    mhWndTarget = 0&                        ' We just set it every time to keep from needing to think about it before this is called.
'    Select Case WindowClass(hwnd)
'    Case "ThunderForm", "ThunderRT6Form"
'        If WindowText(hwnd) = "Persistent Debug Print Window" Then
'            mhWndTarget = hwnd
'            Exit Function
'        End If
'    End Select
'    EnumToFindTargetHwnd = 1&               ' Keep looking.
'End Function

Private Function WindowClass(hwnd As Long) As String
    WindowClass = String$(1024&, vbNullChar)
    WindowClass = Left$(WindowClass, GetClassName(hwnd, WindowClass, 1024&))
End Function

Private Function WindowText(hwnd As Long) As String
    ' Form or control.
    WindowText = String$(GetWindowTextLength(hwnd) + 1&, vbNullChar)
    Call GetWindowText(hwnd, WindowText, Len(WindowText))
    WindowText = Left$(WindowText, InStr(WindowText, vbNullChar) - 1&)
End Function

Private Sub SendStringToAnotherWindow(sMsg As String)
    Dim cds             As COPYDATASTRUCT
    Dim lpdwResult      As Long
    Dim Buf()           As Byte
    Const WM_COPYDATA   As Long = &H4A&
    '
    ReDim Buf(1 To Len(sMsg) + 1&)
    Call CopyMemory(Buf(1&), ByVal sMsg, Len(sMsg)) ' Copy the string into a byte array, converting it to ASCII.
    cds.dwData = 3&
    cds.cbData = Len(sMsg) + 1&
    cds.lpData = VarPtr(Buf(1&))
    Call SendMessage(mhWndTarget, WM_COPYDATA, 0, cds)
    'SendMessageTimeout mhWndTarget, WM_COPYDATA, 0&, cds, 0&, 1000&, lpdwResult ' Return after a second even if receiver didn't acknowledge.
End Sub


'------------------ dzzie basic printf implementation free for any use ----------------------
'implements:
'    \t -> tab
'    \n -> vbcrlf
'    %% -> %
'    %x = hex
'    %X = UCase(Hex(var))
'    %s = string
'    %S = UCase string
'    %c = Chr(var)
'    %d = numeric
Private Function printf(ByVal msg As String, vars() As Variant) As String

    Dim t
    Dim ret As String
    Dim i As Long, base, marker
    
    msg = Replace(msg, Chr(0), Empty)
    msg = Replace(msg, "\t", vbTab)
    msg = Replace(msg, "\n", vbCrLf) 'simplified
    msg = Replace(msg, "%%", Chr(0))
    
    t = Split(msg, "%")
    If UBound(t) <> UBound(vars) + 1 Then
        MsgBox "Format string mismatch.."
        Exit Function
    End If
    
    ret = t(0)
    For i = 1 To UBound(t)
        base = t(i)
        marker = ExtractSpecifier(base)
        If Len(marker) > 0 Then
            ret = ret & HandleMarker(base, marker, vars(i - 1))
        Else
            ret = ret & base
        End If
    Next
    
    ret = Replace(ret, Chr(0), "%")
    printf = ret
    
End Function

Private Function HandleMarker(base, ByVal marker, var) As String
    Dim newBase As String
    Dim mType As Integer
    Dim nVal As String
    Dim spacer As String
    Dim prefix As String
    Dim count As Long
    Dim leftJustify As Boolean
    
    If Len(base) > Len(marker) Then
        newBase = Mid(base, Len(marker) + 1) 'remove the marker..
    End If
    
    mType = Asc(Mid(marker, Len(marker), 1))  'last character
    
    Select Case mType
        Case Asc("x"): nVal = Hex(var)
        Case Asc("X"): nVal = UCase(Hex(var))
        Case Asc("s"): nVal = var
        Case Asc("S"): nVal = UCase(var)
        Case Asc("c"): nVal = Chr(var)
        Case Asc("d"): nVal = var
        
        Case Else: nVal = var
    End Select
    
    If Len(marker) > 1 Then 'it has some more formatting involved..
        marker = Mid(marker, 1, Len(marker) - 1) 'trim off type
        If Left(marker, 1) = "-" Then
            leftJustify = True
            marker = Mid(marker, 2)  'trim off left justify marker
        End If
        If Left(marker, 1) = "0" Then
            spacer = "0"
            marker = Mid(marker, 2)
        Else
            spacer = " "
        End If
        count = CLng(marker) - Len(nVal)
        If count > 0 Then prefix = String(count, spacer)
    End If
    
    If leftJustify Then
        HandleMarker = nVal & prefix & newBase
    Else
        HandleMarker = prefix & nVal & newBase
    End If
    
End Function

Private Function ExtractSpecifier(v)
    
    Dim ret As String
    Dim b() As Byte
    Dim i As Long
    If Len(v) = 0 Then Exit Function
    
    b() = StrConv(v, vbFromUnicode, LANG_US)
    
    For i = 0 To UBound(b)
        ret = ret & Chr(b(i))
        If b(i) = Asc("x") Then Exit For
        If b(i) = Asc("X") Then Exit For
        If b(i) = Asc("c") Then Exit For
        If b(i) = Asc("s") Then Exit For
        If b(i) = Asc("S") Then Exit For
        If b(i) = Asc("d") Then Exit For
    Next
    
    ExtractSpecifier = ret
    
End Function

Private Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo Init
    Dim X
       
    X = UBound(ary)
    ReDim Preserve ary(X + 1)
    
    If IsObject(Value) Then
        Set ary(X + 1) = Value
    Else
        ary(X + 1) = Value
    End If
    
    Exit Sub
Init:
    ReDim ary(0)
    If IsObject(Value) Then
        Set ary(0) = Value
    Else
        ary(0) = Value
    End If
End Sub

Private Function FileExists(path) As Boolean
  On Error GoTo hell
    
  '.(0), ..(0) etc cause dir to read it as cwd!
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If InStr(path, Chr(0)) > 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

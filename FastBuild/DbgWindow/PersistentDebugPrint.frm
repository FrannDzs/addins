VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDebugPrint 
   Caption         =   "Persistent Debug Print Window"
   ClientHeight    =   5295
   ClientLeft      =   1005
   ClientTop       =   3060
   ClientWidth     =   7635
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "PersistentDebugPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7635
   Begin VB.Frame fraSearch 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   4740
      Width           =   7395
      Begin VB.TextBox txtFilter 
         Height          =   330
         Left            =   960
         TabIndex        =   4
         Top             =   60
         Width           =   6315
      End
      Begin VB.Label Label1 
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   60
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H0000FFFF&
      Caption         =   "Continue"
      Height          =   495
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   180
      Width           =   1515
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   3555
      Top             =   2115
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   4545
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   90
      Width           =   7530
   End
   Begin VB.Menu mnuClear 
      Caption         =   "Clear"
   End
   Begin VB.Menu mnuCopy 
      Caption         =   "Copy"
   End
   Begin VB.Menu mnuSeparate 
      Caption         =   "Separate"
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Begin VB.Menu mnuTopMost 
         Caption         =   "TopMost"
      End
      Begin VB.Menu mnuTimeStamp 
         Caption         =   "Timestamp"
      End
      Begin VB.Menu mnuBackColor 
         Caption         =   "BackColor"
      End
      Begin VB.Menu mnuForeColor 
         Caption         =   "ForeColor"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Font"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuOpenHomePage 
         Caption         =   "Homepage"
      End
   End
End
Attribute VB_Name = "frmDebugPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long ' This is +1 (right - left = width)
    Bottom As Long ' This is +1 (bottom - top = height)
End Type
Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type
'
Private Const WM_SETREDRAW      As Long = &HB&
Private Const EM_SETSEL         As Long = &HB1&
Private Const EM_REPLACESEL     As Long = &HC2&
'
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function ShellExecuteA Lib "shell32.dll" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Integer) As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -1

Private Declare Function ChangeWindowMessageFilter Lib "user32" (ByVal msg As Long, ByVal flag As Long) As Long 'Vista+
Const WM_COPYDATA = &H4A
Const WM_COPYGLOBALDATA = &H49

Dim msgs() As String

'Private WithEvents subclass As spSubClass.clsSubClass

'Private Sub subclass_MessageReceived(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
'    If wMsg <> &H4A Then Exit Sub 'copy data
'
'    Dim CopyData As COPYDATASTRUCT
'    Dim Buffer(1 To 2048) As Byte
'    Dim temp As String
'
'    On Error Resume Next
'
'    CopyMemory CopyData, ByVal lParam, Len(CopyData)
'
'    If CopyData.dwFlag = 3 Then
'        CopyMemory Buffer(1), ByVal CopyData.lpData, CopyData.cbSize
'        temp = StrConv(Buffer, vbUnicode)
'        temp = Left$(temp, InStr(1, temp, Chr$(0)) - 1)
'        temp = Trim(temp)
'        frmDebugPrint.Out temp
'    End If
'
'End Sub

Public Function AllowCopyDataAcrossUIPI()
    Dim a, b, c
    Const MSGFLT_ADD = 1
    'a = ChangeWindowMessageFilter(WM_DROPFILES, MSGFLT_ADD)
    b = ChangeWindowMessageFilter(WM_COPYDATA, MSGFLT_ADD) 'we still need this for IPC to get hook data...
    c = ChangeWindowMessageFilter(WM_COPYGLOBALDATA, MSGFLT_ADD)
    'MsgBox a & " " & b & " " & c
End Function

Sub SetWindowTopMost(f As Form, Optional onTop As Boolean = True)
   SetWindowPos f.hWnd, IIf(onTop, HWND_TOPMOST, -2), f.Left / 15, _
        f.Top / 15, f.Width / 15, _
        f.Height / 15, Empty
End Sub

Private Sub cmdContinue_Click()
    cmdContinue.Visible = False
End Sub

Private Sub Form_Load()
        
        On Error Resume Next
        
        If App.PrevInstance Then End
        
        cmdContinue.Visible = False
        SaveSetting "dbgWindow", "settings", "path", App.Path & "\PersistentDebugPrint.exe"
        SaveSetting "dbgWindow", "settings", "hwnd", Me.hWnd
        
        If GetSetting("dbgWindow", "settings", "topMost", 0) <> 0 Then
            mnuTopMost.Checked = True
            SetWindowTopMost Me
        End If
        
        Me.Left = GetSetting(App.Title, "Settings", "Left", 0&)
        Me.Top = GetSetting(App.Title, "Settings", "Top", 0&)
        Me.Width = GetSetting(App.Title, "Settings", "Width", 6600&)
        Me.Height = GetSetting(App.Title, "Settings", "Height", 6600&)
        
        If Not FormIsFullyOnMonitor(Me) Then
            Me.Left = 0&
            Me.Top = 0&
        End If
        '
        txt.FontName = GetSetting(App.Title, "Settings", "FontName", "Fixedsys")
        txt.FontBold = GetSetting(App.Title, "Settings", "FontBold", False)
        txt.FontItalic = GetSetting(App.Title, "Settings", "FontItalic", False)
        txt.FontSize = GetSetting(App.Title, "Settings", "FontSize", 9)
        txt.FontStrikethru = GetSetting(App.Title, "Settings", "FontStrikethru", False)
        txt.FontUnderline = GetSetting(App.Title, "Settings", "FontUnderline", False)
        '
        txt.BackColor = GetSetting(App.Title, "Settings", "BackColor", vbWhite)
        txt.ForeColor = GetSetting(App.Title, "Settings", "ForeCOlor", vbBlack)
        mnuTimeStamp.Checked = GetSetting(App.Title, "Settings", "TimeStamp", 0)
    On Error GoTo 0
    
    'NOTE THESE MUST RUN AT SAME PRIVLEDGE LEVEL
    AllowCopyDataAcrossUIPI
    SubclassFormToReceiveStringMsg Me
    
    'Set subclass = New spSubClass.clsSubClass
    'subclass.AttachMessage Me.hwnd, &H4A
    
    If IsIde Then
        push msgs, "this is message 1"
        push msgs, "i like tacos"
        push msgs, "ack ack adak"
        txt = Join(msgs, vbCrLf)
    End If
    
End Sub


Function IsIde() As Boolean
    On Error GoTo out
    Debug.Print 1 / 0
out: IsIde = Err
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Settings", "Left", Me.Left
    SaveSetting App.Title, "Settings", "Top", Me.Top
    SaveSetting App.Title, "Settings", "Width", Me.Width
    SaveSetting App.Title, "Settings", "Height", Me.Height
    '
    SaveSetting App.Title, "Settings", "FontName", txt.FontName
    SaveSetting App.Title, "Settings", "FontBold", txt.FontBold
    SaveSetting App.Title, "Settings", "FontItalic", txt.FontItalic
    SaveSetting App.Title, "Settings", "FontSize", txt.FontSize
    SaveSetting App.Title, "Settings", "FontStrikethru", txt.FontStrikethru
    SaveSetting App.Title, "Settings", "FontUnderline", txt.FontUnderline
    '
    SaveSetting App.Title, "Settings", "BackColor", txt.BackColor
    SaveSetting App.Title, "Settings", "ForeCOlor", txt.ForeColor
    SaveSetting App.Title, "Settings", "TimeStamp", IIf(mnuTimeStamp.Checked, 1, 0)
    SaveSetting "dbgWindow", "settings", "topMost", IIf(mnuTopMost.Checked, 1, 0)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Not Me.WindowState = vbMinimized Then
        txt.Move 0&, 0&, Me.ScaleWidth, Me.ScaleHeight - fraSearch.Height - 200
        fraSearch.Top = Me.ScaleHeight - fraSearch.Height - 100
        fraSearch.Width = Me.ScaleWidth
        txtFilter.Width = Me.ScaleWidth - txtFilter.Left - 300
    End If
End Sub

Private Sub Label1_Click()
    Const help = "Supports multiple criteria as CSV\nIf first char is - then its all a subtractive filter"
    MsgBox Replace(help, "\n", vbCrLf), vbInformation
End Sub

Private Sub mnuClear_Click()
    Erase msgs
    txt.Text = vbNullString
End Sub

Private Sub mnuCopy_Click()
    On Error Resume Next
    Clipboard.Clear
    If txt.SelLength > 0 Then
         Clipboard.SetText txt.SelText
    Else
        Clipboard.SetText txt.Text
    End If
End Sub

Private Sub mnuOpenHomePage_Click()
    On Error Resume Next
    Const url = "http://www.vbforums.com/showthread.php?874127-Persistent-Debug-Print-Window"
    ShellExecuteA Me.hWnd, "open", url, "", "", 4
End Sub

Private Sub mnuSeparate_Click()
    out "<div>"
End Sub

Private Sub mnuFont_Click()
    cdl.Flags = cdlCFScreenFonts Or cdlCFForceFontExist
    '
    cdl.FontName = txt.FontName
    cdl.FontBold = txt.FontBold
    cdl.FontItalic = txt.FontItalic
    cdl.FontSize = txt.FontSize
    cdl.FontStrikethru = txt.FontStrikethru
    cdl.FontUnderline = txt.FontUnderline
    '
    cdl.ShowFont
    '
    txt.FontName = cdl.FontName
    txt.FontBold = cdl.FontBold
    txt.FontItalic = cdl.FontItalic
    txt.FontSize = cdl.FontSize
    txt.FontStrikethru = cdl.FontStrikethru
    txt.FontUnderline = cdl.FontUnderline
End Sub

Private Sub mnuBackColor_Click()
    ShowColorDialog Me.hWnd, txt.BackColor, , "BackColor"
    If ColorDialogSuccessful Then txt.BackColor = ColorDialogColor
End Sub

Private Sub mnuForeColor_Click()
    ShowColorDialog Me.hWnd, txt.BackColor, , "ForeColor"
    If ColorDialogSuccessful Then txt.ForeColor = ColorDialogColor
End Sub

Private Sub mnuReset_Click()
        Me.Left = 0&
        Me.Top = 0&
        Me.Width = 6600&
        Me.Height = 6600&
        '
        txt.FontName = "Fixedsys"
        txt.FontBold = False
        txt.FontItalic = False
        txt.FontSize = 9
        txt.FontStrikethru = False
        txt.FontUnderline = False
        '
        txt.BackColor = vbWhite
        txt.ForeColor = vbBlack
        mnuTimeStamp.Checked = False
        mnuTopMost.Checked = False
        SetWindowTopMost Me, False
End Sub

Public Sub out(s As String, Optional bHoldLine As Boolean)
    
    Dim supressTimestamp As Boolean
    
    'so apparently this trick doesnt work when SetWindowSubclass is used...
    If s = "<pause>" Or s = "<stop>" Then
        cmdContinue.Visible = True
        While Not cmdContinue.Visible
            DoEvents
            Sleep 15
        Wend
        Exit Sub
    End If
    
    If s = "<div>" Then
        s = vbCrLf & String(50, "-") & vbCrLf
        supressTimestamp = True
    End If
    
    If s = "<cls>" Or s = "<clear>" Then
        Erase msgs
        txt.Text = Empty
    Else
    
        If mnuTimeStamp.Checked And Not supressTimestamp Then
            s = Format(Now, "hh:nn:ss> ") & s
        End If

        push msgs, s
        
        If Len(Trim(txtFilter)) > 0 Then
            If InStr(1, s, txtFilter, vbTextCompare) > 0 Then
                AppendTxt s
            End If
        Else
            AppendTxt s
        End If
        
    End If
    
End Sub

Sub AppendTxt(s As String)
    SendMessageW txt.hWnd, EM_SETSEL, &H7FFFFFFF, ByVal &H7FFFFFFF          ' txt.SelStart = &H7FFFFFFF
    SendMessageW txt.hWnd, EM_REPLACESEL, 0, ByVal StrPtr(s & vbCrLf)   ' txt.SelText = s & vbCrLf
End Sub

Private Function FormIsFullyOnMonitor(frm As Form) As Boolean
    ' This tells us whether or not a form is FULLY visible on its monitor.
    '
    Dim hMonitor As Long
    Dim r1 As RECT
    Dim r2 As RECT
    Dim uMonInfo As MONITORINFO
    '
    hMonitor = hMonitorForForm(frm)
    GetWindowRect frm.hWnd, r1
    uMonInfo.cbSize = LenB(uMonInfo)
    GetMonitorInfo hMonitor, uMonInfo
    r2 = uMonInfo.rcWork
    '
    FormIsFullyOnMonitor = (r1.Top >= r2.Top) And (r1.Left >= r2.Left) And (r1.Bottom <= r2.Bottom) And (r1.Right <= r2.Right)
End Function

Public Function hMonitorForForm(frm As Form) As Long
    ' The monitor that the window is MOSTLY on.
    Const MONITOR_DEFAULTTONULL = &H0
    hMonitorForForm = MonitorFromWindow(frm.hWnd, MONITOR_DEFAULTTONULL)
End Function

Private Sub mnuTimeStamp_Click()
    mnuTimeStamp.Checked = Not mnuTimeStamp.Checked
End Sub

Private Sub mnuTopMost_Click()
    mnuTopMost.Checked = Not mnuTopMost.Checked
    SetWindowTopMost Me, mnuTopMost.Checked
End Sub

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function doSubtractiveFilter()
    Dim x, f(), matches() As String, m, found As Boolean
    
    x = Trim(txtFilter)
    If Len(x) = 1 Then Exit Function
    
    x = Trim(Mid(x, 2))
    If InStr(x, ",") > 0 Then
        matches = Split(x, ",")
    Else
        push matches, x
    End If

    For Each x In msgs
        found = False
        For Each m In matches
            If Len(m) > 0 Then
                If InStr(1, x, m, vbTextCompare) > 0 Then
                    found = True
                    Exit For
                End If
            End If
        Next
        If Not found Then push f, x
    Next
    
    If AryIsEmpty(f) Then
        txt = "0 results"
    Else
        txt = (UBound(f) + 1) & " results: " & vbCrLf & vbCrLf
        AppendTxt Join(f, vbCrLf)
    End If
            
            
End Function

Function doFilter()
    Dim x, f(), matches() As String, m, found As Boolean
    
    x = Trim(txtFilter)
    
    If InStr(x, ",") > 0 Then
        matches = Split(x, ",")
    Else
        push matches, x
    End If

    For Each x In msgs
        found = False
        For Each m In matches
            If Len(m) > 0 Then
                If InStr(1, x, m, vbTextCompare) > 0 Then
                    found = True
                    Exit For
                End If
            End If
        Next
        If found Then push f, x
    Next
    
    If AryIsEmpty(f) Then
        txt = "0 results"
    Else
        txt = (UBound(f) + 1) & " results: " & vbCrLf & vbCrLf
        AppendTxt Join(f, vbCrLf)
    End If
            
            
End Function

Private Sub txtFilter_Change()

    'On Error Resume Next
    
    Dim x, f()
    
    If Len(Trim(txtFilter)) = 0 Then
        txt = Join(msgs, vbCrLf)
    Else
    
        If AryIsEmpty(msgs) Then Exit Sub
            
        If Left(txtFilter, 1) = "-" Then
            doSubtractiveFilter
        Else
            doFilter
        End If
        
    End If
    
End Sub

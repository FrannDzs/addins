VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDebugPrint 
   Caption         =   "Persistent Debug Print Window"
   ClientHeight    =   4725
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
   ScaleHeight     =   4725
   ScaleWidth      =   7635
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

Private Declare Function ShellExecuteA Lib "shell32.dll" ( _
    ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Integer) As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Const HWND_TOPMOST = -1

Sub SetWindowTopMost(f As Form, Optional onTop As Boolean = True)
   SetWindowPos f.hWnd, IIf(onTop, HWND_TOPMOST, -2), f.Left / 15, _
        f.Top / 15, f.Width / 15, _
        f.Height / 15, Empty
End Sub


Private Sub Form_Load()
    On Error Resume Next
        
        If App.PrevInstance Then End
        
        SaveSetting "dbgWindow", "settings", "path", App.Path & "\dbgWindow.exe"
        
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
    '
    SubclassFormToReceiveStringMsg Me
End Sub

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
    If Not Me.WindowState = vbMinimized Then
        txt.Move 0&, 0&, Me.ScaleWidth, Me.ScaleHeight
    End If
End Sub

Private Sub mnuClear_Click()
    txt.Text = vbNullString
End Sub

Private Sub mnuCopy_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText txt.Text
End Sub

Private Sub mnuOpenHomePage_Click()
    On Error Resume Next
    Const url = "http://www.vbforums.com/showthread.php?874127-Persistent-Debug-Print-Window"
    ShellExecuteA Me.hWnd, "open", url, "", "", 4
End Sub

Private Sub mnuSeparate_Click()
     Out "<div>"
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

Public Sub Out(s As String, Optional bHoldLine As Boolean)

    Dim supressTimestamp As Boolean
    
    If s = "<div>" Then
        s = vbCrLf & String(50, "-") & vbCrLf
        supressTimestamp = True
    End If
    
    If s = "<cls>" Then
           txt.Text = Empty
    Else
    
        If mnuTimeStamp.Checked And Not supressTimestamp Then
            s = Format(Now, "hh:nn:ss> ") & s
        End If

        SendMessageW txt.hWnd, EM_SETSEL, &H7FFFFFFF, ByVal &H7FFFFFFF          ' txt.SelStart = &H7FFFFFFF
        If bHoldLine Then
            SendMessageW txt.hWnd, EM_REPLACESEL, 0, ByVal StrPtr(s)            ' txt.SelText = s
        Else
            SendMessageW txt.hWnd, EM_REPLACESEL, 0, ByVal StrPtr(s & vbCrLf)   ' txt.SelText = s & vbCrLf
        End If
    End If
    
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

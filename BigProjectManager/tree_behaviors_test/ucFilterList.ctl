VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ucFilterList 
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7605
   ScaleHeight     =   6315
   ScaleWidth      =   7605
   Begin VB.Timer tmrFilter 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   5880
      Top             =   4440
   End
   Begin VB.TextBox txtFilter 
      Height          =   330
      Left            =   540
      TabIndex        =   3
      Top             =   4320
      Width           =   1995
   End
   Begin MSComctlLib.ListView lvFilter 
      Height          =   3300
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5821
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4155
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image imgX 
      Height          =   225
      Left            =   2580
      Picture         =   "ucFilterList.ctx":0000
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   165
   End
   Begin VB.Label Label1 
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      TabIndex        =   2
      Top             =   4320
      Width           =   420
   End
   Begin VB.Menu mnuTools 
      Caption         =   "mnuTools"
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy All"
      End
      Begin VB.Menu mnuCopySel 
         Caption         =   "Copy Sel"
      End
      Begin VB.Menu mnuCopyColumn 
         Caption         =   "Copy Column"
      End
      Begin VB.Menu mnuTotalCol 
         Caption         =   "Total Column"
      End
      Begin VB.Menu mnuspacer4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilterHelp 
         Caption         =   "Filter Help"
      End
      Begin VB.Menu mnuDistinct 
         Caption         =   "Distinct"
      End
      Begin VB.Menu mnuDistinctStats 
         Caption         =   "Distinct Stats"
      End
      Begin VB.Menu mnuSetFilterCol 
         Caption         =   "Set Filter Column"
      End
      Begin VB.Menu mnuResults 
         Caption         =   "Results:"
      End
      Begin VB.Menu mnuspacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToggleMulti 
         Caption         =   "MultiSelect"
      End
      Begin VB.Menu mnuHideSel 
         Caption         =   "Hide Selection"
      End
      Begin VB.Menu mnuSelectInverse 
         Caption         =   "Inverse Selection"
      End
      Begin VB.Menu mnuAlertColWidths 
         Caption         =   "Alert Column Widths (IDE Only)"
      End
   End
End
Attribute VB_Name = "ucFilterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'author:  David Zimmer <dzzie@yahoo.com>
'site:    http://sandsprite.com
'License: free for any use
'
Option Explicit

Public AllowDelete As Boolean
Private abortFilter As Boolean
Private manuallyCleared As Boolean

Private m_updating As Boolean
Private m_Locked As Boolean
Private m_FilterColumn As Long
Private m_FilterColumnPreset As Long

'we need to track the index map between listviews in case they delete from lvFilter..
Private indexMapping As Collection
Public CustomFilters As New Collection

Event Click()
'Event ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Event DblClick()
Event ItemClick(ByVal Item As MSComctlLib.ListItem)
Event MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
Event OnCustomFilter(ByVal prefix As String, ByVal fullFilter As String)
Event KeyPress(keycode As Integer)
Event KeyDown(keycode As Integer, shift As Integer)

Const LVM_FIRST = &H1000
Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


#If 0 Then
    Dim x, y, Column, nextone 'force lowercase so ide doesnt switch around on its own whim...
#End If

Function ClearFilters()
    Set CustomFilters = New Collection
End Function

Function RegisterFilter(prefix) As Boolean
    On Error Resume Next
    CustomFilters.Add prefix, prefix
    RegisterFilter = (Err.Number = 0)
End Function

Function RemoveFilter(prefix) As Boolean
    On Error Resume Next
    CustomFilters.Remove prefix
    RemoveFilter = (Err.Number = 0)
End Function

'use this instead of usercontrol.setfocus which may cause bug..(you can see the bug lvFiltered will be greyed)
Sub SetFocus2()
    If lvFilter.Visible Then lvFilter.SetFocus Else lv.SetFocus
End Sub

'note when locked you wont receive events, and can not add items..
Property Get Locked() As Boolean
    Locked = m_Locked
End Property

Property Let Locked(x As Boolean)
    m_Locked = x
    txtFilter.BackColor = IIf(x, &HC0C0C0, vbWhite)
    txtFilter.Enabled = Not x
End Property
    
Property Get SelCount() As Long
    Dim v As ListView
    Dim li As ListItem
    Dim cnt As Long
    
    Set v = currentLV
'    For Each li In v.ListItems
'        If li.selected Then cnt = cnt + 1
'    Next
    
    cnt = SendMessage(v.hwnd, LVM_GETSELECTEDCOUNT, 0, 0)
    SelCount = cnt
    
End Property

Property Get selItems() As Collection

    Dim c As New Collection
    Dim li As ListItem
    Dim cnt As Long
    
    Set selItems = c
 
    For Each li In currentLV.ListItems
        If li.selected Then c.Add li
    Next
    
End Property
    
Property Get FilterColumn() As Long
    FilterColumn = m_FilterColumn
End Property

Property Let FilterColumn(x As Long)
    On Error Resume Next
    Dim tmp As String
    Dim ch As ColumnHeader
    
    If lv.ColumnHeaders.Count = 0 Then
        m_FilterColumnPreset = x
        Exit Property
    End If
    
    If x <= 0 Then x = 1
    
    If x > lv.ColumnHeaders.Count Then
        x = lv.ColumnHeaders.Count
    End If
    
    'remove the visual marker that this is the filter column
    Set ch = lv.ColumnHeaders(m_FilterColumn)
    ch.Text = Trim(Replace(ch.Text, "*", Empty))
    
    Set ch = lvFilter.ColumnHeaders(m_FilterColumn)
    ch.Text = Trim(Replace(ch.Text, "*", Empty))

    'add the visual marker to the new column
    Set ch = lv.ColumnHeaders(x)
    ch.Text = ch.Text & " *"
    
    Set ch = lvFilter.ColumnHeaders(x)
    ch.Text = ch.Text & " *"

    m_FilterColumn = x
    
End Property

'doesnt seem to work as intended in all cases?
'note this only hands out a ref to the main listview not filtered
'this is only for compatability with existing code to make integration easier..
Property Get ListItems() As ListItems
    Set ListItems = lv.ListItems
End Property

Property Get MultiSelect() As Boolean
    MultiSelect = lv.MultiSelect
End Property

Property Let MultiSelect(x As Boolean)
    lv.MultiSelect = x
    lvFilter.MultiSelect = x
    mnuToggleMulti.Checked = x
End Property

Property Get HideSelection() As Boolean
    HideSelection = lv.MultiSelect
End Property

Property Let HideSelection(x As Boolean)
    lv.HideSelection = x
    lvFilter.HideSelection = x
    mnuHideSel.Checked = x
End Property

Property Get GridLines() As Boolean
    GridLines = lv.GridLines
End Property

Property Let GridLines(x As Boolean)
    lv.GridLines = x
    lvFilter.GridLines = x
End Property

'which ever one is currently displayed
Property Get currentLV() As ListView
    On Error Resume Next
    If lvFilter.Visible Then
        Set currentLV = lvFilter
    Else
        Set currentLV = lv
    End If
End Property

Property Get mainLV() As ListView
    Set mainLV = lv
End Property


'compatability with normal listview
Property Get SelectedItem() As ListItem
    Set SelectedItem = selItem
End Property

Property Get selItem() As ListItem
    On Error Resume Next
    If lvFilter.Visible Then
        Set selItem = lvFilter.SelectedItem
    Else
        Set selItem = lv.SelectedItem
    End If
End Property

Property Get Filter() As String
    Filter = txtFilter
End Property

Property Let Filter(txt As String)
     txtFilter = txt
End Property

Function AddItem(txt, ParamArray subItems()) As ListItem
    On Error Resume Next
    
    Dim i As Integer, si
    
    If m_Locked Then Exit Function
    
    Set AddItem = lv.ListItems.Add(, , CStr(txt))
    
    For Each si In subItems
        AddItem.subItems(i + 1) = si
        i = i + 1
    Next
    
    ApplyFilter
    
End Function

Function AddAryItem(rowItems) As ListItem
    On Error Resume Next
    
    Dim i As Integer, si
    
    If m_Locked Then Exit Function
    If Not IsArray(rowItems) Then Exit Function
    If AryIsEmpty(rowItems) Then Exit Function
    
    Set AddAryItem = lv.ListItems.Add(, , CStr(rowItems(0)))
    
    For Each si In rowItems
        If i > 0 Then
            If i > lv.ColumnHeaders.Count Then Exit For
            If IsNumeric(si) Then
                AddAryItem.subItems(i) = lpad(si, 8) 'so it sorts properly
            Else
                AddAryItem.subItems(i) = si
            End If
        End If
        i = i + 1
    Next
    
    ApplyFilter
    
End Function

Function LoadRecordset(rs) As Long
    On Error Resume Next
    'Dim rs As Recordset
    Dim i As Long, f() As String, li As ListItem, t(), x
    
    Me.Clear
    Me.ClearColumns
    Me.ClearFilters
    
    If TypeName(rs) <> "Recordset" Then
        MsgBox "FilterList.LoadRecordset unexpected type: " & TypeName(rs)
        Exit Function
    End If
        
    For i = 0 To rs.Fields.Count
        push f, rs.Fields(i).name
    Next
    
    Me.SetColumnHeaders Join(f, ",")
    
    While Not rs.EOF
        Erase t()
        
        For Each x In f
            push t, rs(x)
        Next
        
        Set li = Me.AddAryItem(t)
        rs.MoveNext
    Wend
    
    LoadRecordset = Me.ListItems.Count
    
End Function

Sub Clear(Optional andFilter As Boolean = True)

    m_updating = False
    If m_Locked Then Exit Sub
    
    Dim li As ListItem
    For Each li In lv.ListItems
        If IsObject(li.Tag) Then Set li.Tag = Nothing
    Next
    
    For Each li In lvFilter.ListItems
        If IsObject(li.Tag) Then Set li.Tag = Nothing
    Next
    
    lv.ListItems.Clear
    lvFilter.ListItems.Clear
    If andFilter Then txtFilter.Text = Empty
    
End Sub

Sub SetFont(name As String, size As Long)
    lv.Font.name = name
    lv.Font.size = size
    lvFilter.Font.name = name
    lvFilter.Font.size = size
    txtFilter.Font.name = name
    txtFilter.Font.size = size
End Sub

Sub ClearColumns()
    lv.ColumnHeaders.Clear
    lvFilter.ColumnHeaders.Clear
End Sub

Sub SetColumnHeaders(csvList As String, Optional csvWidths As String)
    
    On Error Resume Next
    Dim i As Long, fc As Long, ch As ColumnHeader, tmp() As String, t
    
    fc = -1
    lv.ColumnHeaders.Clear
    lvFilter.ColumnHeaders.Clear
    
    tmp = Split(csvList, ",")
    For Each t In tmp
        i = i + 1
        If InStr(t, "*") > 0 Then
            fc = i
            t = Trim(Replace(t, "*", Empty))
        End If
        lv.ColumnHeaders.Add , , Trim(t)
        lvFilter.ColumnHeaders.Add , , Trim(t)
    Next
    
    If fc <> -1 Then FilterColumn = fc  'this sets the visual marker on the column if they specified it..
    If m_FilterColumnPreset <> -1 Then FilterColumn = m_FilterColumnPreset 'they called FilterColumn manually first, now apply..
    If m_FilterColumn = -1 Then FilterColumn = 1 'they never specified it so default to first column
    
    If Len(csvWidths) > 0 Then
        tmp = Split(csvWidths, ",")
        For i = 0 To UBound(tmp)
            If Len(tmp(i)) > 0 Then
                lv.ColumnHeaders(i + 1).Width = CLng(tmp(i))
                lvFilter.ColumnHeaders(i + 1).Width = CLng(tmp(i))
            End If
        Next
    End If
    
End Sub

Sub SelectAll(Optional selected As Boolean = True)
    Dim v As ListView, li As ListItem
    
    If Not Me.MultiSelect Then Exit Sub
    If lv.Visible Then Set v = lv Else Set v = lvFilter
    
    For Each li In v.ListItems
        li.selected = selected
    Next
        
End Sub

Private Sub imgX_Click()
    txtFilter.Text = Empty
End Sub

Private Sub Label1_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
    Label1.ToolTipText = currentLV.ListItems.Count
End Sub

Private Sub lv_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub lvFilter_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub lv_KeyDown(keycode As Integer, shift As Integer)

    Dim i As Long
    Dim li As ListItem
    
    On Error Resume Next
    
    If m_Locked Then Exit Sub
    
    If keycode = vbKeyDelete And AllowDelete Then
        For i = lv.ListItems.Count To 1 Step -1
            If lv.ListItems(i).selected Then lv.ListItems.Remove i
        Next
    End If
    
    If keycode = vbKeyA And shift = 2 Then SelectAll 'ctrl-a
    
    If keycode = vbKeyC And shift = 2 Then 'ctrl-C
        Clipboard.Clear
        Clipboard.SetText GetAllElements()
    End If
    
    RaiseEvent KeyDown(keycode, shift)
    
End Sub

Private Sub lvFilter_KeyDown(keycode As Integer, shift As Integer)
    Dim i As Long
    Dim liMain As ListItem
    
    On Error Resume Next
    
    If m_Locked Then Exit Sub
    
    If keycode = vbKeyDelete And AllowDelete Then
        For i = lvFilter.ListItems.Count To 1 Step -1
            If lvFilter.ListItems(i).selected Then
                Set liMain = getMainListItemFor(lvFilter.ListItems(i))
                If Not liMain Is Nothing Then lv.ListItems.Remove liMain.Index
                lvFilter.ListItems.Remove i
            End If
        Next
    End If
    
    If keycode = vbKeyA And shift = 2 Then SelectAll 'ctrl-a
             
    If keycode = vbKeyC And shift = 2 Then 'ctrl-C
        Clipboard.Clear
        Clipboard.SetText GetAllElements()
    End If
    
    RaiseEvent KeyDown(keycode, shift)
    
End Sub


Private Sub mnuAlertColWidths_Click()
    Dim tmp(), c As ColumnHeader
    For Each c In lv.ColumnHeaders
        push tmp, Round(c.Width)
    Next
    InputBox "Column Widths are: ", , Join(tmp, ",")
End Sub

Private Sub Label1_Click()
    If m_Locked Then Exit Sub
    mnuResults.Caption = "Results: " & Me.currentLV.ListItems.Count
    PopupMenu mnuTools
End Sub

Private Sub mnuCopyAll_Click()
    Clipboard.Clear
    Clipboard.SetText Me.GetAllElements()
End Sub

Private Sub mnuCopyColumn_Click()
    On Error Resume Next
    Dim x, c As Long
    x = InputBox("Enter column index or name to copy", , 1)
    If Len(x) = 0 Then Exit Sub
    c = CLng(x) - 1 'we are 0 based internally..
    If Err.Number <> 0 Then
        c = ColIndexForName(x)
    End If
    Clipboard.Clear
    Clipboard.SetText Me.GetAllText(c)
End Sub

Private Sub mnuCopySel_Click()
    Clipboard.Clear
    Clipboard.SetText Me.GetAllElements(True)
End Sub

Private Sub mnuDistinct_Click()
    txtFilter.Text = "/distinct"
End Sub

'Private Sub mnuDistinctStats_Click()
'    On Error Resume Next
'    Dim colIndex As Long, stats As CollectionEx, s, tmp(), cnt As Long, pcent, fName As String, i As Long
'    colIndex = InputBox("Enter column index: ", , m_FilterColumn)
'    If Err.Number <> 0 Then Exit Sub
'
'    Set stats = distinctCounts(colIndex)
'    stats.OptionForceAsKey = False
'    stats.Sort
'
'    push tmp, lv.ColumnHeaders(colIndex).Text & ": Hits : %"
'
'    For i = 1 To stats.Count
'        cnt = stats(i, 0)
'        pcent = Round((cnt / lv.ListItems.Count) * 100, 2)
'        push tmp, stats.keyForIndex(i) & " : " & cnt & " : " & pcent & "%"
'    Next
'
'    push tmp, vbCrLf
'    push tmp, "Distinct Elements: " & stats.Count
'    push tmp, "Hits Total: " & lv.ListItems.Count
'
'    fName = Environ("temp") & "\tmp.txt"
'    If Dir(fName) <> "" Then Kill fName
'    WriteFile fName, Join(tmp, vbCrLf)
'    Shell "notepad.exe """ & fName & """", vbNormalFocus
'
'End Sub

Private Sub WriteFile(path, it)
    On Error Resume Next
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Private Sub mnuFilterHelp_Click()
    
    Dim h As String, v
    
    Const msg = "You can enter multiple criteria to filter \n" & _
                "on by seperating with commas. You can also\n" & _
                "utilize a subtractive filter if the first \n" & _
                "character in the textbox is a minus sign\n" & _
                "Filter also understands: bold,selected, color:red|blue|etc\n\n" & _
                "The FilterColumn is marked with an * this is \n" & _
                "the column that is being searched. You can \n" & _
                "modify it on the filter menu, or by entering\n" & _
                "/[index] in the filter textbox and hitting return\n" & _
                "/t [index|colName] will alert you to the total of the column values\n" & _
                "/a [index|colName] will alert you to the avgerage of the column values\n" & _
                "/c [index|colName] will copy column\n" & _
                "/d will filter to distinct entries only\n" & _
                "/mid(2) or /mid(2,5) /p are also supported (/p means alter parent list must reload to undo)\n" & _
                "numeric columns also support > < = !  filters\n\n" & _
                "Pressing escape in the filter textbox will clear it.\n\n" & _
                "If the AllowDelete property has been set, you can\n" & _
                "select list items and press the delete key to remove\n" & _
                "them."
                
          
    If CustomFilters.Count > 0 Then
        h = vbCrLf & vbCrLf & "Host provided filters: "
        For Each v In CustomFilters
            h = h & v & ", "
        Next
        h = Mid(h, 1, Len(h) - 2)
    End If
    
    MsgBox Replace(msg, "\n", vbCrLf) & h, vbInformation
                
End Sub

Private Sub mnuHideSel_Click()
    Me.HideSelection = Not lv.HideSelection
End Sub

Private Sub mnuSelectInverse_Click()
    InvertSelection
End Sub

Public Sub InvertSelection()
    If Not MultiSelect Then Exit Sub
    Dim li As ListItem
    For Each li In Me.currentLV.ListItems
        li.selected = Not li.selected
    Next
End Sub

Private Sub mnuSetFilterCol_Click()
    On Error Resume Next
    Dim x As Long
    x = InputBox("Enter column that filter searches", , FilterColumn)
    If Len(x) = 0 Then Exit Sub
    x = CLng(x)
    FilterColumn = x
End Sub

Private Sub mnuToggleMulti_Click()
    Me.MultiSelect = Not lv.MultiSelect
End Sub

Function ColorConstantsToLong(ByVal s As String) As Long
    
    Dim c As ColorConstants
    s = LCase(s)
    
    c = -1
    If InStr(s, "black") > 0 Then c = vbBlack
    If InStr(s, "blue") > 0 Then c = vbBlue
    If InStr(s, "cyan") > 0 Then c = vbCyan
    If InStr(s, "green") > 0 Then c = vbGreen
    If InStr(s, "magenta") > 0 Then c = vbMagenta
    If InStr(s, "red") > 0 Then c = vbRed
    If InStr(s, "white") > 0 Then c = vbWhite
    If InStr(s, "yellow") > 0 Then c = vbYellow
    
    ColorConstantsToLong = c
    
End Function


Private Sub mnuTotalCol_Click()
    Dim i As Long, tmp As String, b As Boolean, tot As Long
    On Error Resume Next
    tmp = InputBox("Enter column name or index to total (1-" & (lv.ColumnHeaders.Count) & ")")
    If Len(tmp) = 0 Then Exit Sub
    i = CLng(tmp)
    If Err.Number <> 0 Then
        i = ColIndexForName(tmp) + 1
    End If
    tot = totalColumn(i, b)
    MsgBox "Total for " & lv.ColumnHeaders(i).Text & " = " & tot & IIf(b, " An error was generated", ""), vbInformation
End Sub

Private Sub tmrFilter_Timer()
    tmrFilter.Enabled = False
    Call ApplyFilter
End Sub

'on huge lists it can take a while so let them finish typing first
Private Sub txtFilter_Change()
    abortFilter = True
    If lv.ListItems.Count > 100 Then
        tmrFilter.Enabled = False 'reset the timer it will apply once they pause and wait
        tmrFilter.Enabled = True
    Else
        ApplyFilter
    End If
End Sub

Function distinctFilter()
    Dim vals As New Collection, v As String
    Dim li As ListItem
    
    On Error Resume Next
    
    If m_FilterColumn = -1 Then Exit Function
    
    lvFilter.Visible = True
    lvFilter.ListItems.Clear
    
    For Each li In lv.ListItems
    
        If abortFilter Then Exit For
        
        If m_FilterColumn = 1 Then
            v = li.Text
        Else
            v = li.subItems(m_FilterColumn - 1)
        End If
        If Not KeyExistsInCollection(vals, v) Then
            vals.Add v, v
            CloneListItemTo li, lvFilter
            'Debug.Print "unique: " & v
        End If
    Next
    
End Function

Function midFilter()
    Dim vals As New Collection, v As String
    Dim li As ListItem, li2 As ListItem
    
    On Error Resume Next
    
    'on txtFilter_keypress return requires /mid()
    'supports /mid(5) or /mid(5,9) /p means parent list
    
    Dim a As Long, b As Long
    Dim tmp, parentList As Boolean
    
    tmp = Replace(txtFilter, "/mid(", Empty)
    
    If InStr(tmp, "/p") > 0 Then
        tmp = Trim(Replace(tmp, "/p", Empty))
        parentList = True
    End If
    
    tmp = Trim(Replace(tmp, ")", Empty))
    
    If InStr(tmp, ",") > 0 Then
        tmp = Split(tmp, ",")
        a = CLng(tmp(0))
        b = CLng(tmp(1))
    Else
        a = CLng(tmp)
    End If
    
    If Err.Number > 0 Then
        MsgBox Err.Description
        Exit Function
    End If
    
    If m_FilterColumn = -1 Then Exit Function
    
    If Not parentList Then
        lvFilter.Visible = True
        lvFilter.ListItems.Clear
    End If
    
    For Each li In lv.ListItems
    
        If abortFilter Then Exit For
        
        If m_FilterColumn = 1 Then
            v = li.Text
        Else
            v = li.subItems(m_FilterColumn - 1)
        End If
        
        If b > 0 Then
            v = Mid(v, a, b - a)
        Else
            v = Mid(v, a)
        End If
        
        If parentList Then
            If m_FilterColumn = 1 Then
                li.Text = v
            Else
                li.subItems(m_FilterColumn - 1) = v
            End If
        Else
            Set li2 = CloneListItemTo(li, lvFilter)
            
            If m_FilterColumn = 1 Then
                li2.Text = v
            Else
                li2.subItems(m_FilterColumn - 1) = v
            End If
        End If
        
    Next
    
End Function



Private Function myIsNumeric(ByVal v, ByRef outV As Long) As Boolean
    On Error GoTo hell
    If IsNumeric(v) Then
        outV = CLng(v)
    Else
        v = Replace(v, "0x", Empty)
        outV = CLng("&h" & v)
    End If
    myIsNumeric = True
hell:
End Function

Private Function ColIndexForName(n) As Long
    On Error Resume Next
    Dim i As Long
    If Len(n) > 0 Then
        For i = 1 To lv.ColumnHeaders.Count
            If Left(LCase(lv.ColumnHeaders(i).Text), Len(n)) = LCase(n) Then
                ColIndexForName = i - 1
                Exit Function
            End If
        Next
    End If
    ColIndexForName = -1
End Function

'Function distinctCounts(Optional colIndex As Long = -1) As CollectionEx
'    Dim vals As New CollectionEx, v As String
'    Dim li As ListItem
'
'    vals.OptionForceAsKey = True
'    Set distinctCounts = vals
'
'    If colIndex = -1 Then colIndex = m_FilterColumn
'    If colIndex = -1 Then Exit Function
'
'    For Each li In lv.ListItems
'        If colIndex = 1 Then
'            v = li.Text
'        Else
'            v = li.subItems(colIndex - 1)
'        End If
'        If Not vals.keyExists(v) Then
'            vals.Add 1, v
'        Else
'            vals(v, 1) = vals(v, 1) + 1
'        End If
'    Next
'
'End Function

'so we dont try to apply filter for every single item added when first loading list
'speeds things up a big (big) lot
Sub BeginUpdate()
    m_updating = True
End Sub

Sub EndUpdate()
    m_updating = False
    ApplyFilter
End Sub

Sub ApplyFilter()
    Dim li As ListItem
    Dim t As String
    Dim useSubtractiveFilter As Boolean
    Dim tmp() As String, addIt As Boolean, x
    Dim gtMode As Boolean, ltMode As Boolean, eqMode As Boolean, notMode As Boolean
    Dim uv As Long, v As Long
    
     On Error Resume Next
    'applying the filter can be a long process on huge sets..if they start to enter a command
    'we need to be sure not to trigger to early. so only on return key for / commands..(keypress event)
    
    If m_Locked Then Exit Sub
    If m_updating Then Exit Sub
    
    If manuallyCleared Then
         manuallyCleared = False
         Exit Sub
    End If
    
    If Len(txtFilter) = 0 Then GoTo hideExit
    
    abortFilter = False
    txtFilter.BackColor = vbYellow
    
    'distinct filter leave in place across reloads until they manually delete..
    If txtFilter = "/distinct" Or txtFilter = "/d" Then
         distinctFilter
         txtFilter.BackColor = vbWhite
         Exit Sub
    End If
    
    If VBA.Left(txtFilter, 1) = "/" Then Exit Sub 'GoTo hideExit 'wait until a return is hit to process these..
    
    'I want within:5/10 mode or similar
    If VBA.Left(txtFilter, 1) = ">" Then
        gtMode = True
        If Len(Trim(txtFilter)) = 1 Then GoTo hideExit
    End If
    
    If VBA.Left(txtFilter, 1) = "=" Then
        eqMode = True
        If Len(Trim(txtFilter)) = 1 Then GoTo hideExit
    End If
    
    If VBA.Left(txtFilter, 1) = "!" Then
        notMode = True
        If Len(Trim(txtFilter)) = 1 Then GoTo hideExit
    End If
    
    If VBA.Left(txtFilter, 1) = "<" Then
        ltMode = True
        If Len(Trim(txtFilter)) = 1 Then GoTo hideExit
    End If
    
    If VBA.Left(txtFilter, 1) = "-" Then 'they are typing a subtractive filter..give them time to formulate it..
        If Len(txtFilter) = 1 Then GoTo hideExit
        If VBA.Right(txtFilter, 1) = "," Then Exit Sub 'they are adding more criteria
    End If
    
    'should multiple (csv) filters only apply on hitting return?
    'so you can see full list to work off of?
    
    lvFilter.Visible = True
    lvFilter.ListItems.Clear
    Set indexMapping = New Collection
    txtFilter.BackColor = vbYellow 'in progress indicator
    
    Dim sMatch As String
    Dim isColor As Boolean
    Dim lColor As Long
    Dim startFilter As String
    
    If VBA.Left(txtFilter, 1) = "-" Then
        useSubtractiveFilter = True
        sMatch = Mid(txtFilter, 2)
    ElseIf VBA.Left(txtFilter, 6) = "color:" Then
        isColor = True
        sMatch = Replace(txtFilter, "color:", Empty)
        If Len(sMatch) = 0 Then Exit Sub 'they are still entering it...
        Err.Clear
        lColor = CLng(sMatch)
        If Err.Number <> 0 Then lColor = ColorConstantsToLong(sMatch)
        If lColor = -1 Then Exit Sub
    Else
        sMatch = txtFilter
    End If
    
    If ltMode Or gtMode Or eqMode Or notMode Then
        t = Mid(txtFilter, 2)
        If Not myIsNumeric(t, uv) Then 'we will use converted UserVal (uv) value below..
            ltMode = False
            gtMode = False
            eqMode = False
            notMode = False
        End If
    End If
     
    'we allow for csv multiple criteria, also
    'you can use a subtractive filter like -mnu,cmd,lv
     For Each li In lv.ListItems
     
         If abortFilter Then Exit For
         
         If lvFilter.ListItems.Count > 50 Then LockWindowUpdate lvFilter.hwnd 'they cant see it anymore anyway stop flicker
         
         If FilterColumn = 1 Then
            t = li.Text
         Else
            t = li.subItems(m_FilterColumn - 1)
         End If
         
         addIt = False
         
         If gtMode Or ltMode Or eqMode Or notMode Then
            If myIsNumeric(t, v) Then
                If gtMode Then If v > uv Then addIt = True
                If ltMode Then If v < uv Then addIt = True
                If eqMode Then If v = uv Then addIt = True
                If notMode Then If v <> uv Then addIt = True
            End If
         ElseIf txtFilter = "bold" Then
            If li.Bold = True Then addIt = True
         ElseIf txtFilter = "selected" Then
            If li.selected = True Then addIt = True
         ElseIf isColor Then
            If li.ForeColor = lColor Then addIt = True
         Else
            addIt = useSubtractiveFilter
            If InStr(txtFilter, ",") Then
               tmp = Split(sMatch, ",")
            Else
               push tmp, sMatch
            End If
            
            For Each x In tmp
                If Len(x) > 0 Then
                    If InStr(1, t, x, vbTextCompare) > 0 Then
                        addIt = Not addIt
                        Exit For
                    End If
                End If
            Next
         End If
         
         If addIt Then
             CloneListItemTo li, lvFilter
         End If
      
         If lvFilter.ListItems.Count Mod 20 = 0 Then DoEvents
     Next

     LockWindowUpdate 0
     txtFilter.BackColor = vbWhite
     
Exit Sub

hideExit:
            lvFilter.Visible = False
            txtFilter.BackColor = vbWhite
            Exit Sub
            
    
End Sub

Function RemoveItem(li As ListItem) As Boolean
    On Error Resume Next
    Dim i As ListItem, mainItem As ListItem
    
    If lvFilter.Visible Then
        For Each i In lvFilter.ListItems
            If ObjPtr(li) = ObjPtr(i) Then
                'its from filter list for sure.
                'how added: indexMapping.Add li, "fObj:" & ObjPtr(li2)
                
                lvFilter.ListItems.Remove li.Index
                
                Set mainItem = indexMapping("fObj:" & ObjPtr(li))
                
                If mainItem Is Nothing Then
                    Debug.Print "RemoveItem: Main item not found for " & ObjPtr(li)
                Else
                    lv.ListItems.Remove mainItem.Index
                    RemoveItem = True
                End If
                
                Exit Function
            End If
        Next
        'it wasnt found? from main list then?
        GoTo checkMainLV
    Else
checkMainLV:
        For Each i In lv.ListItems
            If ObjPtr(li) = ObjPtr(i) Then
                'its from main list for sure.
                lv.ListItems.Remove li.Index
                RemoveItem = True
                Exit Function
            End If
        Next
        'it wasnt found was it cached from filter list item? lets just fail...
        Exit Function 'return false
    End If
    
End Function

Function CloneListItemTo(li As ListItem, lv As ListView) As ListItem
    Dim li2 As ListItem, i As Integer
    
    Set li2 = lv.ListItems.Add(, , li.Text)
    Set CloneListItemTo = li2
    
    For i = 1 To lv.ColumnHeaders.Count - 1
        li2.subItems(i) = li.subItems(i)
    Next
    
    If li.ForeColor <> vbBlack Then SetLiColor li2, li.ForeColor
    If li.selected Then li2.selected = True
    
    On Error Resume Next
    If IsObject(li.Tag) Then
        Set li2.Tag = li.Tag
    Else
        li2.Tag = li.Tag
    End If
    
    indexMapping.Add li, "fObj:" & ObjPtr(li2)  'filter list item obj to lvFilter objPtr map
    
End Function

'we had to switch from index mapping to object mapping to account for column click sorts..
Private Function getMainListItemFor(liFilt As ListItem) As ListItem
    On Error Resume Next
    Set getMainListItemFor = indexMapping("fObj:" & ObjPtr(liFilt))
End Function

Private Sub lv_Click()
    If m_Locked Then Exit Sub
    RaiseEvent Click
End Sub

Sub triggerColumnSort(colIndex As Long)
    On Error Resume Next
    If lvFilter.Visible = True Then
        lvFilter_ColumnClick lvFilter.ColumnHeaders(colIndex)
    Else
        lv_ColumnClick lv.ColumnHeaders(colIndex)
    End If
End Sub

Private Sub lv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If m_Locked Then Exit Sub
    Me.ColumnSort ColumnHeader
    'RaiseEvent ColumnClick(ColumnHeader)
End Sub

Private Sub lv_DblClick()
    If m_Locked Then Exit Sub
    RaiseEvent DblClick
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If m_Locked Then Exit Sub
    If Me.SelCount > 1 Then Exit Sub 'uses sendmessage its ok..
    RaiseEvent ItemClick(Item)
End Sub

Private Sub lv_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
    If m_Locked Then Exit Sub
    RaiseEvent MouseUp(Button, shift, x, y)
End Sub

Private Sub lvFilter_Click()
    If m_Locked Then Exit Sub
    RaiseEvent Click
End Sub

Private Sub lvFilter_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If m_Locked Then Exit Sub
    Me.ColumnSort ColumnHeader
    'RaiseEvent ColumnClick(ColumnHeader)
End Sub

Private Sub lvFilter_DblClick()
    If m_Locked Then Exit Sub
    RaiseEvent DblClick
End Sub

Private Sub lvFilter_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If m_Locked Then Exit Sub
    If Me.SelCount > 1 Then Exit Sub 'uses sendmessage its ok..
    RaiseEvent ItemClick(Item)
End Sub

Private Sub lvFilter_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
    If m_Locked Then Exit Sub
    RaiseEvent MouseUp(Button, shift, x, y)
End Sub

Function CopyColumn(OneBasedIndex As Long) As String
    CopyColumn = Me.GetAllText(OneBasedIndex - 1)
End Function

Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    'MsgBox KeyAscii
    
    On Error Resume Next
    Dim t As String, b As Boolean, tot As Long, i As Long
    Dim addIt As Boolean, uv As Long, v As Long, cf
    
    If m_Locked Then Exit Sub
    
    abortFilter = True
    
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Filter = Empty
    End If
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Len(txtFilter) > 0 Then
        
            For Each cf In CustomFilters
                If LCase(Left(txtFilter, Len(cf))) = LCase(cf) Then
                    RaiseEvent OnCustomFilter(cf, txtFilter)
                    Exit Sub
                End If
            Next
        
            If Left(txtFilter, 1) = "/" Then
                t = Replace(txtFilter, "/", Empty)
                
                'total mode /t <index or name>
                If Left(txtFilter, 2) = "/t" Then
                     t = Trim(Mid(txtFilter, 3))
                     If myIsNumeric(t, uv) Then  'we will use converted UserVal (uv) value below..
                        If uv > 0 Or uv <= lv.ColumnHeaders.Count Then
                            v = totalColumn(uv, addIt)
                            MsgBox "Total for " & lv.ColumnHeaders(uv).Text & " = " & v & IIf(addIt, " - An error was generated", ""), vbInformation
                            manuallyCleared = True
                            txtFilter = Empty
                            Exit Sub
                        End If
                     Else
                        uv = ColIndexForName(t) + 1 '0 based , -1 on error
                        If uv > 0 Then
                            tot = totalColumn(uv, addIt)
                            MsgBox "Total for " & lv.ColumnHeaders(uv).Text & " = " & tot & IIf(addIt, vbCrLf & vbCrLf & " - An error was generated", ""), vbInformation
                            manuallyCleared = True
                            txtFilter = Empty
                             Exit Sub
                        End If
                     End If
                    
                End If
                
                'average mode /t <index or name>
                If Left(txtFilter, 2) = "/a" Then
                     t = Trim(Mid(txtFilter, 3))
                     If myIsNumeric(t, uv) Then  'we will use converted UserVal (uv) value below..
                        If uv > 0 Or uv <= lv.ColumnHeaders.Count Then
                            v = avgColumn(uv, addIt)
                            MsgBox "Average for " & lv.ColumnHeaders(uv).Text & " = " & v & IIf(addIt, " - An error was generated", ""), vbInformation
                            manuallyCleared = True
                            txtFilter = Empty
                            Exit Sub
                        End If
                     Else
                        uv = ColIndexForName(t) + 1 '0 based , -1 on error
                        If uv > 0 Then
                            tot = avgColumn(uv, addIt)
                            MsgBox "Average for " & lv.ColumnHeaders(uv).Text & " = " & tot & IIf(addIt, vbCrLf & vbCrLf & " - An error was generated", ""), vbInformation
                            manuallyCleared = True
                            txtFilter = Empty
                             Exit Sub
                        End If
                     End If
                    
                End If
                
                'copy column mode /c <index or name>
                If Left(txtFilter, 2) = "/c" Then
                     t = Trim(Mid(txtFilter, 3))
                     If myIsNumeric(t, uv) Then  'we will use converted UserVal (uv) value below..
                        If uv > 0 Or uv <= lv.ColumnHeaders.Count Then
                            t = CopyColumn(uv)
                            Clipboard.Clear
                            Clipboard.SetText t
                            manuallyCleared = True
                            txtFilter = Empty
                            Exit Sub
                        End If
                     Else
                        uv = ColIndexForName(t) + 1 '0 based , -1 on error
                        If uv > 0 Then
                            t = CopyColumn(uv)
                            Clipboard.Clear
                            Clipboard.SetText t
                            manuallyCleared = True
                            txtFilter = Empty
                             Exit Sub
                        End If
                     End If
                    
                End If
                
                If Left(txtFilter, 5) = "/mid(" And InStr(txtFilter, ")") > 0 Then
                    midFilter
                    Exit Sub
                End If
'                'distinct filter we want to be able to leave this one in place across reloads..
'                If txtFilter = "/distinct" Or txtFilter = "/d" Then
'                     lvFilter.Visible = True
'                     distinctFilter
'                     'manuallyCleared = True
'                     'txtFilter = Empty
'                     Exit Sub
'                End If

                'trying to change the filter column by index or name?
                'this must be last since we shortcut allow partial names in colIndexforName (/t /d would to easily match all)
                If IsNumeric(t) Then
                    FilterColumn = CLng(t)
                    manuallyCleared = True
                    Filter = Empty
                    Exit Sub
                Else
                    i = ColIndexForName(t)
                    If i <> -1 Then
                        FilterColumn = i + 1
                        manuallyCleared = True
                        Filter = Empty
                        Exit Sub
                    End If
                End If
                
                
                
            End If
            
        End If
    End If
            
End Sub


Private Sub UserControl_Initialize()
    m_FilterColumn = -1
    m_FilterColumnPreset = -1
    mnuAlertColWidths.Visible = isIde()
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With UserControl
        lv.Top = 0
        lv.Left = 0
        lv.Width = .Width
        lv.Height = .Height - txtFilter.Height - 300
        txtFilter.Top = .Height - txtFilter.Height - 150
        txtFilter.Width = .Width - txtFilter.Left - imgX.Width '- lblTools.Width - 100
        imgX.Top = txtFilter.Top
        imgX.Left = txtFilter.Width + txtFilter.Left ' 100
        'lblTools.Left = .Width - lblTools.Width
        Label1.Top = txtFilter.Top + 30
        'lblTools.Top = txtFilter.Top + 30
    End With
    lvFilter.Move lv.Left, lv.Top, lv.Width, lv.Height
    lv.ColumnHeaders(lv.ColumnHeaders.Count).Width = lv.Width - lv.ColumnHeaders(lv.ColumnHeaders.Count).Left - 200
    lvFilter.ColumnHeaders(lvFilter.ColumnHeaders.Count).Width = lv.ColumnHeaders(lv.ColumnHeaders.Count).Width
End Sub

Function totalColumn(ByVal colIndex As Long, Optional ByRef hadErr As Boolean) As Long
    On Error Resume Next
    Dim i As Long, tot As Long, li As ListItem
    
    colIndex = colIndex - 1 'we expect a 1 based index to be consistant
    
    If colIndex < 0 Or colIndex > currentLV.ColumnHeaders.Count Then
        hadErr = True
        Exit Function
    End If
    
    hadErr = False
    For Each li In currentLV.ListItems
        If colIndex = 0 Then
           If Len(li.Text) > 0 Then i = CLng(li.Text)
        Else
            If Len(li.subItems(colIndex)) > 0 Then i = CLng(li.subItems(colIndex))
        End If
        tot = tot + i
    Next
        
    hadErr = Not (Err.Number = 0)
    totalColumn = tot

End Function

Function avgColumn(ByVal colIndex As Long, Optional ByRef hadErr As Boolean) As Long
    On Error Resume Next
    Dim tot As Long
    tot = totalColumn(colIndex, hadErr)
    avgColumn = tot / currentLV.ListItems.Count
End Function

Public Sub SetLiColor(li As ListItem, newcolor As Long)
    Dim f As ListSubItem
    On Error Resume Next
    If li Is Nothing Then Exit Sub
    li.ForeColor = newcolor
    For Each f In li.ListSubItems
        f.ForeColor = newcolor
    Next
End Sub

Public Sub ColumnSort(Column As ColumnHeader)
    Dim ListViewControl As ListView
    On Error Resume Next
    
    Set ListViewControl = lv
    If lvFilter.Visible Then Set ListViewControl = lvFilter
        
    With ListViewControl
       If .SortKey <> Column.Index - 1 Then
             .SortKey = Column.Index - 1
             .SortOrder = lvwAscending
       Else
             If .SortOrder = lvwAscending Then
              .SortOrder = lvwDescending
             Else
              .SortOrder = lvwAscending
             End If
       End If
       .Sorted = -1
    End With
    
End Sub

Public Function GetAllElements(Optional selectedOnly As Boolean = False) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li As ListItem
    Dim ListViewControl As ListView
    Dim include  As Boolean
    
    On Error Resume Next
    
    Set ListViewControl = lv
    If lvFilter.Visible Then Set ListViewControl = lvFilter
        
    For i = 1 To ListViewControl.ColumnHeaders.Count
        tmp = tmp & ListViewControl.ColumnHeaders(i).Text & vbTab
    Next

    push ret, tmp
    push ret, String(50, "-")

    For Each li In ListViewControl.ListItems
    
        If selectedOnly Then
            If Not li.selected Then GoTo nextone
        End If
            
        tmp = li.Text & vbTab
        For i = 1 To ListViewControl.ColumnHeaders.Count - 1
            tmp = tmp & li.subItems(i) & vbTab
        Next
        push ret, tmp
        
nextone:
    Next

    GetAllElements = Join(ret, vbCrLf)

End Function

Function GetAllText(Optional subItemRow As Long = 0, Optional selectedOnly As Boolean = False) As String
    Dim i As Long
    Dim tmp() As String, x As String
    Dim ListViewControl As ListView
    
    On Error Resume Next
    
    Set ListViewControl = lv
    If lvFilter.Visible Then Set ListViewControl = lvFilter
    
    For i = 1 To ListViewControl.ListItems.Count
        If subItemRow = 0 Then
            x = ListViewControl.ListItems(i).Text
            If selectedOnly And Not ListViewControl.ListItems(i).selected Then x = Empty
            If Len(x) > 0 Then
                push tmp, x
            End If
        Else
            x = ListViewControl.ListItems(i).subItems(subItemRow)
            If selectedOnly And Not ListViewControl.ListItems(i).selected Then x = Empty
            If Len(x) > 0 Then
                push tmp, x
            End If
        End If
    Next
    
    GetAllText = Join(tmp, vbCrLf)
End Function

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Integer
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Private Function isIde() As Boolean
' Brad Martinez  http://www.mvps.org/ccrp
    On Error GoTo out
    Debug.Print 1 / 0
out: isIde = Err
End Function

Private Sub UserControl_Terminate()
    m_Locked = False
    Me.Clear
End Sub

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Private Function lpad(v, Optional L As Long = 8, Optional char As String = " ")
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < L Then
        lpad = String(L - x, char) & v
    Else
hell:
        lpad = v
    End If
End Function

Private Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function


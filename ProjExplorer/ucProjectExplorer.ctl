VERSION 5.00
Begin VB.UserControl ucProjectExplorer 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   ScaleHeight     =   552
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   -1920
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   1
      Top             =   960
      Width           =   615
   End
   Begin VB.Timer tMouseTimer 
      Enabled         =   0   'False
      Left            =   1800
      Top             =   4920
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7575
      Left            =   4560
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenDesigner 
         Caption         =   "Open Designer"
      End
      Begin VB.Menu mnuOpenCode 
         Caption         =   "Open Code Pane"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewGroup 
         Caption         =   "New Group"
      End
      Begin VB.Menu mnuEditGroup 
         Caption         =   "Rename Group"
      End
      Begin VB.Menu mnuDeleteGroup 
         Caption         =   "Delete Group"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSort 
         Caption         =   "Sort Groups"
      End
   End
End
Attribute VB_Name = "ucProjectExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
   X  As Long
   Y  As Long
End Type

Private Declare Function GetCapture Lib "user32" () As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOutW Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Const DT_LEFT As Long = &H0&
Const DT_VCENTER As Long = &H4&
Const DT_SINGLELINE As Long = &H20&
Const DT_END_ELLIPSIS As Long = &H8000&

Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal iStepIfAniCur As Long, ByVal hBrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Const DI_NORMAL As Long = &H3

Private Const TEXT_FLAGS As Long = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
Private Const GRP_HEADER As String = "*HEADER*"
Private Const ITEM_BACK_CLR As Long = vbWindowBackground, ITEM_BORDER_CLR As Long = vbScrollBars, ITEM_TEXT_CLR As Long = vbWindowText
Private Const ITEM_HOT_CLR As Long = vbInactiveTitleBar, SEL_CLR As Long = vbActiveTitleBar, SEL_BORDER_CLR As Long = vbHighlight

Private mGroups As Collection
Private mDragging As Boolean, mCopyMode As Boolean, mAutoDragDirection As Long
Private mMouseItem As String, mMouseGroup As String, mSelectedItem As String, mSelectedItemGroup As String
Private mMouseX As Long, mMouseY As Long, mDragOffsetX As Long, mDragOffsetY As Long, mMouseDownX As Long, mMouseDownY As Long
Private mViewportH As Long, mViewportW As Long, mViewportY As Long, mViewportMaxY As Long, mTotalH As Long

'Event MouseItemChanged(ByVal ItemKey As String, ByVal ItemCaption As String)
Event DoubleClick(ByVal ItemKey As String)
Event MouseDown(ByVal Button As MouseButtonConstants, ByVal X As Single, ByVal Y As Single, ByVal GroupKey As String, ByVal ItemKey As String, ByVal IsHeader As Boolean)
Public Sub SelectItem(ItemKey As String, Optional EnsureIsVisible As Boolean)
Dim i As Long, j As Long, thisGroup As clsGroup, Found As Boolean, thisItem As clsItem
   mSelectedItem = ItemKey
   For Each thisGroup In mGroups
      For Each thisItem In thisGroup.Items
         If thisItem.Key = mSelectedItem Then
            mSelectedItemGroup = thisGroup.Key
            Found = True 'could potentially find more than one, but the first one will do...
            Exit For
         End If
      Next thisItem
      If Found Then Exit For
   Next thisGroup
   If Found And EnsureIsVisible Then EnsureVisible thisItem.Key, thisGroup.Key
   Refresh
End Sub
Public Sub AddGroup(GroupName As String, Optional ByVal Silent As Boolean)
Dim NewGroup As New clsGroup
   If Exists(mGroups, GroupName) Then Exit Sub
   NewGroup.Expanded = True
   NewGroup.Key = GroupName
   mGroups.Add NewGroup, GroupName
   If Silent Then Exit Sub
   CreateGroupLayout
End Sub
Public Sub AddGroupItem(GroupKey As String, ItemKey As String, ItemCaption As String, ItemType As Long, IconHandle As Long, Optional ExpandGroup As Variant, Optional Silent As Boolean)
Dim W As Long
   AddGroup GroupKey, Silent
   If Not IsMissing(ExpandGroup) Then mGroups(GroupKey).Expanded = ExpandGroup
   UserControl.Font = "Tahoma"
   UserControl.FontBold = False
   UserControl.FontSize = 8
   W = UserControl.TextWidth(ItemCaption) + ICON_SPACE + 3 'a bit of space at the end
   mGroups(GroupKey).AddNewItem ItemKey, ItemCaption, ItemType, W, IconHandle, Silent
   If Not Silent Then CreateGroupLayout
End Sub
Public Sub EnsureVisible(ItemKey As String, GroupKey As String)
Dim thisGroup As clsGroup, thisItem As clsItem

   Set thisGroup = mGroups(GroupKey)
   Set thisItem = mGroups(GroupKey).Items(ItemKey)
   
   If Not thisGroup.Expanded Then DoDropEffect True
   
   If thisGroup.GroupY + thisItem.Top < mViewportY Then
      Scroll (thisGroup.GroupY + thisItem.Top) - mViewportY
   ElseIf thisGroup.GroupY + thisItem.Top + ITEM_H > mViewportY + mViewportH Then
      Scroll (thisGroup.GroupY + thisItem.Top + ITEM_H) - (mViewportY + mViewportH)
   End If
End Sub
Public Sub UnHideAll(Optional Silent As Boolean)
Dim thisGroup As clsGroup, thisItem As clsItem
   For Each thisGroup In mGroups
      For Each thisItem In thisGroup.Items
         thisItem.Hidden = False
      Next thisItem
   Next thisGroup
   If Not Silent Then CreateGroupLayout
End Sub
Public Sub HideItem(ItemKey As String, Optional Silent As Boolean)
Dim Group As clsGroup, Item As clsItem
   For Each Group In mGroups
      For Each Item In Group.Items
         If Item.Key = ItemKey Then Item.Hidden = True
      Next Item
   Next Group
   If Not Silent Then CreateGroupLayout
End Sub
Public Sub RemoveItem(ItemKey As String, Optional GroupKey As String, Optional Silent As Boolean) 'ItemKey can exist in more than one group
Dim Group As clsGroup
   If Len(GroupKey) Then
      Set Group = mGroups(GroupKey)
      Group.RemoveItem ItemKey, Silent
   Else
      For Each Group In mGroups
         Group.RemoveItem ItemKey, Silent
      Next Group
   End If
   If Not Silent Then CreateGroupLayout
End Sub
Public Sub RemoveGroup(GroupKey As String)
   mGroups.Remove GroupKey
   CreateGroupLayout
End Sub
Public Property Get ItemPresenceCount(ItemKey As String) As Long
Dim Group As clsGroup
   For Each Group In mGroups
      If Not Group.Item(ItemKey) Is Nothing Then ItemPresenceCount = ItemPresenceCount + 1
   Next Group
End Property
Public Sub MoveItem(ItemKey As String, FromGroupKey As String, ToGroupKey As String, Optional Silent As Boolean)
Dim Item As clsItem
   
   If FromGroupKey = ToGroupKey Then Exit Sub
   If Not GroupItem(ToGroupKey, ItemKey) Is Nothing Then Err.Raise vbObjectError + 600, , "Already exists in target"
   
   AddGroup ToGroupKey, True
   
   Set Item = GroupItem(FromGroupKey, ItemKey)
   
   mGroups(ToGroupKey).AddItem Item
   mGroups(FromGroupKey).RemoveItem ItemKey, Silent
   
   If Not Silent Then CreateGroupLayout
   
End Sub
Public Sub Sort()
Dim TmpColl As Collection, i As Long, SortedKeys() As String
   Set TmpColl = New Collection
   GetSortedCollectionKeys mGroups, SortedKeys
   For i = 0 To UBound(SortedKeys)
      TmpColl.Add mGroups(SortedKeys(i)), SortedKeys(i)
   Next i
   Set mGroups = TmpColl
   CreateGroupLayout
End Sub
Public Sub Clear()
   Set mGroups = New Collection
   UserControl_Resize
End Sub
Public Sub ExpandAll()
   DoDropEffect True
End Sub
Public Sub CollapseAll()
   DoDropEffect False
End Sub
Public Property Get ItemCount(Optional GroupKey As String) As Long
Dim Group As clsGroup
   If Len(GroupKey) Then
      ItemCount = mGroups(GroupKey).ItemCount
   Else
      For Each Group In mGroups
         ItemCount = ItemCount + Group.ItemCount
      Next Group
   End If
End Property
Public Property Get GroupItem(GroupKey As String, ItemKey As String) As clsItem
   On Error Resume Next
   Set GroupItem = mGroups(GroupKey).Items(ItemKey)
End Property
Public Property Get Groups() As Collection
   Set Groups = mGroups
End Property
Public Property Get GroupItems(GroupKey As String) As Collection
   Set GroupItems = mGroups(GroupKey).Items
End Property
Private Function Exists(pCollection As Collection, ByVal Key As Variant) As Boolean
   On Error Resume Next
   IsObject pCollection.Item(Key)
   Exists = Err.Number = 0
End Function
Public Sub Init()
Dim thisGroup As clsGroup
   For Each thisGroup In mGroups
      thisGroup.SortItems
   Next thisGroup
   CreateGroupLayout
End Sub
Private Sub HitTest(ByVal X As Single, ByVal Y As Single, HitGroup_OUT As String, Optional HitItem_OUT As Variant)
Dim thisGroup As clsGroup
   HitGroup_OUT = vbNullString: HitItem_OUT = vbNullString
   For Each thisGroup In mGroups
      If Y < thisGroup.GroupY + thisGroup.GroupHeight Then
         HitGroup_OUT = thisGroup.Key
         Exit For
      End If
   Next thisGroup
   
   If Not IsMissing(HitItem_OUT) Then
      If Len(HitGroup_OUT) Then HitItem_OUT = mGroups(HitGroup_OUT).HitTest(X, Y)
   End If
End Sub
Private Sub tMouseTimer_Timer()
Dim pt As POINTAPI, NewCopyMode As Boolean, H As Long

   If mDragging Then
      H = GetCapture
      If H <> UserControl.hWnd Then mDragging = False
   
      NewCopyMode = GetKeyState(vbKeyControl) < 0
      If mCopyMode <> NewCopyMode Then mCopyMode = NewCopyMode: Refresh
      If mAutoDragDirection <> 0 Then ShiftViewport ITEM_H * mAutoDragDirection, True
   Else
      If Not tMouseTimer.Enabled Then  'a mouse-enter, effectively...
         tMouseTimer.Interval = 40
         tMouseTimer.Enabled = True '... so enable the timer
      Else '[C2] we're testing for a mouse-leave
         GetCursorPos pt
         If Not WindowFromPoint(pt.X, pt.Y) = UserControl.hWnd Then
            mMouseItem = vbNullString
            tMouseTimer.Enabled = False '... and if it happened, we disable the timer
            If Not mnuOptions.Visible Then Refresh '...and refresh the control
         End If
      End If
   End If
End Sub
Private Sub UserControl_DblClick()
   If mMouseItem = GRP_HEADER Then
      DoDropEffect Not mGroups(mMouseGroup).Expanded, mMouseGroup
   Else
      If Len(mMouseItem) Then RaiseEvent DoubleClick(mMouseItem)
   End If
End Sub
Private Sub DoDropEffect(Expand As Boolean, Optional GroupKey As String)
Dim i As Double, thisGroup As clsGroup
   
   mMouseItem = vbNullString 'forget the mouse-over item
   
   For Each thisGroup In mGroups
      If thisGroup.Key = GroupKey Or Len(GroupKey) = 0 Then
         If Expand And Not thisGroup.Expanded Then thisGroup.YEffectPct = 0 Else thisGroup.YEffectPct = 1
         If Expand Then thisGroup.Expanded = True
      End If
   Next thisGroup

   i = 0.025
   Do
      For Each thisGroup In mGroups
         If thisGroup.Key = GroupKey Or Len(GroupKey) = 0 Then
            If Expand Then
               If thisGroup.YEffectPct < 1 Then thisGroup.YEffectPct = i
            Else
               If thisGroup.YEffectPct > 0 Then thisGroup.YEffectPct = 1 - i
            End If
         End If
      Next thisGroup
      CreateGroupLayout IIf(Expand, GroupKey, vbNullString)
      i = i * 2
   Loop Until i > 1

   For Each thisGroup In mGroups
      If thisGroup.Key = GroupKey Or Len(GroupKey) = 0 Then
         thisGroup.YEffectPct = 1
         If Not Expand Then thisGroup.Expanded = False
      End If
   Next thisGroup
   CreateGroupLayout IIf(Expand, GroupKey, vbNullString)
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim HitGroup As String, HitItem As String

   HitTest X, Y + mViewportY, HitGroup, HitItem
   mMouseItem = HitItem: mMouseGroup = HitGroup
   
   If Len(HitItem) Then mSelectedItem = HitItem: mSelectedItemGroup = HitGroup

   Refresh
   
   RaiseEvent MouseDown(Button, X, Y, HitGroup, IIf(HitItem = GRP_HEADER, vbNullString, HitItem), HitItem = GRP_HEADER)

   If Button = vbLeftButton Then
      If mDragging Then Exit Sub
      If mMouseItem = GRP_HEADER And X < 18 Then UserControl_DblClick
      mMouseDownX = X: mMouseDownY = Y
   End If
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim HitItem As String, HitGroup As String, DragDistance As Long

   Y = Y + mViewportY
   If Not tMouseTimer.Enabled Then tMouseTimer_Timer 'turn on mouse-leave checking
   
   If Button = 0 Then
      HitTest X, Y, HitGroup, HitItem
      
      If mMouseItem <> HitItem Or HitGroup <> mMouseGroup Then
         mMouseItem = HitItem
         'RaiseEvent MouseItemChanged(mMouseItem, vbNullString)
         mMouseGroup = HitGroup
         Refresh
      End If
   ElseIf Button = vbLeftButton And Len(mMouseItem) Then
      If mDragging Then
         mMouseX = X - mDragOffsetX: mMouseY = Y - mDragOffsetY
         Refresh
      Else
         If Abs(X - mMouseDownX) > 4 Or Abs(Y - mMouseDownY) > 4 Then
            If mMouseItem = GRP_HEADER Then
               If mGroups(mMouseGroup).Expanded Then DoDropEffect False, mMouseGroup
               mMouseItem = GRP_HEADER
               mDragOffsetX = X
               mDragOffsetY = Y - mGroups(mMouseGroup).GroupY
            Else
               mDragOffsetX = X - mGroups(mMouseGroup).Item(mMouseItem).Left
               mDragOffsetY = Y - mGroups(mMouseGroup).Item(mMouseItem).Top - mGroups(mMouseGroup).GroupY
            End If
            mDragging = True
            mCopyMode = Shift And vbCtrlMask
         End If
      End If
   End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim HitGroupKey As String
   
   If Not mDragging Then Exit Sub
   
   mDragging = False
   HitTest X, Y + mViewportY, HitGroupKey
   
   If Len(HitGroupKey) = 0 Then HitGroupKey = mGroups(mGroups.Count).Key
   
   If HitGroupKey = mSelectedItemGroup Then
      Refresh
   Else
      If mSelectedItem = GRP_HEADER Then DoHeaderEndDrag HitGroupKey Else DoItemEndDrag HitGroupKey
      CreateGroupLayout
   End If
End Sub
Private Sub DoItemEndDrag(TargetGroupKey As String)
Dim TargetGroup As clsGroup, DraggedItem As clsItem

   On Error GoTo ErrHandler:

   Set TargetGroup = mGroups(TargetGroupKey)
   
   Set DraggedItem = GroupItem(mSelectedItemGroup, mSelectedItem)
   If mCopyMode Then
      AddGroupItem TargetGroupKey, DraggedItem.Key, DraggedItem.Caption, DraggedItem.ItemType, DraggedItem.IconHandle
   Else
      MoveItem mSelectedItem, mSelectedItemGroup, TargetGroupKey
   End If
   mSelectedItemGroup = TargetGroupKey
   mCopyMode = False

   Exit Sub
ErrHandler:
   If Err.Number = -2147220904 Then
      Refresh
      If TargetGroupKey <> mMouseGroup Then
         MsgBox "The Group '" & TargetGroup.Key & "' already contains" & vbCrLf & "the component '" & mMouseItem & "'", , "Invalid Drop Target"
      End If
   Else
      MsgBox Err.Description
   End If
End Sub
Private Sub DoHeaderEndDrag(TargetGroupKey As String)
Dim SourceGroup As clsGroup, MoveToEnd As Boolean

   MoveToEnd = TargetGroupKey = mGroups(mGroups.Count).Key

   Set SourceGroup = mGroups(mSelectedItemGroup)
   mGroups.Remove mSelectedItemGroup
   If MoveToEnd Then
      mGroups.Add SourceGroup, mSelectedItemGroup, , TargetGroupKey
   Else
      mGroups.Add SourceGroup, mSelectedItemGroup, TargetGroupKey
   End If

End Sub
Private Sub CreateGroupLayout(Optional GroupToKeepVisible As String)
Dim thisGroup As clsGroup, Key As String, NewMaxY As Long
   
   mViewportW = UserControl.ScaleWidth - IIf(VScroll1.Enabled, VScroll1.Width, 0)
   mTotalH = 0
   For Each thisGroup In mGroups
      thisGroup.SetGroupWidth mViewportW
      thisGroup.GroupY = mTotalH
      mTotalH = mTotalH + thisGroup.GroupHeight
   Next thisGroup

   If mTotalH > mViewportH And Not VScroll1.Enabled Then
      VScroll1.Enabled = True
      mViewportW = UserControl.ScaleWidth - VScroll1.Width
      CreateGroupLayout GroupToKeepVisible
      Exit Sub
   ElseIf mTotalH < mViewportH And VScroll1.Enabled Then
      VScroll1.Enabled = False
      mViewportW = UserControl.ScaleWidth
      CreateGroupLayout GroupToKeepVisible
      Exit Sub
   End If
   VScroll1.Visible = VScroll1.Enabled
   NewMaxY = mTotalH - mViewportH
   If NewMaxY < 0 Then NewMaxY = 0

   If Len(GroupToKeepVisible) Then 'when we're expanding a group, we do our best to keep the whole thing visible
      Set thisGroup = mGroups(GroupToKeepVisible)
      If thisGroup.GroupY + thisGroup.GroupHeight > mViewportH + mViewportY Then
         mViewportY = mViewportY - (mViewportMaxY - NewMaxY)
         If mViewportY < 0 Then mViewportY = 0
         If mViewportY > thisGroup.GroupY Then mViewportY = thisGroup.GroupY
      End If
   End If

   If mViewportY > NewMaxY Then mViewportY = NewMaxY
   mViewportMaxY = NewMaxY
      
   VScroll1.Tag = "Busy"
      'If VScroll1.Value > mViewportMaxY Then VScroll1.Value = mViewportMaxY
      VScroll1.Max = mViewportMaxY
      VScroll1.Value = mViewportY
      VScroll1.LargeChange = mViewportH
      VScroll1.SmallChange = ITEM_H + SPACING
   VScroll1.Tag = vbNullString
   Refresh
End Sub
'==========================================
'SCROLLING
Private Sub VScroll1_Change()
   If VScroll1.Tag = "Busy" Then Exit Sub
   If Abs(VScroll1.Value - mViewportY) <= VScroll1.SmallChange Then
      ShiftViewport VScroll1.Value
   Else
      Scroll VScroll1.Value - mViewportY
   End If
End Sub
Private Sub VScroll1_Scroll()
   ShiftViewport VScroll1.Value, False
End Sub
Private Sub ShiftViewport(NewY As Long, Optional Relative As Boolean)
   
   If Not mDragging Then mMouseItem = vbNullString
   If mDragging Then mMouseY = mMouseY + NewY
   
   If Relative Then NewY = mViewportY + NewY
   If NewY < 0 Then NewY = 0
   If NewY > mViewportMaxY Then NewY = mViewportMaxY
   mViewportY = NewY
   VScroll1.Tag = "Busy" 'ignore internal changes to the value
      VScroll1.Value = NewY
      Refresh
   VScroll1.Tag = vbNullString
End Sub
Private Sub Scroll(ScrollDistance As Long)
Dim DX As Double, prevDX As Double, i As Double, TotalDX As Double, prevTotalDX As Double
Const STEPS As Long = 8
   DX = 1
   For i = -1 + 1 / STEPS To 1 Step 1 / STEPS
      prevDX = DX
      prevTotalDX = TotalDX
      DX = Sqr(1 - (1 - Abs(i)) ^ 2)
      TotalDX = TotalDX + Abs(prevDX - DX) * ScrollDistance / 2
      ShiftViewport CLng(TotalDX) - CLng(prevTotalDX), True
   Next i
End Sub
'===============================================
Private Sub UserControl_Initialize()
   UserControl.FillStyle = 0
   VScroll1.Enabled = False
   Set mGroups = New Collection
End Sub
Private Sub UserControl_Resize()
   If mViewportW = 0 Then mViewportW = UserControl.ScaleWidth
   mViewportH = UserControl.ScaleHeight
   VScroll1.Move UserControl.ScaleWidth - VScroll1.Width, 0, 18, mViewportH
   CreateGroupLayout
End Sub
Private Sub UserControl_Terminate()
   Set mGroups = Nothing
End Sub
'=================================
'Drawing stuff
Public Sub Refresh() 'drawing in 3 loops is an optimisation (swapping fonts in-and-out is slower)
Dim thisGroup As clsGroup, thisItem As clsItem, HasMouseOver As Boolean, IsVisible As Boolean
Dim X As Long, Y As Long, R As RECT

   mCopyMode = GetKeyState(vbKeyControl) < 0
'New_c.Timing True
   With UserControl
      .AutoRedraw = True
      .Cls
      .DrawStyle = 0
      .Font = "Tahoma"
      .FontSize = 8
      .FontBold = False
      .FillColor = ITEM_BACK_CLR
      .ForeColor = ITEM_BORDER_CLR
      SetTextColor .hdc, ITEM_TEXT_CLR
      
      'Items
      For Each thisGroup In mGroups
         IsVisible = thisGroup.GroupY + thisGroup.GroupHeight >= mViewportY And thisGroup.GroupY <= mViewportY + mViewportH
         If thisGroup.Expanded And IsVisible Then
            For Each thisItem In thisGroup.Items
               IsVisible = thisGroup.GroupY + thisItem.Top + ITEM_H >= mViewportY And thisGroup.GroupY + thisItem.Top * thisGroup.YEffectPct <= mViewportY + mViewportH
               If IsVisible And Not thisItem.Hidden Then
                  If thisGroup.Key = mSelectedItemGroup And thisItem.Key = mSelectedItem Then 'item is selected
                     If mDragging Then
                        .FillColor = IIf(mCopyMode, ITEM_BACK_CLR, .BackColor)
                        .DrawStyle = IIf(mCopyMode, 0, 2)
                        .ForeColor = IIf(mCopyMode, ITEM_BORDER_CLR, SEL_BORDER_CLR)
                     Else
                        .FillColor = SEL_CLR
                        .ForeColor = IIf(mDragging, ITEM_BORDER_CLR, SEL_BORDER_CLR)
                     End If
                     SetTextColor .hdc, ITEM_TEXT_CLR
                     DrawItem thisItem, thisItem.Left, thisGroup.GroupY - mViewportY + thisItem.Top * thisGroup.YEffectPct, mDragging And Not mCopyMode
                     .ForeColor = ITEM_BORDER_CLR 'restore
                     .FillColor = ITEM_BACK_CLR 'restore
                     .DrawStyle = 0 'restore
                     SetTextColor .hdc, ITEM_TEXT_CLR 'restore
                  Else
                     If thisGroup.Key = mMouseGroup And thisItem.Key = mMouseItem Then 'item has mouse over
                        .FillColor = ITEM_HOT_CLR
                        DrawItem thisItem, thisItem.Left, thisGroup.GroupY - mViewportY + thisItem.Top * thisGroup.YEffectPct, False
                        .FillColor = ITEM_BACK_CLR 'restore
                     Else 'normal draw
                        DrawItem thisItem, thisItem.Left, thisGroup.GroupY - mViewportY + thisItem.Top * thisGroup.YEffectPct, False
                     End If
                  End If
               End If
            Next thisItem
         End If
      Next thisGroup
      
      'Headers
      .FontBold = True
      For Each thisGroup In mGroups
         If thisGroup.GroupY + HEADER_H >= mViewportY And thisGroup.GroupY <= mViewportY + mViewportH Then
            HasMouseOver = thisGroup.Key = mMouseGroup And mMouseItem = GRP_HEADER
            .FillColor = IIf(HasMouseOver, vbWindowFrame, vbGrayText)
            .ForeColor = .FillColor
            .DrawStyle = IIf(mDragging And HasMouseOver, 2, 0)
            SetTextColor .hdc, vbWhite
            DrawGroupHeader thisGroup, thisGroup.GroupY - mViewportY, mDragging And HasMouseOver
         End If
      Next thisGroup
      .FontBold = False
      
      'Expand/Collapse Glyphs
      .Font = "Marlett": .FontSize = 12
      For Each thisGroup In mGroups
         If thisGroup.GroupY + HEADER_H >= mViewportY And thisGroup.GroupY <= mViewportY + mViewportH Then
            R.Left = 2: R.Top = thisGroup.GroupY - mViewportY: R.Right = mViewportW: R.Bottom = R.Top + HEADER_H
            HasMouseOver = thisGroup.Key = mMouseGroup And mMouseItem = GRP_HEADER
            If Not (HasMouseOver And mDragging) Then DrawText .hdc, IIf(thisGroup.Expanded, "6", "4"), -1, R, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
         End If
      Next thisGroup
      
      'dragged object (item or header)
      If mDragging Then
         .Font = "Tahoma": .FontSize = 8
         Set thisGroup = mGroups(mMouseGroup)
         X = mMouseX: Y = mMouseY - mViewportY
         If Y < 0 Then
            Y = 0: mAutoDragDirection = -1
         ElseIf Y > mViewportH - ITEM_H Then
            Y = mViewportH - ITEM_H: mAutoDragDirection = 1
         Else
            mAutoDragDirection = 0
         End If
      
         If mMouseItem = GRP_HEADER Then
            .FillColor = vbWindowFrame
            .ForeColor = .FillColor
            .FontBold = True
            .DrawStyle = 0
            SetTextColor .hdc, vbWhite
            DrawGroupHeader thisGroup, Y, False
            .Font = "Marlett": .FontSize = 12
            R.Left = 2: R.Top = Y - mViewportY: R.Right = mViewportW: R.Bottom = R.Top + HEADER_H
            DrawText .hdc, "4", -1, R, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
         Else
            If X > mViewportW - thisGroup.Item(mMouseItem).Width Then X = mViewportW - thisGroup.Item(mMouseItem).Width
            If X < 0 Then X = 0
            .FillColor = SEL_CLR
            .ForeColor = SEL_BORDER_CLR
            SetTextColor .hdc, ITEM_TEXT_CLR
            DrawItem thisGroup.Item(mMouseItem), X, Y, False
            .FontBold = True
            .FontSize = 14
            SetTextColor .hdc, SEL_BORDER_CLR
            If mCopyMode Then TextOutW .hdc, X - 4 + thisGroup.Item(mMouseItem).Width, Y, StrPtr("+"), 1
         End If
      End If
      
      .AutoRedraw = False
   End With
'Debug.Print New_c.Timing
End Sub
Private Sub DrawGroupHeader(Group As clsGroup, Y As Long, OutlineOnly As Boolean)
Dim R As RECT
   With UserControl
      If OutlineOnly Then .FillColor = .BackColor
      R.Top = Y: R.Right = mViewportW: R.Bottom = R.Top + HEADER_H
      Rectangle .hdc, 2, R.Top + 1, mViewportW - 2, R.Bottom - 2
      
      If OutlineOnly Then Exit Sub
      
      R.Left = R.Left + ICON_SPACE + 2
      DrawText .hdc, Group.Key & " (" & Group.VisibleItemCount & ")", -1, R, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
   End With
End Sub
Private Sub DrawItem(Item As clsItem, X As Long, Y As Long, OutlineOnly As Boolean)
   With UserControl
      Rectangle .hdc, X, Y, X + Item.Width, Y + ITEM_H
      
      If OutlineOnly Then Exit Sub
      
      DrawIconEx .hdc, X + 2, Y + 1, Item.IconHandle, 16, 16, 0, 0, DI_NORMAL
      TextOutW .hdc, X + ICON_SPACE, Y + 2, StrPtr(Item.Caption), Len(Item.Caption)
   End With
End Sub


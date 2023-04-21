VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GroupY As Long
Public Key As String
Public Expanded As Boolean
Public YEffectPct As Double 'used for the FX associated with expanding/collapsing nodes

Private mGroupHeight As Long, mGroupWidth As Long
Private mVisibleCount As Long
Private mItems As Collection
Private Sub Class_Initialize()
   YEffectPct = 1
   Set mItems = New Collection
End Sub
Private Sub Class_Terminate()
   Set mItems = Nothing
End Sub
Public Property Get Items() As Collection
   Set Items = mItems
End Property
Public Sub AddNewItem(ByVal ItemKey As String, ByVal ItemCaption As String, ByVal ItemType As Long, ByVal ItemWidth As Long, ByVal IconHandle As Long, Optional ByVal Silent As Boolean)
Dim NewItem As clsItem
   Set NewItem = New clsItem
   NewItem.Key = ItemKey
   NewItem.Caption = ItemCaption
   NewItem.ItemType = ItemType
   NewItem.IconHandle = IconHandle
   NewItem.Width = ItemWidth
   NewItem.Height = ITEM_H
   mItems.Add NewItem, ItemKey
   If Silent Then Exit Sub
   SortItems
End Sub
Public Sub AddItem(Item As clsItem)
   mItems.Add Item, Item.Key
   SortItems
End Sub
Public Sub RemoveItem(Key As String, Optional Silent As Boolean) 'if silent=true, Init needs to be called on the UserControl to 'catch-up'!
   If Not Exists(mItems, Key) Then Exit Sub
   mItems.Remove Key
   If Not Silent Then SortItems
End Sub
Public Property Get Item(Key As String) As clsItem
   On Error Resume Next
   Set Item = mItems(Key)
End Property
Public Property Get ItemCount() As Long
   ItemCount = mItems.Count
End Property
Public Property Get VisibleItemCount() As Long
   VisibleItemCount = mVisibleCount 'mItems.Count
End Property
Public Function SetGroupWidth(Width As Long) As Boolean
   SetGroupWidth = mGroupWidth <> Width
   If Not SetGroupWidth Then Exit Function
   mGroupWidth = Width
   CreateItemlayout
End Function
Public Sub SortItems()
Dim TmpColl As Collection, i As Long, SortedKeys() As String
   If mItems.Count > 0 Then
      Set TmpColl = New Collection
      GetSortedCollectionKeys mItems, SortedKeys
      For i = 0 To UBound(SortedKeys)
         TmpColl.Add mItems(SortedKeys(i)), SortedKeys(i)
      Next i
      Set mItems = TmpColl
   End If
   CreateItemlayout
End Sub
Public Property Get GroupHeight() As Long
   GroupHeight = HEADER_H
   If Expanded Then GroupHeight = GroupHeight + (mGroupHeight - HEADER_H) * YEffectPct
End Property
Public Sub CreateItemlayout()
Dim X As Long, Y As Long
Dim Index As Long, LineStartIndex As Long, j As Long, ItemW As Long
   
   mVisibleCount = 0
   If mItems.Count > 0 Then
      LineStartIndex = 1
      Y = HEADER_H + SPACING
      'because, on each new line, we go back and fix the previous line, the following can only possibly occur once
      '(if at all) and so has been optimised out of the loop
      If mItems(1).Width > mGroupWidth And Not mItems(1).Hidden Then Y = Y - ITEM_H - SPACING
      
      For Index = 1 To mItems.Count
         With mItems(Index)
            If .Hidden Then
               ItemW = 0
            Else
               ItemW = .Width
               mVisibleCount = mVisibleCount + 1
            End If
            
            If X + ItemW > mGroupWidth - SPACING * 2 Then
               j = Index - 1
               Do Until j < LineStartIndex
                  mItems(j).Left = mItems(j).Left + (mGroupWidth - X + SPACING) * 0.5
                  j = j - 1
               Loop

               LineStartIndex = Index
               X = 0: Y = Y + ITEM_H + SPACING
            End If

            .Left = X: .Top = Y
            X = X + ItemW + IIf(ItemW > 0, SPACING, 0)

         End With
      Next Index
      j = mItems.Count
      Do Until j < LineStartIndex
         mItems(j).Left = mItems(j).Left + (mGroupWidth - X + SPACING) * 0.5
         j = j - 1
      Loop
         
      If mVisibleCount > 0 Then mGroupHeight = Y + ITEM_H + SPACING Else mGroupHeight = HEADER_H
   Else
      mGroupHeight = HEADER_H
   End If

End Sub
Public Function HitTest(X As Single, Y As Single) As String
Dim i As Long, j As Long
   
   Y = Y - GroupY
   
   If Y <= ITEM_H Then HitTest = "*HEADER*": Exit Function
   
   If mItems.Count = 0 Then Exit Function
   
   j = SearchStartIndex(Y - ITEM_H, 1, mItems.Count)
   For i = j To mItems.Count
      With mItems(i)
         If .Top > Y Then
            Exit For
         ElseIf Not .Hidden Then
            If .Left < X And .Left + .Width > X And .Top < Y And .Top + .Height > Y Then
               HitTest = mItems(i).Key
               Exit For
            End If
         End If
      End With
   Next i

End Function
Private Function SearchStartIndex(ByVal Target As Long, ByVal pFrom As Long, ByVal pTo As Long, Optional ByRef Position As Long = -1) As Long
Dim i As Long
   
   i = pFrom + Abs(pTo - pFrom) \ 2
   
   If mItems(i).Top = Target Then
      Position = pTo
   ElseIf i = pFrom Then
      Position = pFrom
   ElseIf mItems(i).Top > Target Then
      SearchStartIndex Target, pFrom, i, Position
   ElseIf mItems(i).Top < Target Then
      SearchStartIndex Target, i, pTo, Position
   End If
   
   SearchStartIndex = Position
End Function
Private Function Exists(pCollection As Collection, ByVal Key As Variant) As Boolean
   On Error Resume Next
   IsObject pCollection.Item(Key)
   Exists = Err.Number = 0
End Function

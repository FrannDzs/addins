Attribute VB_Name = "modGlobals"
Option Explicit

Public gVBInstance As VBIDE.VBE
Public FormDisplayed As Boolean

Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Src As Any, ByRef Dst As Any) As Long
Private Declare Function SysAllocString Lib "oleaut32" (ByRef pOlechar As Any) As Long

Public Type RECT
   Left   As Long
   Top    As Long
   Right  As Long
   Bottom As Long
End Type

Public Const HEADER_H As Long = 20
Public Const ITEM_H As Long = 18
Public Const SPACING As Long = 3
Public Const ICON_SPACE As Long = 20

' Thanks to The Trick for the following technique - slightly modded as i only want the keys in ascending order
' // Returns the sorted keys and items
Public Function GetSortedCollectionKeys(ByVal cCol As Collection, ByRef sKeys() As String) As Long
Dim pCur As Long, pNull As Long, lMask As Long, lIndex As Long
    
   If cCol.Count = 0 Then Exit Function
   
   ReDim sKeys(cCol.Count - 1)
   
   GetMem4 ByVal ObjPtr(cCol) + &H24, pCur
   GetMem4 ByVal ObjPtr(cCol) + &H28, pNull
   
   Fill pCur, pNull, lIndex, sKeys
   
   GetSortedCollectionKeys = lIndex
    
End Function
Private Function Fill(ByVal pItem As Long, ByVal pNull As Long, ByRef lIndex As Long, ByRef sKeys() As String)
Dim pKey As Long, pLeft As Long, pRight As Long
    
   If pItem = pNull Or pItem = 0 Then Exit Function
   
   GetMem4 ByVal pItem + (&H28), pLeft
   
   If pLeft <> pNull Then Fill pLeft, pNull, lIndex, sKeys

   ' // Extract key
   GetMem4 ByVal pItem + &H10, pKey
   GetMem4 SysAllocString(ByVal pKey), ByVal VarPtr(sKeys(lIndex))
   
   lIndex = lIndex + 1
   
   GetMem4 ByVal pItem + (&H24), pRight
   
   If pRight = pNull Then
       Exit Function
   Else
       Fill pRight, pNull, lIndex, sKeys
   End If

End Function




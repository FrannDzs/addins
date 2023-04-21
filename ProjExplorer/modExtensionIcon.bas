Attribute VB_Name = "modExtensionIcon"
Option Explicit

Private Const SHGFI_ICON As Long = &H100&
Private Const SHGFI_SMALLICON As Long = &H1&  '16x16 pixels.
Private Const SHGFI_USEFILEATTRIBUTES As Long = &H10&

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type

Private Type PictDesc_Icon
    cbSizeofStruct As Long
    picType As Long
    hIcon As Long
End Type

Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoW" ( _
    ByVal pszPath As Long, _
    ByVal dwFileAttributes As Long, _
    ByVal psfi As Long, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long
Public Function GetAssocIconHandle(ByVal Extension As String) As Long
Dim SFI As SHFILEINFO, Desc As PictDesc_Icon
    
    If SHGetFileInfo(StrPtr(Extension), 0, VarPtr(SFI), LenB(SFI), SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_USEFILEATTRIBUTES) = 0 Then Exit Function
    
    With Desc
       .cbSizeofStruct = Len(Desc)
       .picType = vbPicTypeIcon
       .hIcon = SFI.hIcon
    End With

    GetAssocIconHandle = Desc.hIcon

End Function





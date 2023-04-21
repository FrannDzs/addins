Attribute VB_Name = "modIniFile"
Option Explicit

Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpReturnBuffer As String, ByVal nSize As Long, ByVal lpName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Function ReadFromINI(Filename As String, ByVal Section As String, ByVal Key As String, Optional ByVal Default As String) As String
Dim Buffer As String
   Buffer = String$(750, Chr$(0&))
   ReadFromINI$ = Left$(Buffer, GetPrivateProfileString(Section, ByVal LCase$(Key), Default, Buffer, Len(Buffer), Filename))
End Function
Public Sub WriteToINI(Filename As String, ByVal Section As String, ByVal Key As String, ByVal KeyValue As String)
   WritePrivateProfileString Section, Key, KeyValue, Filename
End Sub
Public Sub DeleteIniSection(Filename As String, ByVal Section As String)
   WritePrivateProfileString Section, 0&, 0&, Filename
End Sub
Public Sub DeleteIniKey(Filename As String, ByVal Section As String, ByVal KeyName As String)
   WritePrivateProfileString Section, KeyName, 0&, Filename
End Sub
Public Function CheckIfIniKeyExists(Filename As String, ByVal Section, ByVal KeyName As String) As Boolean
Dim str_A As String, str_B As String
   str_A = ReadFromINI(Filename, Section, KeyName, "A")
   str_B = ReadFromINI(Filename, Section, KeyName, "B")
   If str_A = str_B Then CheckIfIniKeyExists = True
End Function
Public Function CheckIfIniSectionExists(Filename As String, ByVal Section As String) As Boolean
Dim Buffer As String
   Buffer = String$(750, Chr$(0&))
   CheckIfIniSectionExists = CBool(GetPrivateProfileSection(Section, Buffer, Len(Buffer), Filename) > 0)
End Function
Public Function GetLongFromINI(Filename As String, ByVal Section, ByVal KeyName As String, Optional ByVal Default As Long) As Long
   GetLongFromINI = GetPrivateProfileInt(Section, KeyName, Default, Filename)
End Function
'Public Function GetSectionNames(Filename As String) As String()
'Dim sSections As String * 1000, lngResult As Long, sTmp() As String
'
'    lngResult = GetPrivateProfileSectionNames(sSections, Len(sSections), Filename)
'    sTmp = Split(sSections, vbNullChar & vbNullChar)
'    GetSectionNames = Split(sTmp(0), vbNullChar)
'End Function
'Public Function GetSectionItems(Filename As String, Section As String) As String()
'Dim lngResult As Long, sItems As String * 5000, sTmp() As String
'
'   lngResult = GetPrivateProfileSection(Section, sItems, Len(sItems), Filename)
'   sTmp = Split(sItems, vbNullChar & vbNullChar)
'   GetSectionItems = Split(sTmp(0), vbNullChar)
'End Function

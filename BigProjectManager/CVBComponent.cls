VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVBComponent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'todo: cleanup - left over from when this was external and done over IPC


Public Typ  As Long
Public name As String
Public path As String
Private m_raw As String

Public icon As String
Public defFolder As String
Public n As Node

Property Get raw() As String
    'if there was a rename we need to rebuild it..
    raw = Typ & "|" & name & "|" & path
End Property

Function loadFromList(raw)
    Dim yy() As String, icn As String
    
    m_raw = raw 'why save it now..we need to rebuild it on demand anyway...
    yy = Split(raw, "|") 'type|name|path  path may not exist yet if not saved...
    Typ = yy(0)
    path = yy(2)
    
    If Typ = &HA Then 'related document no name
        name = fso.FileNameFromPath(path)
    Else
        name = yy(1)
    End If
    
    defFolder = DefaultFolderForType(Typ, icn)
    icon = icn
        
End Function

Function loadFromEvent(e As CComponentEvent)
    Dim icn As String
    
    Typ = e.ComponentType
    name = e.ComponentName
    path = e.filename
    
    If Typ = &HA Then 'related document no name
        name = fso.FileNameFromPath(path)
    End If
    
    defFolder = DefaultFolderForType(Typ, icn)
    icon = icn
    
End Function

Function loadFromFile(ByVal fpath As String) As Boolean
    On Error Resume Next
    Dim icn As String
    path = fpath
    Typ = typeFromPath(fpath)
    name = PreloadComponentName(fpath)
    
    'If Typ = &HA Then 'related document no name
    '    name = fso.FileNameFromPath(path)
    'End If
    
    defFolder = DefaultFolderForType(Typ, icn)
    icon = icn
    loadFromFile = (Typ <> 0)
End Function

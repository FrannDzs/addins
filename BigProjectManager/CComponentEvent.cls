VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CComponentEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'todo: cleanup - left over from when this was external and done over IPC

    'ipc.Send "Component|Renamed|" & c.Collection.Parent.name & "|" & c.Type & "|" & c.name & "|" & c.FileNames(1) & "|" & OldName
    'ipc.Send "Component|Selected|" & c.Collection.Parent.name & "|" & c.Type & "|" & c.name & "|" & c.FileNames(1)
    'ipc.Send "Component|Removed|" & c.Collection.Parent.name & "|" & c.Type & "|" & c.name & "|" & c.FileNames(1)
    'ipc.Send "Component|Added|" & c.Collection.Parent.name & "|" & c.Type & "|" & c.name & "|" & c.FileNames(1)

Enum eComponentEvents
    ec_Unknown = 0
    ec_Rename = 1
    ec_Select = 2
    ec_Remove = 3
    ec_Add = 4
End Enum

Public raw As String
Public EventName As String
Public EventType As eComponentEvents
Public ParentName As String
Public ComponentType As Long
Public ComponentName As String
Public filename As String
Public OldName As String

Sub init(raw As String)

    On Error Resume Next
    Dim x() As String
    
    Me.raw = raw
    x = Split(raw, "|")
    EventName = x(1)
    
    Select Case EventName
        Case "Renamed": EventType = ec_Rename
        Case "Selected": EventType = ec_Select
        Case "Removed": EventType = ec_Remove
        Case "Added": EventType = ec_Add
   End Select
   
    ParentName = x(2)
    ComponentType = x(3)
    ComponentName = x(4)
    filename = x(5)
    OldName = x(6) 'only exists for rename
    
End Sub


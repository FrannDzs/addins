VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSaveTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Serialize(n As Node, ByRef appendTag As String, ByVal index As Long)
Event DeSerialize(n As Node, ByVal appendTag As String, ByVal index As Long)

Private m_baseDir As String
Private f As Integer
Private intIndentLevel As Integer
Private index As Long
Private maxDepth As Long

Public NodeCount As Long 'mainly for RestoreTree clients who might want to use a progress bar

Property Get BaseDir() As String
    BaseDir = m_baseDir
End Property

Function saveTree(tv As TreeView, fpath As String)
 
     Dim objNode As Node
     Dim cfg() As Byte
     
     m_baseDir = GetParentFolder(fpath) & "\"
     If FileExists(fpath) Then Kill fpath
     
     index = 0
     maxDepth = 0
     intIndentLevel = 0
     NodeCount = tv.Nodes.Count
      
     Set objNode = tv.Nodes(1)
     
     f = FreeFile
     Open fpath For Output As #f
     Print #f, "# " & Space(100) 'leave room to embed config data after generated
     SaveTree2 objNode
     Close #f
     
     cfg() = StrConv("MaxDepth:" & maxDepth & " NodeCount:" & tv.Nodes.Count, vbFromUnicode, &H409)
     Open fpath For Binary As #f
     Put f, 3, cfg()
     Close #f
     
 End Function
 
'recursive
Private Sub SaveTree2(objNode As Node)
        
        Dim tag As String
        
        RaiseEvent Serialize(objNode, tag, index)
        tag = Replace(tag, vbTab, "\t")
        tag = Replace(tag, vbCr, "\r")
        tag = Replace(tag, vbLf, "\n")
        Print #f, String(intIndentLevel, vbTab) & Replace(objNode.Text, ":", Chr(5)) & ":" & tag
        index = index + 1
        
        If objNode.Children > 0 Then
            intIndentLevel = intIndentLevel + 1
            If intIndentLevel > maxDepth Then maxDepth = intIndentLevel
            SaveTree2 objNode.Child
        End If
        
        Set objNode = objNode.Next
       
        If TypeName(objNode) <> "Nothing" Then
            SaveTree2 objNode
        Else
            intIndentLevel = intIndentLevel - 1
        End If
        
End Sub

Function RestoreTree(tv As TreeView, fpath As String) As Long
    
     Dim n As Node, pn As Node
     Dim parentNodes() As Node  'i = level, up to 10 levels supported by default
     Dim x As String, tabs As Long
     Dim tag As String
      
     m_baseDir = GetParentFolder(fpath) & "\"
     If Not FileExists(fpath) Then Exit Function
     
     index = 0
     NodeCount = 0 'extracted from first cfg line (next also)
     maxDepth = 20
     ReDim parentNodes(maxDepth)
     tv.Nodes.Clear
     
     f = FreeFile
     Open fpath For Input As #f
    
     Do While Not EOF(f)
        
        Line Input #f, x
        If Left(x, 1) = "#" Then 'its a comment line extract config details
            extractcfg x 'expected to set NodesCount and MaxDepth
            If maxDepth > UBound(parentNodes) - 1 Then ReDim parentNodes(maxDepth + 5)
        Else
            tabs = ReadTreeLine(x, tag)
            
            If tabs = 0 Then 'new top level node
                Set n = tv.Nodes.Add(, , , x)
            Else
                Set pn = parentNodes(tabs - 1)
                Set n = tv.Nodes.Add(pn, tvwChild, , x)
            End If
            
            RaiseEvent DeSerialize(n, tag, index)
            Set parentNodes(tabs) = n
            ClearOldParents parentNodes, tabs
            index = index + 1
        End If
        
     Loop
     
     Close #f
     RestoreTree = tv.Nodes.Count
     
End Function

Sub extractcfg(x)
    Dim a, b
    If Extract(x, "MaxDepth:", " ", a) Then maxDepth = CLng(a)
    If Extract(x, "NodeCount:", " ", b) Then NodeCount = CLng(b)
End Sub

Private Sub ClearOldParents(pn() As Node, curIndex As Long)
    Dim i As Long
    For i = curIndex + 1 To UBound(pn)
        Set pn(i) = Nothing
    Next
End Sub

Private Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Private Function CountTabs(it) As Integer
    Dim tmp() As String
    If InStr(1, it, vbTab, vbTextCompare) < 1 Then Exit Function
    tmp = Split(it, vbTab, , vbTextCompare)
    CountTabs = UBound(tmp)
End Function

Private Function ReadTreeLine(ByRef lineData As String, ByRef TagData As String) As Long
    Dim tmp As String, tabs As Long, v() As String
    tabs = CountTabs(lineData)
    lineData = Replace(lineData, vbTab, Empty)
    If InStr(lineData, ":") > 0 Then
        v = Split(lineData, ":", 2)
        lineData = Replace(v(0), Chr(5), ":")
        TagData = v(1)
        TagData = Replace(TagData, "\t", vbTab)
        TagData = Replace(TagData, "\r", vbCr)
        TagData = Replace(TagData, "\n", vbLf)
    End If
    ReadTreeLine = tabs
End Function

Private Function Extract(value, marker1, marker2, ByRef outVar, _
                Optional includeMarkers As Boolean = False, _
                Optional start As Long = 1, _
                Optional ByRef lastPos As Long, _
                Optional method As VbCompareMethod = vbBinaryCompare _
) As Long

    Dim a As Long, b As Long
    
    lastPos = 0
    outVar = Empty
    
    If Len(marker1) = 0 Then
        a = 1
    Else
        a = InStr(start, value, marker1, method)
        If a < 1 Then Exit Function
    End If
    
    a = a + Len(marker1)
    If Len(marker2) = 0 Then
        outVar = Mid(value, a)
    Else
        b = InStr(a, value, marker2, method)
        If b < 1 Then Exit Function
        
        lastPos = b + Len(marker2)
        outVar = Mid(value, a, b - a)
    End If
    
    If includeMarkers Then outVar = marker1 & outVar & marker2
    Extract = Len(outVar)
    
End Function

Private Function GetParentFolder(path) As String
    Dim tmp() As String, ub
    tmp = Split(path, "\")
    ub = tmp(UBound(tmp))
    GetParentFolder = Replace(Join(tmp, "\"), "\" & ub, "")
End Function

 

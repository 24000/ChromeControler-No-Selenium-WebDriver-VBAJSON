VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a3_BasicInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)


Public PageDefaultTargetId As String

Public TargetInfoJSON As String
Public WebSocketURL As String
Public targetId As String
Public WebSocketUrlPath  As String
Public Context As String

Public DocumentNodeId As Long
Public DocumentObjectId As String

Public IsolatedFrameId As String
Public ExecutionContextId As Long


Private Sub Class_Initialize()
    TargetInfoJSON = GetTragetInfoJSON
    
    WebSocketURL = xJson.GetValue( _
                                TargetInfoJSON, 0, "webSocketDebuggerUrl")
                                
    targetId = xJson.GetValue(TargetInfoJSON, 0, "id")
    
    PageDefaultTargetId = targetId
    
    WebSocketUrlPath = Replace(WebSocketURL, _
                                    "ws://" & xG.LoopBackAddress & ":" & xG.PORT, "")
                                    
    Context = Replace(WebSocketUrlPath, targetId, "")
End Sub


Private Function GetTragetInfoJSON() As String
    Dim client As Object:

On Error GoTo Loop_
Do_:
    Set client = Nothing
    Set client = CreateObject("MSXML2.ServerXMLHTTP")
    client.Open "GET", xG.LoopBackAddress & ":" & xG.PORT & xG.TARGET_LIST_PATH
    client.setRequestHeader "Content-Type", "application/json"
    client.send
    GetTragetInfoJSON = client.responsetext
    Exit Function
Loop_:
DoEvents
Sleep 500
Resume Do_

End Function


Public Function Update(newTargetId As String, _
                                    isPageSwitch As Boolean) As Collection
    If isPageSwitch Then PageDefaultTargetId = newTargetId
    
    targetId = newTargetId
    TargetInfoJSON = GetTragetInfoJSON
    WebSocketURL = GetWebSocketDebuggerUrl()
    WebSocketUrlPath = Replace(WebSocketURL, "ws://" & xG.LoopBackAddress & ":" & xG.PORT, "")
    Context = Replace(WebSocketUrlPath, targetId, "")
End Function

Private Function GetWebSocketDebuggerUrl() As String
    Dim collect As Collection: Set collect = New Collection
    Dim targetInfo As String
    Dim Count As Long
    Do
        targetInfo = xJson.GetObject(TargetInfoJSON, Count)
        If targetInfo = "" Then Exit Do
        collect.add targetInfo
        Count = Count + 1
    Loop
    
    Dim i As Long
    Dim targetType As String, id As String
    For i = 1 To collect.Count
            targetType = xJson.GetValue(collect(i), "type")
            If targetType <> "page" Then GoTo next_
            id = xJson.GetValue(collect(i), "id")
            If id <> targetId Then GoTo next_
            GetWebSocketDebuggerUrl = xJson.GetValue(collect(i), "webSocketDebuggerUrl")
            Exit Function
next_:
    Next
End Function


Public Function IsTargetPageDefault() As Boolean
    If targetId = PageDefaultTargetId Then
        IsTargetPageDefault = True
    Else
        IsTargetPageDefault = False
    End If
End Function


'çƒìxCreateIsolatedWorldÇ∑ÇÈèàóùÅB
Public Sub UpdateExecutionContextId(CDP As a4_ExecuteCDP)
    If IsolatedFrameId <> "" Then
        ExecutionContextId = CDP.CreateIsolatedFrameWorld(IsolatedFrameId)
        DoEvents
    End If
End Sub



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a_0_x3_BasicInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const HOST_ As String = "127.0.0.1"
Private Const DEBUGGING_PORT As Long = 9222
Private Const RESOURCE_PATH As String = "/json" 'タブのリスト
Private Const HTTP_SCHEMA As String = "http://"
Private Const WEBSOCKET_SCHEMA As String = "ws://"

Private Const KEY_WEBSOCKETURL As String = "webSocketDebuggerUrl"

Public PageDefaultTargetId As String

Public TargetInfoJSON As String
Public WebSocketURL As String
Public TargetId As String
Public WebSocketUrlPath  As String
Public Context As String

Public Property Get Host() As String
    Host = HOST_
End Property
Public Property Get DebuggingPort() As String
    DebuggingPort = DEBUGGING_PORT
End Property
Public Property Get ResourcePath() As String
    ResourcePath = RESOURCE_PATH
End Property

Public Sub Init(json As a_0_x4_JSONHandler)
    TargetInfoJSON = GetTragetInfoJSON
    
    WebSocketURL = json.GetValueFrom( _
                                TargetInfoJSON, Array(0, KEY_WEBSOCKETURL))
                                
    TargetId = json.GetValueFrom(TargetInfoJSON, Array(0, "id"))
    
    PageDefaultTargetId = TargetId
    
    WebSocketUrlPath = Replace(WebSocketURL, _
                                    WEBSOCKET_SCHEMA & HOST_ & ":" & DEBUGGING_PORT, "")
                                    
    Context = Replace(WebSocketUrlPath, TargetId, "")
End Sub

Private Function GetTragetInfoJSON() As String
    Dim client As Object
    Set client = CreateObject("MSXML2.ServerXMLHTTP")

On Error GoTo Loop_
do_:
    client.Open "GET", HTTP_SCHEMA & HOST_ & ":" & DEBUGGING_PORT & RESOURCE_PATH
    client.setRequestHeader "Content-Type", "application/json"
    client.send
    GetTragetInfoJSON = client.responsetext
    Exit Function
Loop_:
DoEvents
Resume do_

End Function


Public Function Update(newTargetId As String, json As a_0_x4_JSONHandler, _
                                    isPageSwitch As Boolean) As Collection
    If isPageSwitch Then PageDefaultTargetId = newTargetId
    
    TargetId = newTargetId
    TargetInfoJSON = GetTragetInfoJSON
    WebSocketURL = GetWebSocketDebuggerUrl(json)
    WebSocketUrlPath = Replace(WebSocketURL, WEBSOCKET_SCHEMA & HOST_ & ":" & DEBUGGING_PORT, "")
    Context = Replace(WebSocketUrlPath, TargetId, "")
End Function

Private Function GetWebSocketDebuggerUrl(json As a_0_x4_JSONHandler) As String
    Dim collect As Collection: Set collect = New Collection
    Dim targetInfo As String
    Dim Count As Long
    Do
        targetInfo = json.GetObjectFrom(TargetInfoJSON, Array(Count))
        If targetInfo = "" Then Exit Do
        collect.add targetInfo
        Count = Count + 1
    Loop
    
    Dim i As Long
    Dim targetType As String, Id As String
    For i = 1 To collect.Count
            targetType = json.GetValueFrom(collect(i), Array("type"))
            If targetType <> "page" Then GoTo next_
            Id = json.GetValueFrom(collect(i), Array("id"))
            If Id <> TargetId Then GoTo next_
            GetWebSocketDebuggerUrl = json.GetValueFrom(collect(i), Array(KEY_WEBSOCKETURL))
            Exit Function
next_:
    Next
End Function


Public Function IsTargetPageDefault() As Boolean
    If TargetId = PageDefaultTargetId Then
        IsTargetPageDefault = True
    Else
        IsTargetPageDefault = False
    End If
End Function

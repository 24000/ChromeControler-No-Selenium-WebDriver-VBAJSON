VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a_0_x3_ReturnInfos"
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
Private Const KEY_ID As String = "id"

Public TargetInfoJSON As String
Public WebSocketURL As String
Public targetId As String
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

Public Sub Init(jsonhandler As a_0_x4_JSONHandler)
    TargetInfoJSON = GetTragetInfoJSON
    WebSocketURL = jsonhandler.GetValueFromJson(TargetInfoJSON, _
                                                    Array(0, KEY_WEBSOCKETURL))
    targetId = jsonhandler.GetValueFromJson(TargetInfoJSON, _
                                                    Array(0, KEY_ID))
    WebSocketUrlPath = Replace(WebSocketURL, _
                                    WEBSOCKET_SCHEMA & HOST_ & ":" & DEBUGGING_PORT, "")
    Context = Replace(WebSocketUrlPath, targetId, "")
End Sub

Private Function GetTragetInfoJSON() As String
    Dim client As Object
    Set client = CreateObject("MSXML2.ServerXMLHTTP")
    client.Open "GET", HTTP_SCHEMA & Host & ":" & DEBUGGING_PORT & RESOURCE_PATH
    client.setRequestHeader "Content-Type", "application/json"
    client.send
    GetTragetInfoJSON = client.responsetext
End Function



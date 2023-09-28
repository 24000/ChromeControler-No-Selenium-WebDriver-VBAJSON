VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a3_BrowserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents event_ As a5_CDPEventNotifier    'event受信のみ
Attribute event_.VB_VarHelpID = -1

Public Enum Target_Event_Name
    IsScheduledNavigation
    IsStartedLoading
    IsClearedScheduledNavigation
    IsTargetNavigated
End Enum


Private page_ As a3_x1_Page
Private iframe_ As a3_x2_Iframe
Private Dialog_ As a3_x3_Dialog

'contextID(flameをisolatedCreate)関連
Private IsCurrentContextIdDestroyed As Boolean
Public ExecutionContextId As Long

Public Property Get WebSocketPath(pageNum As Long) As String
    Dim URL As String: URL = page_.GetWebSocketDebuggerUrl(pageNum)
    WebSocketPath = Mid(URL, (InStr(URL, "devtools") - 1))
End Property
Public Property Get targetId(pageNum As Long) As String
    targetId = page_.GetTargetId(pageNum)
End Property
Public Property Get PageInfosCount() As Long
    PageInfosCount = page_.List.Count
End Property
Public Property Get URL() As String
    URL = page_.List(1)("url")
End Property


Public Sub Init(json As String, notifyer As a5_CDPEventNotifier)
    Set event_ = notifyer
    Set page_ = New a3_x1_Page
    Set iframe_ = New a3_x2_Iframe
    Set Dialog_ = New a3_x3_Dialog
    
    Dim target As Object
    Dim i As Long, jsObj As String
    For i = 0 To (xJson.GetJsonLength(json) - 1)
        jsObj = xJson.GetObject(json, i)
        Set target = GetTarget(jsObj)
        If target("type") = "page" Then
            page_.List.add target("id"), target
        ElseIf target("type") = "iframe" Then
            iframe_.List.add target("id"), target
        End If
    Next
End Sub


'JSオブジェクト文字列からターゲットの情報を辞書として返す
Private Function GetTarget(jsObj As String) As Object
    Dim target As Object: Set target = CreateObject("Scripting.Dictionary")
    target.add "id", xJson.GetValue(jsObj, "id")
    target.add "title", xJson.GetValue(jsObj, "title")
    target.add "type", xJson.GetValue(jsObj, "type")
    target.add "url", xJson.GetValue(jsObj, "url")
    target.add "webSocketDebuggerUrl", xJson.GetValue(jsObj, "webSocketDebuggerUrl")
    target.add "browserContextId", xJson.GetValue(jsObj, "browserContextId")
    If target("type") = "iframe" Then target.add "parentId", xJson.GetValue(jsObj, "parentId")
    
    Set GetTarget = target
End Function




'Public Function Update(newTargetId As String, _
'                                    isPageSwitch As Boolean) As Collection
'    If isPageSwitch Then PageDefaultTargetId = newTargetId
'
'    TargetId = newTargetId
'    TargetInfoJSON = GetTragetInfoJSON
'    WebSocketURL = GetWebSocketDebuggerUrl()
'    WebSocketUrlPath = Replace(WebSocketURL, "ws://" & xG.LoopBackAddress & ":" & xG.PORT, "")
'    Context = Replace(WebSocketUrlPath, TargetId, "")
'End Function

'Private Function GetWebSocketDebuggerUrl() As String
'    Dim collect As Collection: Set collect = New Collection
'    Dim targetInfo As String
'    Dim Count As Long
'    Do
'        targetInfo = xJson.GetObject(TargetInfoJSON, Count)
'        If targetInfo = "" Then Exit Do
'        collect.add targetInfo
'        Count = Count + 1
'    Loop
'
'    Dim i As Long
'    Dim targetType As String, Id As String
'    For i = 1 To collect.Count
'            targetType = xJson.GetValue(collect(i), "type")
'            If targetType <> "page" Then GoTo next_
'            Id = xJson.GetValue(collect(i), "id")
'            If Id <> TargetId Then GoTo next_
'            GetWebSocketDebuggerUrl = xJson.GetValue(collect(i), "webSocketDebuggerUrl")
'            Exit Function
'next_:
'    Next
'End Function


'Public Function IsTargetPageDefault() As Boolean
'    If TargetId = PageDefaultTargetId Then
'        IsTargetPageDefault = True
'    Else
'        IsTargetPageDefault = False
'    End If
'End Function


'再度CreateIsolatedWorldする処理。
Public Sub UpdateExecutionContextId(CDP As a4_ExecuteCDP)
    If iframe_.IsolatedId <> "" Then
        ExecutionContextId = CDP.CreateIsolatedFrameWorld(iframe_.IsolatedId)
        DoEvents
    End If
End Sub



Public Function SetPageStatus(eName As Target_Event_Name, bool As Boolean)
    Select Case eName
        Case Target_Event_Name.IsScheduledNavigation
            page_.IsScheduledNavigation = bool
        Case Target_Event_Name.IsClearedScheduledNavigation
            page_.IsClearedScheduledNavigation = bool
        Case Target_Event_Name.IsStartedLoading
            page_.IsStartedLoading = bool
        Case Target_Event_Name.IsTargetNavigated
            page_.IsTargetNavigated = bool
    End Select
End Function
Public Function GetPageStatus(eName As Target_Event_Name) As Boolean
    Select Case eName
        Case Target_Event_Name.IsScheduledNavigation
            GetPageStatus = page_.IsScheduledNavigation
        Case Target_Event_Name.IsClearedScheduledNavigation
            GetPageStatus = page_.IsClearedScheduledNavigation
        Case Target_Event_Name.IsStartedLoading
            GetPageStatus = page_.IsStartedLoading
        Case Target_Event_Name.IsTargetNavigated
            GetPageStatus = page_.IsTargetNavigated
    End Select
End Function
Public Function SetIframeStatus(eName As Target_Event_Name, bool As Boolean)
    Select Case eName
        Case Target_Event_Name.IsScheduledNavigation
            iframe_.IsScheduledNavigation = bool
        Case Target_Event_Name.IsClearedScheduledNavigation
            iframe_.IsClearedScheduledNavigation = bool
        Case Target_Event_Name.IsStartedLoading
            iframe_.IsStartedLoading = bool
        Case Target_Event_Name.IsTargetNavigated
            iframe_.IsTargetNavigated = bool
    End Select
End Function
Public Function GetIframeStatus(eName As Target_Event_Name) As Boolean
    Select Case eName
        Case Target_Event_Name.IsScheduledNavigation
            GetIframeStatus = page_.IsScheduledNavigation
        Case Target_Event_Name.IsClearedScheduledNavigation
            GetIframeStatus = page_.IsClearedScheduledNavigation
        Case Target_Event_Name.IsStartedLoading
            GetIframeStatus = page_.IsStartedLoading
        Case Target_Event_Name.IsTargetNavigated
            GetIframeStatus = page_.IsTargetNavigated
    End Select
End Function



























'イベント処理＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝

Private Sub event__TargetCreated(json As String)
    'タイプとIDを取得
    Dim t As String: t = xJson.GetValue(json, "type")
    Dim id As String: id = xJson.GetValue(json, "targetId")
    
    'タイプによりそれぞれのリストをチェックし、既に含まれている場合は捨てる、含まれていない場合はリストへ追加
    If t = "page" Then
        If page_.Exists(id) Then
            Exit Sub
        Else
            page_.AddTarget id, t, json
        End If
    ElseIf t = "iframe" Then
        If iframe_.Exists(id) Then
            Exit Sub
        Else
            iframe_.AddTarget id, t, json
        End If
    Else
        Stop
    End If
    
End Sub
Private Sub event__TargetInfoChange(json As String)
    'タイプとIDを取得
    Dim t As String: t = xJson.GetValue(json, "type")
    Dim id As String: id = xJson.GetValue(json, "targetId")
    
    Dim target As Object
    If t = "page" Then
        page_.ChangeTarget id, json
    ElseIf t = "iframe" Then
        iframe_.ChangeTarget id, json
    End If
End Sub
Private Sub event__TargetDestroyed(json As String)
    Debug.Print "TargetDestroyed"
    'タイプが情報に含まれないため、IDを取得
    Dim id As String: id = xJson.GetValue(json, "targetId")
    
    Dim target As Object
    If page_.Exists(id) Then
        page_.RemoveTarget id
    ElseIf iframe_.Exists(id) Then
        iframe_.RemoveTarget id
    End If
End Sub


Private Sub event__FrameScheduledNavigation(json As String)
'            If InStr(json, "anchorClick") > 0 Then
'                If isCurrentTarget Then IsCurrentTargetScheduledNavigation = True
'                If isIsolatedFrameId Then
'                    IsIsolatedFrameScheduledNavigation = True
'                    Debug.Print "★★IsolatedFrameScheduledNavigation"
'                End If
'            End If
End Sub

Private Sub event__frameClearedScheduledNavigation(json As String)
'            If isCurrentTarget Then IsCurrentTargetClearedScheduledNavigation = True
'            If isIsolatedFrameId Then IsIsolatedFrameClearedScheduledNavigation = True
End Sub
Private Sub event__FrameStartedLoading(json As String)
'            If isCurrentTarget Then IsCurrentTargetStartedLoading = True
'            If isIsolatedFrameId Then
'                IsIsolatedFrameStartedLoading = True
'                Sleep 1000
'            End If
End Sub
Private Sub event__FrameNavigated(json As String)
'            If isCurrentTarget Then IsCurrentTargetNavigated = True
'            If isIsolatedFrameId Then
'                IsIsolatedFrameNavigated = True
'                Debug.Print "★★IsolatedFrameNavigated"
'            End If
End Sub

Private Sub event__ExecutionContextCreated(json As String)
    Debug.Print "ExecutionContextCreated"
End Sub

Private Sub event__ExecutionContextDestroyed(json As String)
'            Dim contextId As Long
'            contextId = xJson.GetValue(json, "params", "executionContextId")
'            If contextId = basicInfo_.ExecutionContextId Then
'                IsCurrentContextIdDestroyed = True
'                Debug.Print "★★IsCurrentContextIdDestroyed"
'            End If
End Sub



Private Sub event__JavaScriptDialogOpening(json As String)
    Debug.Print "JavaScriptDialogOpening"
End Sub
Private Sub event__JavaScriptDialogClosed(json As String)
    Debug.Print "JavaScriptDialogClosed"
End Sub



'========================================================



Public Sub ForDebug_ShowtargetInfo()
    Dim i  As Long
    Debug.Print "PageInfos↓==============="
    For i = 1 To page_.List.Count
        Debug.Print i & "：" & page_.List(i)("targetId")
    Next
    Debug.Print "IframeInfos↓==============="
    For i = 1 To iframe_.List.Count
        Debug.Print i & "：" & iframe_.List(i)("targetId")
    Next
    Debug.Print "========================"
End Sub










VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a_0_x2_CDPMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private id_ As Long

Public Property Get Id() As Long
    Id = id_
End Property

'TargetDomain*******************************************************************
'pageDomainの通知を有効にする。
Public Function PageEnable() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.enable""}"
    PageEnable = message
End Function

Public Function SetDiscoverTargets(TargetId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.setDiscoverTargets""," & _
                      """params"": {""discover"": true}}"
    SetDiscoverTargets = message
End Function

Public Function AttachToTarget(TargetId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.attachToTarget""," & _
                      """params"": {""targetId"": """ & TargetId & """}}"
    AttachToTarget = message
End Function

Public Function ActivateTarget(TargetId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.activateTarget""," & _
                      """params"": {""targetId"": """ & TargetId & """}}"
    ActivateTarget = message
End Function

Public Function DetachFromTarget(TargetId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.detachFromTarget""," & _
                      """params"": {""targetId"": """ & TargetId & """}}"
    DetachFromTarget = message
End Function

Public Function TargetClose(TargetId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.closeTarget""," & _
                      """params"": {""targetId"": """ & TargetId & """}}"
    TargetClose = message
End Function
'利用可能なターゲットのリストを取得する。
Public Function GetTargets() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Target.getTargets""}"
    GetTargets = message
End Function
'*******************************************************************


'BrowserDomain*******************************************************************
Public Function SessionClose() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Browser.close""}"
    SessionClose = message
End Function
'*******************************************************************

'PageDomain*******************************************************************
Public Function URLOpen(URL As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.navigate""," & _
                      """params"": {""url"": """ & URL & """}}"
    URLOpen = message
End Function

Public Function HandleJavaScriptDialog() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.handleJavaScriptDialog""," & _
                     """params"": {""accept"": true}}"
    HandleJavaScriptDialog = message
End Function

Public Function GetFrameTree() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.getFrameTree""}"
    GetFrameTree = message
End Function

Public Function CreateIsolatedFrameWorld(FrameId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Page.createIsolatedWorld""," & _
                     """params"": {""frameId"": """ & FrameId & """,""grantUniveralAccess"": true}}"
    CreateIsolatedFrameWorld = message
End Function

'*******************************************************************
'Input Domain*******************************************************************
'キーイベントをページにディスパッチする。keyDown , keyUp, rawKeyDown, charのいずれか。
Public Function DispatchKeyEvent_DOMKeyCode(eventType As String, code As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Input.dispatchKeyEvent""," & _
                      """params"": {""type"":  """ & eventType & """,""code"": """ & code & """}}"
    DispatchKeyEvent_DOMKeyCode = message
    
'    windowsVirtualKeyCode
End Function

'マウスイベントをxy座標にディスパッチする。mousePressed, mouseReleased, mouseMoved, mouseWheelのいずれか。
Public Function DispatchMouseEvent(eventType As String, x As Long, y As Long, _
                                                        button As String, clickCount As Long, _
                                                        Optional modifiers As Long = 0, _
                                                        Optional deltaX As Long = 0, Optional deltaY As Long = 0) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Input.dispatchMouseEvent""," & _
                      """params"": {""type"":  """ & eventType & """,""x"": " & x & ",""y"": " & y & "," & _
                                            """button"": """ & button & """,""clickCount"": " & clickCount
                                            'none, left, middle, right, back, forwardのいずれか
                                            
    'Alt=1, Ctrl=2, Meta/Command=4, Shift=8
    If modifiers <> 0 Then
        message = message & ",""modifiers"": " & modifiers
    End If
    
    If deltaX <> 0 Or deltaY <> 0 Then
        message = message & ",""deltaX"": " & deltaX & ",""deltaY"": " & deltaY
    End If
    
    message = message & "}}"
    DispatchMouseEvent = message
End Function
'*******************************************************************
'DOM Domain*******************************************************************
'Domの有効化
Public Function DOM_Enable()
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.enable""}"
    DOM_Enable = message
End Function

'指定された要素にフォーカスを当てる。
Public Function Focus(objectId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.focus""," & _
                      """params"": {""objectId"":  """ & objectId & """}}"
    Focus = message
End Function

'DocumentObjectのNodeIdを取得
Public Function GetDocumentID() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.getDocument""," & _
                     """params"": {""depth"": " & 1 & "}}"
    GetDocumentID = message
End Function

'CSSセレクタによりNode取得
Public Function querySelector(Selector As String, Optional nodeId As Long = 1) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.querySelector""," & _
                      """params"": {""nodeId"":  " & nodeId & ",""selector"": """ & Selector & """}}"
    querySelector = message
End Function

'objectIdからnodeId取得
Public Function RequestNode(objectId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.requestNode""," & _
                      """params"": {""objectId"":  """ & objectId & """}}"
    RequestNode = message
End Function

'nodeIdからobjectId取得
Public Function ResolveNode(nodeId As Long)
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.resolveNode""," & _
                      """params"": {""nodeId"":  " & nodeId & "}}"
    ResolveNode = message
End Function
'idで指定されたノードの属性情報を取得
Public Function GetAttributes(nodeId As Long)
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.getAttributes""," & _
                      """params"":{""nodeId"":  " & nodeId & "}}"
    GetAttributes = message
End Function

'id で指定されたノード情報を叙述
Public Function DescribeNode(Id As Variant) As Variant
    id_ = id_ + 1
    Dim message As String
    If IsNumeric(Id) Then
        message = "{""id"": " & id_ & ",""method"": ""DOM.describeNode""," & _
                          """params"": {""nodeId"":  " & Id & ",""depth"": 1,""pierce"": false}}"
    Else
        message = "{""id"": " & id_ & ",""method"": ""DOM.describeNode""," & _
                         """params"": {""objectId"":  """ & Id & """,""depth"": 1,""pierce"": false}}"
    End If
    DescribeNode = message
End Function

Public Function RequestChildNodes() As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.RequestChildNodes""," & _
                      """params"": {""nodeId"":  " & id_ & ",""depth"":  -1}}"
    RequestChildNodes = message
End Function

'指定された要素にフォーカスを当てる。
Public Function GetBoxModel(objectId As String) As String
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.getBoxModel""," & _
                      """params"": {""objectId"":  """ & objectId & """}}"
    GetBoxModel = message
End Function

'ファイル登録インプットボックスへファイルPath登録
Public Function SetFileInputFiles(objectId As Variant, files As String) As String
    id_ = id_ + 1
    files = "[""" & files & """]"
    
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""DOM.setFileInputFiles""," & _
                      """params"": " & _
                     "{""files"": " & files & ",""objectId"": """ & objectId & """}}"
    
    SetFileInputFiles = message
End Function

'RunTime Domain*******************************************************************
Public Function RunTime_Enable()
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Runtime.enable""}"
    RunTime_Enable = message
End Function
Public Function JavaScritpEvaluate(script As String)
    id_ = id_ + 1
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Runtime.evaluate""," & _
                      """params"": {""expression"": """ & script & """,""returnByValue"": false}}"
    JavaScritpEvaluate = message
End Function

Public Function CallJavaScriptFunction(funcDeclear As String, objectId As Variant, argments As Variant)
    id_ = id_ + 1
    Dim JavaScriptObjectArgs As String
    
    If argments = "true" Or argments = "false" Then
        JavaScriptObjectArgs = "[{""value"": " & argments & "}]"
    Else
        JavaScriptObjectArgs = "[{""value"": """ & argments & """}]"
    End If
    
    Dim message As String
    message = "{""id"": " & id_ & ",""method"": ""Runtime.callFunctionOn""," & _
                      """params"": " & _
                     "{""functionDeclaration"": """ & funcDeclear & """,""arguments"": " & JavaScriptObjectArgs
    If IsNumeric(objectId) Then
        message = message & ",""executionContextId"": " & objectId & "}}"
    Else
        message = message & ",""objectId"": """ & objectId & """}}"
    End If
    CallJavaScriptFunction = message
End Function
'*******************************************************************

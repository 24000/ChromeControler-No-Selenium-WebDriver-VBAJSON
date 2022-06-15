VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromeControler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IWebDriver

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private Const CHROME_EXE As String = """chrome.exe"""
Private Const DEBUG_MODE_OPTION As String = " --remote-debugging-Port=9222"
Private Const USER_DATA_DIR_OPTION As String = " --user-data-dir="
Private Const ENABLE_AUTOMATION_OPTION As String = " --enable-automation"

Private userDataDir As String

Private webSocket_ As a_0_x1_WebSocketCommunicator
Private CDPMethod_ As a_0_x2_CDPMethods
Private jsonHandler_ As a_0_x4_JSONHandler
Private return_ As a_0_x3_ReturnInfos
Private events_ As a_0_x1_x2_CDPEventHandler

Public documentNodeId_ As Long
Public documentObjectId_ As String

Public isolatedFrameId_ As String
Public executionContextId_ As Long

Public Property Get WebSocketURL() As String
    WebSocketURL = return_.WebSocketURL
End Property
Public Property Get UrlPath() As String
    UrlPath = return_.WebSocketUrlPath
End Property

Public Property Get WebSocket() As a_0_x1_WebSocketCommunicator
    Set WebSocket = webSocket_
End Property
Public Property Get CDPMethod() As a_0_x2_CDPMethods
    Set CDPMethod = CDPMethod_
End Property
Public Property Get jsonhandler() As a_0_x4_JSONHandler
    Set jsonhandler = jsonHandler_
End Property
Public Property Get CDPEvents() As a_0_x1_x2_CDPEventHandler
    Set CDPEvents = events_
End Property


Private Sub Class_Initialize()
    
    userDataDir = ThisWorkbook.Path & "\マクロ実行で作成されるフォルダ_削除OK\"
    If Dir(userDataDir, vbDirectory) = "" Then
        MkDir userDataDir
    End If
    
    CreateObject("WScript.Shell").Run _
        CHROME_EXE & DEBUG_MODE_OPTION & ENABLE_AUTOMATION_OPTION _
                              & USER_DATA_DIR_OPTION & "" & userDataDir & ""
    Sleep 2000
    
    Set jsonHandler_ = New a_0_x4_JSONHandler
    Set return_ = New a_0_x3_ReturnInfos: return_.Init jsonhandler
    Set webSocket_ = New a_0_x1_WebSocketCommunicator
    If webSocket_.Init(return_.Host, return_.WebSocketUrlPath, return_.DebuggingPort) = False Then End
    
    Set CDPMethod_ = New a_0_x2_CDPMethods
    Set events_ = New a_0_x1_x2_CDPEventHandler
End Sub

'message処理送信＆受信
'受信時にはクロームからイベント発生が通知される場合があるので、idで送信に対する受信を判断
'※イベント通知にはidが含まれない
Public Function ProcessMessage( _
    sendId As Long, message As String, jsonhandler As a_0_x4_JSONHandler) As String

    If webSocket_.SendMessage(message) = False Then GoTo quit

    Dim responseJson As String, responseId As Variant
    Do
        responseJson = webSocket_.GetMessage
        If responseJson = "受信エラー発生" Then GoTo quit
            
        responseId = jsonhandler.GetValueFromJson(responseJson, Array("id"))
        If responseId = "" Then
            If InStr(responseJson, "method") > 0 Then
                events_.GetInfo responseJson, Me
            End If
        End If
    Loop While (sendId <> responseId)
    
    ProcessMessage = responseJson
Exit Function

quit:
Dim errMsg As String
If responseJson = "" Then
    errMsg = "送信エラー発生：イミディエイト参照"
Else
    errMsg = "受信エラー発生:イミディエイト参照"
End If
MsgBox errMsg, vbCritical, "エラー"
webSocket_.Http.CloseHInternetHandles
End
End Function

'ElementId系************************************************
'CDPMethodにないのでJavascript実行
Private Function GetDocumentObjectID() As String
    Dim responseJson  As String
    responseJson = ExecuteJavaScript_Inner("document;")
    documentObjectId_ = jsonHandler_.GetValueFromJson(responseJson, _
                                                        Array("result", "result", "objectId"))
    
    GetDocumentObjectID = documentObjectId_
End Function

Private Function GetDocumentNodeID() As Long
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.GetDocumentID
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
    GetDocumentNodeID = jsonHandler_.GetValueFromJson(responseJson, Array("result", "root", "nodeId"))
End Function

Private Function ChangeObjectIdToNodeId(objectId As String) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.RequestNode(objectId)
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
    
    Dim nodeId As String
    nodeId = jsonHandler_.GetValueFromJson(responseJson, Array("result", "nodeId"))
    ChangeObjectIdToNodeId = nodeId
End Function

Private Function ChangeNodeIdToObjectId(nodeId As Long) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.ResolveNode(nodeId)
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
    
    Dim objectId As String
    objectId = jsonHandler_.GetValueFromJson(responseJson, Array("result", "objectId"))
    ChangeNodeIdToObjectId = objectId
End Function
'************************************************

'Element情報系************************************************
Private Function GetAttributes(nodeId As Long) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.GetAttributes(nodeId)
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
    GetAttributes = responseJson
End Function

Private Function DescribeNode(IdName As String, ID As Variant) As Variant
    If IdName <> "nodeId" And IdName <> "objectId" Then
        Err.Raise 10000, , "第一引数には""nodeId""か""objectId""のどちらか指定してください。"
    End If
    
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.DescribeNodeBy_nodeId_or_objectId(IdName, ID)
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
    Dim attributes As String
    attributes = jsonHandler_.GetValueFromJson(responseJson, Array("result", "node", "attributes"))
    DescribeNode = attributes
End Function
'************************************************

'Runtime************************************************
'JavaScriptコード実行
Private Function ExecuteJavaScript_Inner(Script As String) As String
    Dim jsonForSending  As String
    jsonForSending = CDPMethod_.JavaScritpEvaluate(Script)
    ExecuteJavaScript_Inner = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
End Function

'JavaScript_Function実行
Public Function JavaScritpCallFunction(funcName As String, objectId As Variant, args As String) As String
    Dim jsonForSending  As String
    jsonForSending = CDPMethod_.JavaScritpCallFunction(funcName, objectId, args)
    JavaScritpCallFunction = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
End Function
'************************************************


'DOM操作系************************************************
Public Sub DOMEnable()
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.DOM_Enable()
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
End Sub

Public Sub FindElementByQuerySelector(Selector As String)
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.querySelector(Selector)
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
    
    Dim nodeId As Long
    nodeId = jsonHandler_.GetValueFromJson(responseJson, Array("result", "nodeId"))
    Dim nodeInfoJson As String
    nodeInfoJson = DescribeNode("nodeId", nodeId)
End Sub

Public Function FindElementByID(ID As String) As String
    
    
End Function
'************************************************

'Input系************************************************
Public Function Click(objId As String) As String
    JavaScritpCallFunction "function(value){this.click();}", objId, ""
End Function
'************************************************

'Target:Page系************************************************
Public Function PageEnable() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.PageEnable()
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
End Function

Public Function GetFrameTree() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.GetFrameTree()
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
    GetFrameTree = responseJson
End Function

Public Function CreateIsolatedFrameWorld(frameId As String) As Long
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.CreateIsolatedFrameWorld(frameId)
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
    CreateIsolatedFrameWorld = _
        jsonHandler_.GetValueFromJson(responseJson, Array("result", "executionContextId"))
End Function

Public Function SetDiscoverTargets() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.SetDiscoverTargets(return_.targetId)
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
End Function

Public Function AttachBrowser() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.AttachBrowser(return_.targetId)
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
End Function
Public Function TargetClose() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.TargetClose(return_.targetId)
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
End Function
Public Function GetTargetInfo() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.GetTargets()
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
End Function













'以下､インターフェイス公開メソッドおよびその関連メソッド
'************************************************

'Public
'第1引数で指定のURLを開く
Private Sub IWebDriver_URLOpen(url As String, Optional maxWaitTime_ms As Long = 30000)
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.URLOpen(url)
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
    If InStr(responseJson, "invalid URL") > 0 Then GoTo quit
    
    WaitingForWebPageDisplay (maxWaitTime_ms)
    
    SetDiscoverTargets
    PageEnable
    DOMEnable
    documentObjectId_ = GetDocumentObjectID
    documentNodeId_ = GetDocumentNodeID
Exit Sub

quit:
DisplayErrorAndEnd "指定された以下URLは無効なURLです。" & vbCrLf & url
End Sub

'ページ表示待ちループ
Private Function WaitingForWebPageDisplay(maxWaitTime_ms As Long)
    Dim javaScriptCode As String, responseJson As String, state As String
    Do
        javaScriptCode = "document.readyState;"
        responseJson = ExecuteJavaScript_Inner(javaScriptCode)
        state = jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "value"))
        DoEvents
        Sleep 800
    Loop Until (state = "complete")
End Function

'Public
'第1引数のJavaScriptコードの実行
Private Function IWebDriver_ExecuteJavaScript(Script As String) As String
    Dim responseJson As String
    responseJson = ExecuteJavaScript_Inner(Script)
    IWebDriver_ExecuteJavaScript = jsonHandler_.GetValueFromJson(responseJson, Array("result"))
End Function

'Public
'第1引数で指定されるXpathの要素を取得し、エレメントObjectとして返却
Private Function IWebDriver_GetElementByXpath(xpath As String) As IWebElement
    ReCreateIsolatedFrameWorld_IfFlagOn
    
    Dim objectId As String
    objectId = GetObjectIdOfElementByXpath(xpath)
    If objectId = "" Then GoTo quit
    Set IWebDriver_GetElementByXpath = MakeElementObject(objectId)
Exit Function

quit:
DisplayErrorAndEnd "指定された以下Xpathの要素は存在しませんでした。" & vbCrLf & xpath
End Function
'Xpathで要素を探し、オブジェクトIDを返す
Private Function GetObjectIdOfElementByXpath(xpath As String) As String
    Dim javaScriptCode As String, responseJson As String
    If executionContextId_ = 0 Then
        javaScriptCode = _
            "document.evaluate('" & xpath & "', document, null, " & _
                "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;"
        responseJson = ExecuteJavaScript_Inner(javaScriptCode)
    Else
        javaScriptCode = _
            "function(xpath){" & _
                "return  document.evaluate(xpath, document, null, " & _
                    "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;" & _
            "}"
        responseJson = JavaScritpCallFunction(javaScriptCode, executionContextId_, xpath)
    End If
    
    GetObjectIdOfElementByXpath = _
        jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "objectId"))
End Function
'エレメントを作成し、戻り値として返却
Private Function MakeElementObject(objectId As String) As IWebElement
    DOMEnable
    Dim nodeId As Long
    nodeId = ChangeObjectIdToNodeId(objectId)
    Dim attributes As String
    attributes = GetAttributes(nodeId)
    
    Dim element As WebElement: Set element = New WebElement
    element.Init Me, objectId, nodeId
    Set MakeElementObject = element
End Function

'Public
'第1引数で指定される値のフレームにフレームを切替
Private Sub IWebDriver_SwitchFrameByNameOrUrlOrIndex(value As Variant)
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim frameId As String
    If IsNumeric(value) Then
        If frameInfos.count < value Then GoTo quit
        frameId = jsonHandler_.GetValueFromJson(frameInfos(value), Array("id"))
    Else
        frameId = GetTargetFrameId(frameInfos, value)
    End If
    If frameId = "" Then GoTo quit
    executionContextId_ = CreateIsolatedFrameWorld(frameId)
    isolatedFrameId_ = frameId
    
Exit Sub

quit:
DisplayErrorAndEnd "指定されたframeは存在しません。" & vbCrLf & value & vbCrLf & vbCrLf & _
                              "※もしUrl指定の場合、Chromeが情報を返さないケースがあります。" & vbCrLf & _
                              "   とれないようであればIndex、またはあればName指定を推奨します。"
End Sub
'frameツリーを取得し、各フレームの情報（JSONオブジェクト文字列）をコレクションとして集約
Private Function GetFrameInfos() As Collection
    Dim collect As Collection: Set collect = New Collection
    Dim json As String: json = GetFrameTree
    collect.Add jsonHandler_.GetObjectFromJson(json, Array("result", "frameTree", "frame"))
    
    Dim frmInfo As String, frminfo2 As String
    Dim count As Long
    Do
        frmInfo = jsonHandler_.GetObjectFromJson(json, _
                        Array("result", "frameTree", "childFrames", count))
        If frmInfo = "" Then Exit Do
        frminfo2 = jsonHandler_.GetObjectFromJson(frmInfo, Array("frame"))
        collect.Add frminfo2
        If InStr(frmInfo, "childFrames") > 0 Then
            GetChildFrameInfo collect, frmInfo
        End If
        count = count + 1
    Loop
    Set GetFrameInfos = collect
End Function
'childFrameがある場合、再帰処理
Public Sub GetChildFrameInfo(collect As Collection, json As String)
    Dim frmInfo As String, frminfo2 As String
    Dim count As Long
    Do
        frmInfo = jsonHandler_.GetObjectFromJson(json, _
                        Array("childFrames", count))
        If frmInfo = "" Then Exit Do
        frminfo2 = jsonHandler_.GetObjectFromJson(frmInfo, Array("frame"))
        collect.Add frminfo2
        If InStr(frmInfo, "childFrames") > 0 Then
            GetChildFrameInfo collect, frmInfo
        End If
        count = count + 1
    Loop
End Sub
'各フレームのname,urlと指定された値を比較。一致するフレームのIdを返す。
Private Function GetTargetFrameId(frameInfos As Collection, value As Variant) As String
    Dim frameInfo As Variant
    Dim name As String, url As String
    For Each frameInfo In frameInfos
        name = jsonHandler_.GetValueFromJson(frameInfo, Array("name"))
        url = jsonHandler_.GetValueFromJson(frameInfo, Array("url"))
        If name = value Or url = value Then
            GetTargetFrameId = jsonHandler_.GetValueFromJson(frameInfo, Array("id"))
            Exit Function
        End If
    Next
End Function

'Public
'デフォルトのフレームに戻す
Private Sub IWebDriver_SwitchFrameToDefault()
    isolatedFrameId_ = ""
    executionContextId_ = 0
End Sub

'Public
'第1引数で指定されるXpathの要素有無の確認
Private Function IWebDriver_IsExistsElementByXpath(xpath As String) As Boolean
    ReCreateIsolatedFrameWorld_IfFlagOn
    
    Dim objectId As String
    objectId = GetObjectIdOfElementByXpath(xpath)
    If objectId = "" Then
        IWebDriver_IsExistsElementByXpath = False
    Else
        IWebDriver_IsExistsElementByXpath = True
    End If
End Function

'Public
'Xpath(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearByXpath(xpath As String, maxWaitTime_ms As Long)
    ReCreateIsolatedFrameWorld_IfFlagOn
    
    Dim startTime As Date: startTime = Timer()
    Do
        If IWebDriver_IsExistsElementByXpath(xpath) Then
            Exit Sub
        ElseIf CInt((Timer() - startTime) * 1000) > maxWaitTime_ms Then
            DisplayErrorAndEnd maxWaitTime_ms & _
                                         "ミリ秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
                                          xpath
        End If
        DoEvents: Sleep 500
    Loop
End Sub

'Public
'AlertダイアログのOKボタンを押して閉じる
Private Sub IWebDriver_PushOkButtonOnDialog()
    Dim jsonForSending  As String, responseJson As String
    'イベント受信するために一回無害なメソッドを送受信する。
    jsonForSending = CDPMethod_.GetTargets
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
    
    jsonForSending = CDPMethod_.handleJavaScriptDialog()
    responseJson = ProcessMessage(CDPMethod_.ID, jsonForSending, jsonHandler_)
    If InStr(responseJson, "No dialog is showing") > 0 Then GoTo quit
Exit Sub

quit:
    DisplayErrorAndEnd "ダイアログが表示されていません" & vbCrLf & _
                                   "※表示されるタイミング調整が必要かもしれません。"
End Sub

Private Sub IWebDriver_SleepByWindowsAPI(milliSecond As Long)
    Sleep milliSecond
End Sub

'共通エラー対応。メッセージ表示、かつインターネットハンドルを解放して強制終了とする｡
'※エラー発生時にデバッグモードに入り、ユーザーに任せると解放されない可能性が高いので強制終了とする。
Public Sub DisplayErrorAndEnd(displayText As String)
    displayText = displayText & vbCrLf & "処理を終了します。"
    MsgBox displayText, vbCritical, "エラー"
    webSocket_.Http.CloseHInternetHandles
    End
End Sub

'共通対応。
'IsolatedWorldしているフレームに更新が入る場合、イベントで検知してフラグが立つので、再度IsolatedWorldする処理。
'【使用しているメソッド】
'IWebDriver_GetElementByXpath
'IWebDriver_WaitForElementToAppearByXpath
'IWebDriver_IsExistsElementByXpath
Public Sub ReCreateIsolatedFrameWorld_IfFlagOn()
    DOMEnable '念のため一回無害なメソッドを投げてフレーム更新イベントの発生チェック。発生なし→終
    If events_.IsIsolatedFrameScheduledNavigation = False Then Exit Sub
    
    '更新が終わるまでループ
    If events_.IsIsolatedFrameScheduledNavigationEnd = False Then
        Dim Counter As Long
        Do
            Sleep 1000
            DOMEnable
            Counter = Counter + 1
            If Counter > 30 Then GoTo quit '最大約30秒待ってもダメだったら強制終了
        Loop Until events_.IsIsolatedFrameScheduledNavigationEnd
    End If
    
    executionContextId_ = CreateIsolatedFrameWorld(isolatedFrameId_)
    events_.IsIsolatedFrameScheduledNavigation = False
    events_.IsIsolatedFrameScheduledNavigationEnd = False
Exit Sub

quit:
DisplayErrorAndEnd "フレームの更新を検知しましたが、30秒経過しても更新が終わりませんでした。"
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChromeDriver"
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
Private basicInfo_ As a_0_x3_BasicInfos
Private events_ As a_0_x1_x2_CDPEventHandler

Public documentNodeId_ As Long
Public documentObjectId_ As String

Public isolatedFrameId_ As String
Public executionContextId_ As Long

Public Property Get WebSocketURL() As String
    WebSocketURL = basicInfo_.WebSocketURL
End Property
Public Property Get UrlPath() As String
    UrlPath = basicInfo_.WebSocketUrlPath
End Property
Public Property Get TargetId() As String
    TargetId = basicInfo_.TargetId
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
    Set basicInfo_ = New a_0_x3_BasicInfos: basicInfo_.Init jsonhandler
    Set webSocket_ = New a_0_x1_WebSocketCommunicator
    If webSocket_.Init(basicInfo_.Host, basicInfo_.WebSocketUrlPath, basicInfo_.DebuggingPort) = False Then End
    
    Set CDPMethod_ = New a_0_x2_CDPMethods
    Set events_ = New a_0_x1_x2_CDPEventHandler
    
    SetDiscoverTargets
    AttachToTarget basicInfo_.TargetId
    
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
            ElseIf InStr(responseJson, "error") > 0 Then
                GoTo quit
            End If
        End If
    Loop While (sendId <> responseId)
    
    ProcessMessage = responseJson
Exit Function

quit:
Dim errMsg As String
If responseJson = "" Then
    errMsg = "送信エラー発生：イミディエイト参照"
ElseIf responseJson = "受信エラー発生" Then
    errMsg = "受信エラー発生:イミディエイト参照"
Else
    errMsg = "クロームからのエラー通知あり:イミディエイト参照"
End If
MsgBox errMsg, vbCritical, "エラー"
webSocket_.Http.CloseHInternetHandles
End
End Function

'Browser系************************************************
Private Function SessionClose() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.SessionClose
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

'************************************************

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
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    GetDocumentNodeID = jsonHandler_.GetValueFromJson(responseJson, Array("result", "root", "nodeId"))
End Function

Private Function ChangeObjectIdToNodeId(ObjectId As String) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.RequestNode(ObjectId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    
    Dim nodeId As String
    nodeId = jsonHandler_.GetValueFromJson(responseJson, Array("result", "nodeId"))
    ChangeObjectIdToNodeId = nodeId
End Function

Private Function ChangeNodeIdToObjectId(nodeId As Long) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.ResolveNode(nodeId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    
    Dim ObjectId As String
    ObjectId = jsonHandler_.GetValueFromJson(responseJson, Array("result", "objectId"))
    ChangeNodeIdToObjectId = ObjectId
End Function
'************************************************

'Element情報系************************************************
Private Function GetAttributes(nodeId As Long) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.GetAttributes(nodeId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    GetAttributes = responseJson
End Function

Private Function DescribeNode(id As Variant) As Variant
    Dim jsonForSending  As String
    jsonForSending = CDPMethod_.DescribeNode(id)
    DescribeNode = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function
'************************************************

'Runtime************************************************
'executionContextCreated イベントによる実行コンテキスト生成の報告を有効にします。
'レポートが有効になると、イベントは既存の実行コンテキストごとに直ちに送信されます。
Public Sub RuntimeEnable()
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.RunTime_Enable()
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Sub
'JavaScriptコード実行
Private Function ExecuteJavaScript_Inner(script As String) As String
    Dim jsonForSending  As String
    jsonForSending = CDPMethod_.JavaScritpEvaluate(script)
    ExecuteJavaScript_Inner = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

'JavaScript_Function実行
Public Function JavaScritpCallFunction(funcName As String, ObjectId As Variant, args As Variant) As String
    Dim jsonForSending  As String
    jsonForSending = CDPMethod_.JavaScritpCallFunction(funcName, ObjectId, args)
    JavaScritpCallFunction = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function
'************************************************
'Input Domain*******************************************************************
'キーイベントをページにディスパッチする。keyDown , keyUp, rawKeyDown, charのいずれか。
Public Function DispatchKeyEvent_DOMKeyCode(eventType As String, code As String) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.DispatchKeyEvent_DOMKeyCode(eventType, code)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

'*******************************************************************

'DOM操作系************************************************
Public Sub DOMEnable()
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.DOM_Enable()
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Sub

Public Sub Focus(ObjectId As String)
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.Focus(ObjectId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Sub

Public Sub FindElementByQuerySelector(Selector As String)
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.querySelector(Selector)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    
    Dim nodeId As Long
    nodeId = jsonHandler_.GetValueFromJson(responseJson, Array("result", "nodeId"))
    Dim nodeInfoJson As String
    nodeInfoJson = DescribeNode(nodeId)
End Sub

Public Function FindElementByID(id As String) As String
    
    
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
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

Public Function GetFrameTree() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.GetFrameTree()
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    GetFrameTree = responseJson
End Function

Public Function CreateIsolatedFrameWorld(FrameId As String) As Long
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.CreateIsolatedFrameWorld(FrameId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    If InStr(responseJson, "error") > 0 Then
        CreateIsolatedFrameWorld = 0
    Else
        CreateIsolatedFrameWorld = _
            jsonHandler_.GetValueFromJson(responseJson, Array("result", "executionContextId"))
        Sleep 1000
    End If
End Function

Public Function SetDiscoverTargets() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.SetDiscoverTargets(basicInfo_.TargetId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

Public Function AttachToTarget(TargetId As String) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.AttachToTarget(TargetId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

Public Function ActivateTarget(TargetId As String) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.ActivateTarget(TargetId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

Public Function DetachFromTarget(TargetId As String) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.DetachFromTarget(TargetId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

Public Function TargetClose() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.TargetClose(basicInfo_.TargetId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function
Public Function GetTargetInfo() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.GetTargets()
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function





















'以下､インターフェイス公開メソッドおよびその関連メソッド
'************************************************
'Public
'現在のページのURL
Private Property Get IWebDriver_CurrentURL() As String
    SetDiscoverTargets  'これを使ってPage情報最新化
    Dim i As Long
    For i = 1 To events_.PageInfos.Count
        If events_.PageInfos(i)("targetId") = basicInfo_.TargetId Then
            IWebDriver_CurrentURL = events_.PageInfos(i)("url")
            Exit Property
        End If
    Next
End Property



'Public
'第1引数で指定のURLを開く
Private Sub IWebDriver_URLOpen(url As String)
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.URLOpen(url)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    If InStr(responseJson, "invalid URL") > 0 Then GoTo quit
    
    WaitingForWebPageDisplay
    
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
Private Function WaitingForWebPageDisplay()
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
Private Function IWebDriver_ExecuteJavaScript(script As String) As String
    Dim responseJson As String
    responseJson = ExecuteJavaScript_Inner(script)
    IWebDriver_ExecuteJavaScript = jsonHandler_.GetValueFromJson(responseJson, Array("result"))
End Function

'Public
'第1引数で指定されるXpathの要素を取得し、エレメントObjectとして返却
Private Function IWebDriver_GetElementByXpath(xPath As String) As IWebElement
    ReCreateIsolatedFrameWorld_IfFlagOn
    
    If InStr(xPath, """") > 0 Then xPath = Replace(xPath, """", "\""")
    
    Dim scriptCode As String, functionCode As String
    scriptCode = "document.evaluate('" & xPath & "', document, null, " & _
                        "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;"
    functionCode = "function(xpath){return  document.evaluate(xpath, document, null, " & _
                            "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;}"
            
    Dim ObjectId As String
    ObjectId = GetObjectIdOfElementBy(xPath, scriptCode, functionCode)
    If ObjectId = "" Then GoTo quit
    Set IWebDriver_GetElementByXpath = MakeElementObject(ObjectId)
Exit Function

quit:
DisplayErrorAndEnd "指定された以下Xpathの要素は存在しませんでした。" & vbCrLf & xPath
End Function
'Xpathで要素を探し、オブジェクトIDを返す
Private Function GetObjectIdOfElementBy(way As String, scriptCode As String, functionCode As String) As String
    Dim responseJson As String
    If executionContextId_ = 0 Then
        responseJson = ExecuteJavaScript_Inner(scriptCode)
    Else
        responseJson = JavaScritpCallFunction(functionCode, executionContextId_, way)
    End If
    
    GetObjectIdOfElementBy = _
        jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "objectId"))
End Function
'エレメントを作成し、戻り値として返却
Public Function MakeElementObject(ObjectId As String) As IWebElement
    DOMEnable
    
    
    Dim info  As String
    info = DescribeNode(ObjectId)
    If InStr(info, "attributes") > 0 Then
        Dim attributes As Object
        Set attributes = GetAttributesDic( _
                                jsonHandler_.GetValueFromJson(info, Array("result", "node", "attributes")))
    End If
    If InStr(info, "frameId") > 0 Then
        Dim FrameId As String
        FrameId = jsonHandler_.GetValueFromJson(info, Array("result", "node", "frameId"))
    End If
    Dim nodeId As Long
    nodeId = 1
    
    Dim element As WebElement: Set element = New WebElement
    element.Init Me, ObjectId, nodeId, attributes, FrameId
    Set MakeElementObject = element
End Function
Private Function GetAttributesDic(attributes As String) As Object
    Dim ary As Variant: ary = Split(attributes, ",")
    Dim attributesDic As Object: Set attributesDic = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(ary) To UBound(ary) - 1 Step 2
        attributesDic.add ary(i), ary(i + 1)
    Next
    Set GetAttributesDic = attributesDic
End Function

'Public
Private Function IWebDriver_GetElementByName(Name As String) As IWebElement
    ReCreateIsolatedFrameWorld_IfFlagOn
    
    Dim scriptCode As String, functionCode As String
    scriptCode = "document.getElementsByName('" & Name & "')[0];"
    functionCode = "function(name){return  document.getElementsByName(name)[0];}"
    
    Dim ObjectId As String
    ObjectId = GetObjectIdOfElementBy(Name, scriptCode, functionCode)
    If ObjectId = "" Then GoTo quit
    Set IWebDriver_GetElementByName = MakeElementObject(ObjectId)
Exit Function

quit:
DisplayErrorAndEnd "指定された以下Nameの要素は存在しませんでした。" & vbCrLf & Name
End Function

'Public
Private Function IWebDriver_GetElementsByTagName(tagName As String) As IWebElements
    
    Dim elemsCount As Long
    elemsCount = GetElementsCount(tagName)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    Dim index As Long
    For index = 0 To elemsCount - 1
        elements.add CreateElementsItemByIndex(tagName, index)
    Next
    
    Dim elems As WebElements: Set elems = New WebElements
    elems.Init elements
    
    Set IWebDriver_GetElementsByTagName = elems
Exit Function

quit:
DisplayErrorAndEnd "指定された以下TagNameの要素は存在しませんでした。" & vbCrLf & tagName
End Function

Private Function GetElementsCount(tagName As String) As Long
    Dim javaScriptCode As String, responseJson As String
    If executionContextId_ = 0 Then
        javaScriptCode = "document.getElementsByTagName('" & tagName & "').length;"
        responseJson = ExecuteJavaScript_Inner(javaScriptCode)
    Else
        javaScriptCode = _
        "function(tagName){return  document.getElementsByTagName('" & tagName & "').length;}"
        responseJson = JavaScritpCallFunction(javaScriptCode, executionContextId_, tagName)
    End If
    GetElementsCount = jsonHandler_.GetValueFromJson(responseJson, Array("result", "result", "value"))
End Function

Private Function CreateElementsItemByIndex(tagName As String, index As Long) As IWebElement
    Dim javaScriptCode As String, responseJson As String
    Dim ObjectId As String
    Dim elem As IWebElement
    
    If executionContextId_ = 0 Then
        javaScriptCode = "document.getElementsByTagName('" & tagName & "')[" & index & "];"
        responseJson = ExecuteJavaScript_Inner(javaScriptCode)
    Else
        javaScriptCode = _
        "function(tagName){return  document.getElementsByTagName(tagName)[" & index & "];}"
        responseJson = JavaScritpCallFunction(javaScriptCode, executionContextId_, tagName)
    End If
    ObjectId = jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "objectId"))
    Set CreateElementsItemByIndex = MakeElementObject(ObjectId)
End Function


'Public
'第1引数で指定される値のフレームにフレームを切替
Private Sub IWebDriver_SwitchFrameByNameOrUrlOrIndex(value As Variant)
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim FrameId As String
    If IsNumeric(value) Then
        If frameInfos.Count < value Then GoTo quit
        FrameId = jsonHandler_.GetValueFromJson(frameInfos(value), Array("id"))
    Else
        FrameId = GetTargetFrameId(frameInfos, value)
    End If
    If FrameId = "" Then GoTo quit
    
    executionContextId_ = CreateIsolatedFrameWorld(FrameId)
    isolatedFrameId_ = FrameId
    
Exit Sub

quit:
DisplayErrorAndEnd "指定されたframeは存在しません。" & vbCrLf & value & vbCrLf & vbCrLf & _
                              "※もしUrl指定の場合、Chromeが情報を返さないケースがあります。" & vbCrLf & _
                              "   とれないようであればIndex、またはあればName指定を推奨します。"
End Sub

'frameツリーを取得し、各フレームの情報（JSONオブジェクト文字列）をコレクションとして集約
'frameTreeはresult→frameTree
'→frame(Pageの親フレーム)とChildFrames(配列の中にPageの子frame情報(オブジェクト)）の構成
Private Function GetFrameInfos() As Collection
    Dim collect As Collection: Set collect = New Collection
    Dim json As String: json = GetFrameTree
'    collect.add jsonHandler_.GetObjectFromJson(json, Array("result", "frameTree", "frame"))
    
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        'まずchildFramesから一つのFrame情報（オブジェクト）取得
        frmInfo = jsonHandler_.GetObjectFromJson(json, _
                        Array("result", "frameTree", "childFrames", Count))
        If frmInfo = "" Then Exit Do
        'そのオブジェクトの中のframeプロパティ(オブジェクト)を取得し、コレクションに加える
        frminfo2 = jsonHandler_.GetObjectFromJson(frmInfo, Array("frame"))
        collect.add frminfo2
        'childFramesプロパティがある場合、再帰構造なので再帰関数で同様の処理を回す
        If InStr(frmInfo, "childFrames") > 0 Then
            GetChildFrameInfo collect, frmInfo
        End If
        Count = Count + 1
    Loop
    Set GetFrameInfos = collect
End Function
'childFrameがある場合、再帰処理
Public Sub GetChildFrameInfo(collect As Collection, json As String)
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        frmInfo = jsonHandler_.GetObjectFromJson(json, _
                        Array("childFrames", Count))
        If frmInfo = "" Then Exit Do
        frminfo2 = jsonHandler_.GetObjectFromJson(frmInfo, Array("frame"))
        collect.add frminfo2
        If InStr(frmInfo, "childFrames") > 0 Then
            GetChildFrameInfo collect, frmInfo
        End If
        Count = Count + 1
    Loop
End Sub
'各フレームのname,urlと指定された値を比較。一致するフレームのIdを返す。
Private Function GetTargetFrameId(frameInfos As Collection, value As Variant) As String
    Dim frameInfo As Variant
    Dim Name As String, url As String
    For Each frameInfo In frameInfos
        Name = jsonHandler_.GetValueFromJson(frameInfo, Array("name"))
        url = jsonHandler_.GetValueFromJson(frameInfo, Array("url"))
        If Name = value Or url = value Then
            GetTargetFrameId = jsonHandler_.GetValueFromJson(frameInfo, Array("id"))
            Exit Function
        End If
    Next
End Function

'Public
'切り替え可能なフレームの情報をイミディエイトウィンドウに表示
Private Sub IWebDriver_DebugPrintAccessibleFrameInfo()
    PageEnable
    RuntimeEnable
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim i As Long
    For i = 1 To frameInfos.Count
        Debug.Print "index：" & i & _
                           "　name：" & jsonHandler_.GetValueFromJson(frameInfos(i), Array("name")) & _
                           "　url：" & jsonHandler_.GetValueFromJson(frameInfos(i), Array("url"))
    Next
End Sub

'Public
Private Sub IWebDriver_SwitchFrameByIframeElement(IframeElement As IWebElement)
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim elem As WebElement: Set elem = IframeElement
    If elem.FrameId = "" Then
        DisplayErrorAndEnd "指定されたElementがIframeではありません" & vbCrLf & vbCrLf & _
                                      "Iframe要素を引数に渡して下さい" & vbCrLf
    End If
    
'    events_.ForDebug_ShowtargetInfo
    GetTargetInfo
    SetDiscoverTargets
'    events_.ForDebug_ShowtargetInfo
    'frameIdがframeIdならIsolatedWorldで対応
    executionContextId_ = CreateIsolatedFrameWorld(elem.FrameId)
    If executionContextId_ <> 0 Then
        isolatedFrameId_ = elem.FrameId
        Exit Sub
    Else
        DisplayErrorAndEnd "このIframeはこのメソッドでは切り替えができません。" & vbCrLf & vbCrLf & _
                                      "※" & vbCrLf
    End If
    
    
'    テスト
'    Dim docElem As WebElement
'    Dim functionCode As String, responseJson As String, docObjectId As String
'    functionCode = "function(name){return  this.contentDocument;}"
'    responseJson = JavaScritpCallFunction(functionCode, elem.ObjectId, "")
'    docObjectId = jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "objectId"))
'    Dim info  As String
'    info = DescribeNode(docObjectId)
    
    
'    frameIdがtargetIdなら接続切り替えで対応
'    basicInfo_.Update elem.FrameId, jsonHandler_, False
'    isolatedFrameId_ = ""
'
'    webSocket_.CloseWebSocket
'    Set webSocket_ = Nothing
'    Set webSocket_ = New a_0_x1_WebSocketCommunicator
'    If webSocket_.Init(basicInfo_.Host, basicInfo_.WebSocketUrlPath, _
'                                basicInfo_.DebuggingPort) = False Then
'        DisplayErrorAndEnd "SwitchIframeに失敗しました" & vbCrLf
'    End If
'
'    AttachToTarget elem.FrameId
'    PageEnable
'    DOMEnable
'
'    documentObjectId_ = GetDocumentObjectID
'    documentNodeId_ = GetDocumentNodeID

End Sub

'Public
'デフォルトのフレームに戻す
Private Sub IWebDriver_SwitchFrameToDefault()
    If isolatedFrameId_ <> "" Then
        isolatedFrameId_ = ""
        executionContextId_ = 0
    End If
    
'    If basicInfo_.IsTargetPageDefault = False Then
'        SwitchTargetToPageDefault
'    End If
End Sub
'Private Sub SwitchTargetToPageDefault()
'    Dim newTargetId As String
'    newTargetId = basicInfo_.PageDefaultTargetId
'    basicInfo_.Update newTargetId, jsonHandler_, False
'
'    webSocket_.CloseWebSocket
'    Set webSocket_ = Nothing
'    Set webSocket_ = New a_0_x1_WebSocketCommunicator
'    If webSocket_.Init(basicInfo_.Host, basicInfo_.WebSocketUrlPath, basicInfo_.DebuggingPort) = False Then End
'
'    ActivateTarget basicInfo_.TargetId
'    PageEnable
'    DOMEnable
'
'    documentObjectId_ = GetDocumentObjectID
'    documentNodeId_ = GetDocumentNodeID
'End Sub

'Public
'第1引数で指定されるXpathの要素有無の確認
Private Function IWebDriver_IsExistsElementByXpath(xPath As String) As Boolean
    ReCreateIsolatedFrameWorld_IfFlagOn
    
    Dim scriptCode As String, functionCode As String
    scriptCode = "document.evaluate('" & xPath & "', document, null, " & _
                        "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;"
    functionCode = "function(xpath){return  document.evaluate(xpath, document, null, " & _
                            "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;}"
    Dim ObjectId As String
    ObjectId = GetObjectIdOfElementBy(xPath, scriptCode, functionCode)
    If ObjectId = "" Then
        IWebDriver_IsExistsElementByXpath = False
    Else
        IWebDriver_IsExistsElementByXpath = True
    End If
End Function

'Public
'Xpath(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearByXpath(xPath As String, maxWaitTime_sec As Long)
    
    Dim startTime As Date: startTime = Timer()
    Do
        If IWebDriver_IsExistsElementByXpath(xPath) Then
            Exit Sub
        ElseIf CInt((Timer() - startTime)) > maxWaitTime_sec Then
            DisplayErrorAndEnd maxWaitTime_sec & _
                                         "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
                                          xPath
        End If
        ReCreateIsolatedFrameWorld_IfFlagOn
        DoEvents: Sleep 500
    Loop
End Sub

'Public
'AlertダイアログのOKボタンを押して閉じる
Private Sub IWebDriver_PushOkButtonOnDialog()
    Dim jsonForSending  As String, responseJson As String
    'イベント受信するために一回無害なメソッドを送受信する。
    jsonForSending = CDPMethod_.GetTargets
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    
    jsonForSending = CDPMethod_.handleJavaScriptDialog()
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
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
'IWebDriver_GetElementsByTagName
Public Sub ReCreateIsolatedFrameWorld_IfFlagOn()
'event検知がうまくいかないため、一旦DefaultFrame以外の操作継続時には毎回CreateIsolatedFrameWorldする
    If isolatedFrameId_ <> "" Then
        GetFrameTree '一旦無害なメソッドを投げてイベント発生チェック
        executionContextId_ = CreateIsolatedFrameWorld(isolatedFrameId_)
    End If
    
'    DOMEnable '念のため一回無害なメソッドを投げてフレーム更新イベントの発生チェック。発生なし→終
'    If events_.IsIsolatedFrameScheduledNavigation = False Then Exit Sub
'
'    '更新が終わるまでループ
'    If events_.IsIsolatedFrameScheduledNavigationEnd = False Then
'        Dim Counter As Long
'        Do
'            Sleep 1000
'            DOMEnable
'            Counter = Counter + 1
'            If Counter > 30 Then GoTo quit '最大約30秒待ってもダメだったら強制終了
'        Loop Until events_.IsIsolatedFrameScheduledNavigationEnd
'    End If
'
'    executionContextId_ = CreateIsolatedFrameWorld(isolatedFrameId_)
''    events_.IsIsolatedFrameScheduledNavigation = False
''    events_.IsIsolatedFrameScheduledNavigation = False
Exit Sub

'quit:
'DisplayErrorAndEnd "フレームの更新を検知しましたが、30秒経過しても更新が終わりませんでした。"
End Sub

'右隣のタブへ移動
Private Sub IWebDriver_SwitchTabByIndex(index As Long)
    Dim newTargetId As String
    newTargetId = events_.PageInfos(index)("targetId")
    basicInfo_.Update newTargetId, jsonHandler_, True
    
    webSocket_.CloseWebSocket
    Set webSocket_ = Nothing
    Set webSocket_ = New a_0_x1_WebSocketCommunicator
    If webSocket_.Init(basicInfo_.Host, basicInfo_.WebSocketUrlPath, basicInfo_.DebuggingPort) = False Then End
    
    ActivateTarget basicInfo_.TargetId
    AttachToTarget basicInfo_.TargetId
    PageEnable
    DOMEnable
    
    documentObjectId_ = GetDocumentObjectID
    documentNodeId_ = GetDocumentNodeID
End Sub

Private Sub IWebDriver_CloseWindow()
    SessionClose
End Sub

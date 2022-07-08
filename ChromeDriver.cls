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

Private Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const CLASS_NAME As String = "Chrome_WidgetWin_1"
Private Const CAPTION_NAME As String = "data:, - Google Chrome"

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32.dll" ( _
        ByVal hwnd As LongPtr, ByRef ProcessId As Long) As Long

Private Declare PtrSafe Sub SetForeGroundWindow Lib "User32" Alias "SetForegroundWindow" (ByVal ms As LongPtr)

Private Const CHROME_EXE_REG_KEY As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\"
Private Const CHROME_EXE As String = """chrome.exe"""
Private Const DEBUG_MODE_OPTION As String = " --remote-debugging-Port=9222"
Private Const USER_DATA_DIR_OPTION As String = " --user-data-dir="
Private Const ENABLE_AUTOMATION_OPTION As String = " --enable-automation"

Private userDataDir As String

Private webSocket_ As a_0_x1_WebSocketCommunicator
Private CDPMethod_ As a_0_x2_CDPMethods
Private json_ As a_0_x4_JSONHandler
Private basicInfo_ As a_0_x3_BasicInfos
Private events_ As a_0_x1_x2_CDPEventHandler

Private processId_ As Long
Private chromeHwnd_ As LongPtr

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
Public Property Get Json() As a_0_x4_JSONHandler
    Set Json = json_
End Property
Public Property Get CDPEvents() As a_0_x1_x2_CDPEventHandler
    Set CDPEvents = events_
End Property


Private Sub Class_Initialize()
    
    userDataDir = ThisWorkbook.Path & "\マクロ実行で作成されるフォルダ_削除OK\"
    If Dir(userDataDir, vbDirectory) = "" Then MkDir userDataDir
    
    processId_ = StartChrome
    '検証したところ、複数Chromeが立ち上がっていても最新のWindowHandleが返ってくる模様。これで一旦問題なしか？
    chromeHwnd_ = FindWindow(CLASS_NAME, vbNullString)
    
    Set json_ = New a_0_x4_JSONHandler
    Set basicInfo_ = New a_0_x3_BasicInfos: basicInfo_.Init json_
    Set webSocket_ = New a_0_x1_WebSocketCommunicator
    If webSocket_.Init(basicInfo_.Host, basicInfo_.WebSocketUrlPath, basicInfo_.DebuggingPort) = False Then End
    
    Set CDPMethod_ = New a_0_x2_CDPMethods
    Set events_ = New a_0_x1_x2_CDPEventHandler
    
    SetDiscoverTargets
    AttachToTarget basicInfo_.TargetId
End Sub

Private Function StartChrome() As Long
    Dim chromeExe As String
    Dim objExec As Object
    With CreateObject("WScript.Shell")
        chromeExe = """" & .RegRead(CHROME_EXE_REG_KEY) & """"
        Set objExec = .Exec( _
                            chromeExe & DEBUG_MODE_OPTION & ENABLE_AUTOMATION_OPTION _
                            & USER_DATA_DIR_OPTION & "" & userDataDir & "")
        StartChrome = objExec.ProcessId
        Set objExec = Nothing
    End With
    Sleep 2000
End Function

'message処理送信＆受信
'受信時には、送信に対する受信ではなく、クロームからのイベント通知を受信する場合あり。
'イベント通知の場合、responsにidが含まれないため、idの有無で送信に対する受信なのかイベント通知なのかを判断
'補足:送信自体は成功しても、送信したJsonに不備がある場合も、responsにidが含まれない。
Public Function ProcessMessage( _
    sendId As Long, sendMsg As String, Json As a_0_x4_JSONHandler) As String
    
    Debug.Print vbCrLf & ">>>>>>" & "送信内容：" & sendMsg
    If webSocket_.SendMessage(sendMsg) = False Then GoTo quit

    Dim response As String, responseId As Variant
    Do
        response = webSocket_.GetMessage
        If response = "受信エラー発生" Then GoTo quit
        
        responseId = Json.GetValueFrom(response, Array("id"))
        If responseId = "" Then
            If InStr(response, "method") > 0 Then
                events_.GetInfo response, Me
                
                'ダイアログ出現時のみの特殊対応。ダイアログが出現したら、利用者は必ず次は(HandleAlertDialogを使用し)
                'ダイアログを消すと思われるので、その前提でこのループから抜ける。（他の操作をされるとアウト）
                '理由：ダイアログによっては出現時に処理をしないと、それ以降クロームがメッセージを返さなくなるケースがある。
                'そのダイアログの場合にこの受信ループを継続すると、クロームから返信なし⇒受信から抜け出せなくなって完全に止まるため。
                If events_.DialogInfoDic("IsExistDialog") Then Exit Function
                
            ElseIf InStr(response, "error") > 0 Then
                GoTo quit
            End If
        End If
    Debug.Print "受信：" & response
    Loop While (sendId <> responseId)
    
    ProcessMessage = response
Exit Function

quit:
Dim errMsg As String
If response = "" Then
    errMsg = "送信エラー発生：イミディエイト参照"
ElseIf response = "受信エラー発生" Then
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
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.SessionClose
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function

'************************************************

'DOM(ElementId)系************************************************
'CDPMethodにないのでJavascript実行
Private Function GetDocumentObjectID() As String
    Dim responseMsg  As String
    responseMsg = ExecuteJavaScript_Inner("document;")
    documentObjectId_ = json_.GetValueFrom(responseMsg, Array("result", "result", "objectId"))
    
    GetDocumentObjectID = documentObjectId_
End Function

Private Function GetDocumentNodeID() As Long
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.GetDocumentID
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
    documentNodeId_ = json_.GetValueFrom(responseMsg, Array("result", "root", "nodeId"))
    GetDocumentNodeID = documentNodeId_
End Function

Private Function ChangeObjectIdToNodeId(objectId As String) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.RequestNode(objectId)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
    ChangeObjectIdToNodeId = json_.GetValueFrom(responseMsg, Array("result", "nodeId"))
End Function

Private Function ChangeNodeIdToObjectId(nodeId As Long) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.ResolveNode(nodeId)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
    ChangeNodeIdToObjectId = json_.GetValueFrom(responseMsg, Array("result", "objectId"))
End Function
'************************************************

'DOM(Element情報)系************************************************
Private Function GetAttributes(nodeId As Long) As String
    Dim sendMsg  As String
    sendMsg = CDPMethod_.GetAttributes(nodeId)
    GetAttributes = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function

Private Function DescribeNode(id As Variant) As Variant
    Dim sendMsg  As String
    sendMsg = CDPMethod_.DescribeNode(id)
    DescribeNode = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function
'************************************************

'Runtime************************************************
'executionContextCreated イベントによる実行コンテキスト生成の報告を有効にします。
'レポートが有効になると、イベントは既存の実行コンテキストごとに直ちに送信されます。
Public Sub RuntimeEnable()
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.RunTime_Enable()
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Sub
'JavaScriptコード実行
Private Function ExecuteJavaScript_Inner(script As String) As String
    Dim sendMsg  As String
    sendMsg = CDPMethod_.JavaScritpEvaluate(script)
    ExecuteJavaScript_Inner = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function

'JavaScript_Function実行
Public Function JavaScritpCallFunction(funcName As String, objectId As Variant, args As Variant) As String
    Dim sendMsg  As String
    sendMsg = CDPMethod_.JavaScritpCallFunction(funcName, objectId, args)
    JavaScritpCallFunction = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function
'************************************************
'Input Domain*******************************************************************
'キーイベントをページにディスパッチする。keyDown , keyUp, rawKeyDown, charのいずれか。
Public Function DispatchKeyEvent_DOMKeyCode(eventType As String, code As String) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.DispatchKeyEvent_DOMKeyCode(eventType, code)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function

'マウスイベントをxy座標にディスパッチする。mousePressed, mouseReleased, mouseMoved, mouseWheelのいずれか。
Public Function DispatchMouseEvent(eventType As String, x As Long, y As Long, _
                                                        button As String, clickCount As Long, _
                                                        Optional modifiers As Long = 0, _
                                                        Optional deltaX As Long = 0, Optional deltaY As Long = 0) As String
                                                        
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.DispatchMouseEvent(eventType, x, y, button, clickCount, _
                                                                                    modifiers, deltaX, deltaY)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function

'*******************************************************************

'DOM操作系************************************************
Public Sub DOMEnable()
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.DOM_Enable()
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Sub

Public Sub Focus(objectId As String)
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.Focus(objectId)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Sub

Public Function GetBoxModel(objectId As String) As Variant
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.GetBoxModel(objectId)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
    Dim result As Variant
    result = json_.GetValueFrom(responseMsg, Array("result", "model", "content"))
    result = Split(result, ",")
    Dim width As Long, height As Long
    width = json_.GetValueFrom(responseMsg, Array("result", "model", "width"))
    height = json_.GetValueFrom(responseMsg, Array("result", "model", "height"))
    Dim xy(1) As Variant
    xy(0) = result(0) + (width / 2)
    xy(1) = result(1) + (height / 2)
    GetBoxModel = xy
End Function

Public Sub SetFileInputFiles(objectId As Variant, files As String)
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.SetFileInputFiles(objectId, files)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Sub

Public Sub FindElementByQuerySelector(Selector As String)
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.querySelector(Selector)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
    
    Dim nodeId As Long
    nodeId = json_.GetValueFrom(responseMsg, Array("result", "nodeId"))
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
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.PageEnable()
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function

Public Function GetFrameTree() As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.GetFrameTree()
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
    GetFrameTree = responseMsg
End Function

Public Function CreateIsolatedFrameWorld(FrameId As String) As Long
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.CreateIsolatedFrameWorld(FrameId)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
    If InStr(responseMsg, "error") > 0 Then
        CreateIsolatedFrameWorld = 0
    Else
        CreateIsolatedFrameWorld = _
            json_.GetValueFrom(responseMsg, Array("result", "executionContextId"))
        Sleep 200
    End If
End Function

Public Function SetDiscoverTargets() As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.SetDiscoverTargets(basicInfo_.TargetId)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function

Public Function AttachToTarget(TargetId As String) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.AttachToTarget(TargetId)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function

Public Function ActivateTarget(TargetId As String) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.ActivateTarget(TargetId)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function

Public Function DetachFromTarget(TargetId As String) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.DetachFromTarget(TargetId)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function

Public Function TargetClose() As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.TargetClose(basicInfo_.TargetId)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function
Public Function GetTargetInfo() As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.GetTargets()
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
End Function

Public Sub SleepByWindowsAPIInner(ms As Long)
    Sleep ms
End Sub





















'以下､インターフェイス公開メソッドおよびその関連メソッド
'************************************************
'Public
'現在のページのURL
Private Property Get IWebDriver_URL() As String
    SetDiscoverTargets  'これを使ってPage情報最新化
    Dim i As Long
    For i = 1 To events_.PageInfos.Count
        If events_.PageInfos(i)("targetId") = basicInfo_.TargetId Then
            IWebDriver_URL = events_.PageInfos(i)("url")
            Exit Property
        End If
    Next
End Property





'Public
'第1引数で指定のURLを開く
Private Sub IWebDriver_OpenURL(targetURL As String)
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.URLOpen(targetURL)
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
    If InStr(responseMsg, "invalid URL") > 0 Then GoTo quit
    
    WaitingForWebPageDisplay
    
    SetDiscoverTargets
    PageEnable
    DOMEnable
    documentObjectId_ = GetDocumentObjectID
    documentNodeId_ = GetDocumentNodeID
Exit Sub

quit:
DisplayErrorAndEnd "＜OpenURL＞" & vbCrLf & _
                                "指定された以下URLは無効なURLです。" & vbCrLf & _
                                targetURL & vbCrLf
End Sub

'ページ表示待ちループ
Private Function WaitingForWebPageDisplay()
    Dim javaScriptCode As String, responseMsg As String, state As String
    Do
        javaScriptCode = "document.readyState;"
        responseMsg = ExecuteJavaScript_Inner(javaScriptCode)
        state = json_.GetValueFrom(responseMsg, Array("result", "result", "value"))
        DoEvents
        Sleep 800
    Loop Until (state = "complete")
End Function

'Public
'第1引数のJavaScriptコードの実行
Private Function IWebDriver_ExecuteScript(script As String) As String
    Dim responseMsg As String
    responseMsg = ExecuteJavaScript_Inner(script)
    IWebDriver_ExecuteScript = json_.GetValueFrom(responseMsg, Array("result"))
End Function

'Public
'第1引数で指定されるXpathの要素を取得し、エレメントObjectとして返却
Private Function IWebDriver_FindElementByXpath(xPath As String) As IWebElement
    
    If InStr(xPath, """") > 0 Then xPath = Replace(xPath, """", "\""")
    
    Dim scriptCode As String, FunctionCode As String
    scriptCode = "document.evaluate('" & xPath & "', document, null, " & _
                        "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;"
    FunctionCode = "function(xpath){return  document.evaluate(xpath, document, null, " & _
                            "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;}"
            
    Dim objectId As String
    objectId = GetObjectIdOfElementBy(xPath, scriptCode, FunctionCode)
    If objectId = "" Then GoTo quit
    Set IWebDriver_FindElementByXpath = MakeElementObject(objectId)
Exit Function

quit:
DisplayErrorAndEnd "＜FindElementByXpath＞" & vbCrLf & _
                                "指定された以下Xpathの要素は存在しませんでした。" & vbCrLf & _
                                xPath & vbCrLf & vbCrLf & _
                                "※引数のXPathにはChromeの検証ツールで「完全なXPath」が" & vbCrLf & _
                                "　 コピペできるので、それを指定してください。" & vbCrLf & vbCrLf & _
                                "※XPathが正しくとも要素の出現前に取得しようとすると失敗します。" & vbCrLf & _
                                "　 必要であればWaitForElementToAppearByXpathで" & vbCrLf & _
                                "　 要素の出現を確認する処理を入れましょう。" & vbCrLf
End Function
'Xpathで要素を探し、オブジェクトIDを返す
Private Function GetObjectIdOfElementBy(way As String, scriptCode As String, FunctionCode As String) As String
    Dim responseMsg As String
    If executionContextId_ = 0 Then
        responseMsg = ExecuteJavaScript_Inner(scriptCode)
    Else
        responseMsg = JavaScritpCallFunction(FunctionCode, executionContextId_, way)
        If InStr(responseMsg, "Cannot find context with specified id") > 0 Then
            ReCreateIsolatedFrameWorld
            responseMsg = JavaScritpCallFunction(FunctionCode, executionContextId_, way)
        End If
    End If
    
    GetObjectIdOfElementBy = _
        json_.GetValueFrom(responseMsg, Array("result", "result", "objectId"))
End Function
'エレメントを作成し、戻り値として返却
'エレメント内容：ObjectId,NodeId,Attributes(辞書),フレームの場合のみframeId追加
Public Function MakeElementObject(objectId As String) As IWebElement
    
    Dim nodeId As Long
    nodeId = ChangeObjectIdToNodeId(objectId)
    
    Dim info  As String
    info = DescribeNode(objectId)
    If InStr(info, "attributes") > 0 Then
        Dim attributes As Object
        Set attributes = MakeAttributesDic( _
                                json_.GetValueFrom(info, Array("result", "node", "attributes")))
    End If
    If InStr(info, "frameId") > 0 Then
        Dim FrameId As String
        FrameId = json_.GetValueFrom(info, Array("result", "node", "frameId"))
    End If
    
    
    Dim element As WebElement: Set element = New WebElement
    element.Init Me, objectId, nodeId, attributes, FrameId
    Set MakeElementObject = element
End Function
'elementの属性情報を辞書にして返す。
Private Function MakeAttributesDic(attributes As String) As Object
    Dim ary As Variant: ary = Split(attributes, ",")
    Dim attributesDic As Object: Set attributesDic = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(ary) To UBound(ary) - 1 Step 2
        If attributesDic.exists(ary(i)) = False Then
            attributesDic.add ary(i), ary(i + 1)
        End If
    Next
    Set MakeAttributesDic = attributesDic
End Function

'Public
Private Function IWebDriver_FindElementByName(Name As String) As IWebElement
    
    Dim scriptCode As String, FunctionCode As String
    scriptCode = "document.getElementsByName('" & Name & "')[0];"
    FunctionCode = "function(name){return  document.getElementsByName(name)[0];}"
    
    Dim objectId As String
    objectId = GetObjectIdOfElementBy(Name, scriptCode, FunctionCode)
    If objectId = "" Then GoTo quit
    Set IWebDriver_FindElementByName = MakeElementObject(objectId)
Exit Function

quit:
DisplayErrorAndEnd "＜FindElementByName＞" & vbCrLf & _
                                "指定された以下Nameの要素は存在しませんでした。" & vbCrLf & _
                                Name & vbCrLf & vbCrLf & _
                                "※要素の出現前に取得しようとすると失敗します。" & vbCrLf & _
                                "　 必要であればWaitForElementToAppearByXpathで" & vbCrLf & _
                                "　 要素の出現を確認する処理を入れましょう。" & vbCrLf
End Function

'Public
'tag名で要素コレクションを取得（要素グループのLengthを取得し、Length分要素作成ループ）
Private Function IWebDriver_FindElementsByTag(tagName As String) As IWebElements
    Dim scriptCode As String, FunctionCode As String
    scriptCode = "document.getElementsByTagName('" & tagName & "').length;"
    FunctionCode = "function(tagName){return  document.getElementsByTagName('" & tagName & "').length;}"
    Dim elemsCount As Long
    elemsCount = GetElementsCountBy(tagName, scriptCode, FunctionCode)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    Dim objectId As String
    Dim index As Long
    For index = 0 To elemsCount - 1
        scriptCode = "document.getElementsByTagName('" & tagName & "')[" & index & "];"
        FunctionCode = "function(tagName){return  document.getElementsByTagName(tagName)[" & index & "];}"
        objectId = GetObjectIdOfElementBy(tagName, scriptCode, FunctionCode)
        elements.add MakeElementObject(objectId)
    Next
    
    Dim elems As WebElements: Set elems = New WebElements
    elems.Init elements
    
    Set IWebDriver_FindElementsByTag = elems
Exit Function

quit:
DisplayErrorAndEnd "＜FindElementsByTag＞" & vbCrLf & _
                                "指定された以下TagNameの要素は存在しませんでした。" & vbCrLf & _
                                tagName & vbCrLf & vbCrLf & _
                                "※要素の出現前に取得しようとすると失敗します。" & vbCrLf & _
                                "　 必要であればWaitForElementToAppearByXpathで" & vbCrLf & _
                                "　 要素の出現を確認する処理を入れましょう。" & vbCrLf
End Function

Private Function GetElementsCountBy(way As String, scriptCode As String, FunctionCode As String) As Long
    Dim responseMsg As String
    If executionContextId_ = 0 Then
        responseMsg = ExecuteJavaScript_Inner(scriptCode)
    Else
        responseMsg = JavaScritpCallFunction(FunctionCode, executionContextId_, way)
        If InStr(responseMsg, "Cannot find context with specified id") > 0 Then
            ReCreateIsolatedFrameWorld
            responseMsg = JavaScritpCallFunction(FunctionCode, executionContextId_, way)
        End If
    End If
    GetElementsCountBy = json_.GetValueFrom(responseMsg, Array("result", "result", "value"))
End Function


'Public
'第1引数で指定される値のフレームにフレームを切替
Private Sub IWebDriver_SwitchFrameByNameOrUrlOrIndex(value As Variant)
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim FrameId As String
    If IsNumeric(value) Then
        If frameInfos.Count < value Then GoTo quit
        FrameId = json_.GetValueFrom(frameInfos(value), Array("id"))
    Else
        FrameId = GetTargetFrameId(frameInfos, value)
    End If
    If FrameId = "" Then GoTo quit
    
    executionContextId_ = CreateIsolatedFrameWorld(FrameId)
    isolatedFrameId_ = FrameId
    
Exit Sub

quit:
DisplayErrorAndEnd "＜SwitchFrameByNameOrUrlOrIndex＞" & vbCrLf & _
                                "指定された以下のframeは存在しません。" & vbCrLf & _
                                value & vbCrLf & vbCrLf & _
                                "※iframeにName属性があれば、Nameでの指定を推奨します。" & vbCrLf & vbCrLf & _
                                "※フレームのIndex番号、Name、URLを確認したい場合、" & vbCrLf & _
                                "　 DebugPrintAccessibleFrameInfoを使用してください。" & vbCrLf & _
                                "　 クロームから通知されるフレーム情報がイミディエイトに出力されます。" & vbCrLf & vbCrLf & _
                                "※このメソッドでうまくいかない場合、" & vbCrLf & _
                                "　 SwitchFrameByIframeElementの" & vbCrLf & _
                                "　 使用を検討してください。" & vbCrLf
End Sub

'frameツリーを取得し、各フレームの情報（JSONオブジェクト文字列）をコレクションとして集約
'frameTreeはresult→frameTree
'→frame(Pageの親フレーム)とChildFrames(配列の中にPageの子frame情報(オブジェクト)）の構成
Private Function GetFrameInfos() As Collection
    Dim collect As Collection: Set collect = New Collection
    Dim Json As String: Json = GetFrameTree
'    collect.add json_.GetObjectFrom(json, Array("result", "frameTree", "frame"))
    
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        'まずchildFramesから一つのFrame情報（オブジェクト）取得
        frmInfo = json_.GetObjectFrom(Json, _
                        Array("result", "frameTree", "childFrames", Count))
        If frmInfo = "" Then Exit Do
        'そのオブジェクトの中のframeプロパティ(オブジェクト)を取得し、コレクションに加える
        frminfo2 = json_.GetObjectFrom(frmInfo, Array("frame"))
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
Public Sub GetChildFrameInfo(collect As Collection, Json As String)
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        frmInfo = json_.GetObjectFrom(Json, _
                        Array("childFrames", Count))
        If frmInfo = "" Then Exit Do
        frminfo2 = json_.GetObjectFrom(frmInfo, Array("frame"))
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
    Dim Name As String, URL As String
    For Each frameInfo In frameInfos
        Name = json_.GetValueFrom(frameInfo, Array("name"))
        URL = json_.GetValueFrom(frameInfo, Array("url"))
        If Name = value Or URL = value Then
            GetTargetFrameId = json_.GetValueFrom(frameInfo, Array("id"))
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
                           "　name：" & json_.GetValueFrom(frameInfos(i), Array("name")) & _
                           "　url：" & json_.GetValueFrom(frameInfos(i), Array("url"))
    Next
End Sub

'Public
Private Sub IWebDriver_SwitchFrameByIframeElement(IframeElement As IWebElement)
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim elem As WebElement: Set elem = IframeElement
    If elem.FrameId = "" Then
        DisplayErrorAndEnd "＜SwitchFrameByIframeElement＞" & vbCrLf & _
                                      "指定されたElementがIframeではありません" & vbCrLf & vbCrLf & _
                                      "Iframe要素を引数に渡して下さい" & vbCrLf
    End If
    
'    events_.ForDebug_ShowtargetInfo
    GetTargetInfo
    SetDiscoverTargets
'    events_.ForDebug_ShowtargetInfo
    'frameElementのframeIdが正当なframeIdならIsolatedWorldで対応
    executionContextId_ = CreateIsolatedFrameWorld(elem.FrameId)
    If executionContextId_ <> 0 Then
        isolatedFrameId_ = elem.FrameId
        Exit Sub
    Else
        DisplayErrorAndEnd "＜SwitchFrameByIframeElement＞" & vbCrLf & _
                                       "このIframeはこのメソッドでは切り替えができません。" & vbCrLf & vbCrLf & _
                                      "※SwitchFrameByNameOrUrlOrIndexの" & vbCrLf & _
                                      "　 使用を検討してください。" & vbCrLf
    End If
    
    
'    テスト
'    Dim docElem As WebElement
'    Dim functionCode As String, responseMsg As String, docObjectId As String
'    functionCode = "function(name){return  this.contentDocument;}"
'    responseMsg = JavaScritpCallFunction(functionCode, elem.ObjectId, "")
'    docObjectId = json_.GetValueFrom(responseMsg, Array("result", "result", "objectId"))
'    Dim info  As String
'    info = DescribeNode(docObjectId)
    
    
'    frameIdがtargetIdなら接続切り替えで対応
'    basicInfo_.Update elem.FrameId, json_, False
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
'    basicInfo_.Update newTargetId, json_, False
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
    
    Dim scriptCode As String, FunctionCode As String
    scriptCode = "document.evaluate('" & xPath & "', document, null, " & _
                        "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;"
    FunctionCode = "function(xpath){return  document.evaluate(xpath, document, null, " & _
                            "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;}"
    Dim objectId As String
    objectId = GetObjectIdOfElementBy(xPath, scriptCode, FunctionCode)
    If objectId = "" Then
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
        DoEvents: Sleep 500
    Loop
End Sub

'Public
'AlertダイアログのOKボタンを押して閉じる
'★events_.DialogInfoDic("IsExistDialog")の判定、特殊対応を入れている箇所★
'①ProcessMessage②これ③WebElemets：Click
'制御を奪うDialogの場合、Dialog出現以降に、このCDPメソッド以外のCDPメソッドを投げると、
'制御が奪われているため、返信がなく詰む。
Private Sub IWebDriver_HandleAlertDialog()
    Dim sendMsg  As String, responseMsg As String
    If events_.DialogInfoDic("IsExistDialog") = False Then
        DOMEnable 'イベント受信状況更新のため、一回無害なメソッドを投げる
    End If
    
    sendMsg = CDPMethod_.HandleJavaScriptDialog()
    responseMsg = ProcessMessage(CDPMethod_.id, sendMsg, json_)
    If InStr(responseMsg, "No dialog is showing") > 0 Then GoTo quit
Exit Sub

quit:
    DisplayErrorAndEnd "<HandleAlertDialog>" & vbCrLf & _
                                   "ダイアログが表示されていません。" & vbCrLf & _
                                   "※このメソッドはAlertが表示されてから使用します。" & vbCrLf & _
                                   "　 必要があればタイミング調整処理を入れてください。" & vbCrLf
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
'再度CreateIsolatedWorldする処理。
'【使用しているメソッド】
'IWebDriver_GetElementByXpath
'IWebDriver_WaitForElementToAppearByXpath
'IWebDriver_IsExistsElementByXpath
'IWebDriver_GetElementsByTagName
Public Sub ReCreateIsolatedFrameWorld()
    If isolatedFrameId_ <> "" Then
        executionContextId_ = CreateIsolatedFrameWorld(isolatedFrameId_)
    End If
End Sub

'右隣のタブへ移動
Private Sub IWebDriver_SwitchTabByIndex(index As Long)
    Dim newTargetId As String
    newTargetId = events_.PageInfos(index)("targetId")
    basicInfo_.Update newTargetId, json_, True
    
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

Private Sub IWebDriver_ScrollWindow(yPixel As Long, Optional xPixel As Long = 0&)
    DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, xPixel, yPixel
End Sub

Private Sub IWebDriver_SetForeGroundWindow()
    SetForeGroundWindow chromeHwnd_
End Sub

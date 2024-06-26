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

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const CLASS_NAME As String = "Chrome_WidgetWin_1"
Private Const CAPTION_NAME As String = "新しいタブ - Google Chrome" '"data:, - Google Chrome"

'Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
'        ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
'

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "USER32.DLL" ( _
        ByVal hWnd As LongPtr, ByRef ProcessId As Long) As Long

Private Declare PtrSafe Sub SetForeGroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal ms As LongPtr)

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
Public Property Get json() As a_0_x4_JSONHandler
    Set json = json_
End Property
Public Property Get CDPEvents() As a_0_x1_x2_CDPEventHandler
    Set CDPEvents = events_
End Property


Private Sub Class_Initialize()
    
    userDataDir = ThisWorkbook.Path & "\マクロ実行専用UserDataフォルダ\"
    Dim isFirst As Boolean
    If Dir(userDataDir, vbDirectory) = "" Then
        MkDir userDataDir
        isFirst = True
    End If
    
    processId_ = StartChrome
    
    '適切に取得ができないケースあり、要改善（よってまだどこでもhwndは使用しない）
    chromeHwnd_ = FindWindow(CLASS_NAME, CAPTION_NAME)
    
    Set json_ = New a_0_x4_JSONHandler
    Set basicInfo_ = New a_0_x3_BasicInfos: basicInfo_.Init json_
    Set webSocket_ = New a_0_x1_WebSocketCommunicator
    If webSocket_.Init(basicInfo_.Host, basicInfo_.WebSocketUrlPath, basicInfo_.DebuggingPort) = False Then End
    
    Set CDPMethod_ = New a_0_x2_CDPMethods
    Set events_ = New a_0_x1_x2_CDPEventHandler
    
    If isFirst Then  'この位置で判定必要（CloseWindowするため通信確立後）
        IWebDriver_CloseWindow
        ChromeFirstLaunch
    End If
    
    SetDiscoverTargets
    AttachToTarget basicInfo_.TargetId
    
    Sleep 2000
End Sub

'クロームの真の初回起動時の画面遷移
Private Sub ChromeFirstLaunch()
    MsgBox "今回、初回起動のため、このファイルと同一フォルダ内に" & vbCrLf & _
                "マクロ実行専用のUserDataフォルダを作成しました。" & vbCrLf & vbCrLf & _
                "お手数ですが、再度マクロを実行してください。" & vbCrLf & _
                "※次回以降、このメッセージは表示されません。", _
                vbInformation, "Chrome初回起動メッセージ"
    End
'    Sleep 2000
'    TabKey 1
'    SendKeys "{ENTER}", True
'    Sleep 1500
'    TabKey 6
'    SendKeys "{ENTER}", True
'    Sleep 1500
'    TabKey 8
'    SendKeys "{ENTER}", True
'    Sleep 1500
'    TabKey 2
'    SendKeys "{ENTER}", True
'    Sleep 1500
End Sub
'Private Sub TabKey(num As Long)
'    Dim i As Long
'    For i = 1 To num
'        SendKeys "{TAB}", True
'    Next
'End Sub

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
End Function

'message処理送信＆受信
'受信時には、送信に対する受信ではなく、クロームからのイベント通知を受信する場合あり。
'イベント通知の場合、responsにidが含まれないため、idの有無で送信に対する受信なのかイベント通知なのかを判断
'補足:送信自体は成功しても、送信したJsonに不備がある場合も、responsにidが含まれない。
Public Function ProcessMessage( _
    sendId As Long, sendMsg As String, json As a_0_x4_JSONHandler) As String
    
    Debug.Print vbCrLf & ">>>>>>" & "送信内容：" & sendMsg
    If webSocket_.SendMessage(sendMsg) = False Then GoTo quit

    Dim response As String, responseId As Variant
    Do
        response = webSocket_.GetMessage
        If response = "受信エラー発生" Then GoTo quit
        
        responseId = json.GetValueFrom(response, Array("id"))
        If responseId = "" Then
            If InStr(response, "method") > 0 Then
                events_.GetInfo response, Me
                '制御を奪うダイアログが存在し、そのDialogOpenのイベント通知がくるとその後、ループ受信でとまる。
                '例：クリック送信⇒クリックに対する受信までの間にそのイベント通知が来ると、受信ができずに止まる。
                'そのため、DialogOpenの通知のみ特殊対応。抜けた後にはすぐダイアログ処理する必要あり。
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
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
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
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
    documentNodeId_ = json_.GetValueFrom(responseMsg, Array("result", "root", "nodeId"))
    GetDocumentNodeID = documentNodeId_
End Function

Private Function ChangeObjectIdToNodeId(objectId As String) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.RequestNode(objectId)
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
    ChangeObjectIdToNodeId = json_.GetValueFrom(responseMsg, Array("result", "nodeId"))
End Function

Private Function ChangeNodeIdToObjectId(nodeId As Long) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.ResolveNode(nodeId)
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
    ChangeNodeIdToObjectId = json_.GetValueFrom(responseMsg, Array("result", "objectId"))
End Function
'************************************************

'DOM(Element情報)系************************************************
Private Function GetAttributes(nodeId As Long) As String
    Dim sendMsg  As String
    sendMsg = CDPMethod_.GetAttributes(nodeId)
    GetAttributes = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function

Private Function DescribeNode(Id As Variant) As Variant
    Dim sendMsg  As String
    sendMsg = CDPMethod_.DescribeNode(Id)
    DescribeNode = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function
'************************************************

'Runtime************************************************
'executionContextCreated イベントによる実行コンテキスト生成の報告を有効にします。
'レポートが有効になると、イベントは既存の実行コンテキストごとに直ちに送信されます。
Public Sub RuntimeEnable()
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.RunTime_Enable()
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Sub
'JavaScriptコード実行
Private Function ExecuteJavaScript_Inner(script As String) As String
    Dim sendMsg  As String
    sendMsg = CDPMethod_.JavaScritpEvaluate(script)
    ExecuteJavaScript_Inner = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function

'JavaScript_Function実行
Private Function CallJavaScriptFunction(funcName As String, objectId As Variant, args As Variant) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.CallJavaScriptFunction(funcName, objectId, args)
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
    '↓このエラーメッセージが返ってくるのはexecutionContextIdが渡ってきた場合のみなので、ObjectIdが渡ってくるときは問題なし
    If InStr(responseMsg, "Cannot find context with specified id") > 0 Then
        ReCreateIsolatedFrameWorld
        sendMsg = CDPMethod_.CallJavaScriptFunction(funcName, executionContextId_, args)
        responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
    End If
    CallJavaScriptFunction = responseMsg
End Function
'JavaScript_Function実行(WebElementから呼ばれる用）
Public Function CallJavaScriptFunctionForElement(funcName As String, objectId As Variant, args As Variant) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.CallJavaScriptFunction(funcName, objectId, args)
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
    CallJavaScriptFunctionForElement = responseMsg
End Function
'************************************************
'Input Domain*******************************************************************
'キーイベントをページにディスパッチする。keyDown , keyUp, rawKeyDown, charのいずれか。
Public Function DispatchKeyEvent_DOMKeyCode(eventType As String, code As String) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.DispatchKeyEvent_DOMKeyCode(eventType, code)
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function

'マウスイベントをxy座標にディスパッチする。mousePressed, mouseReleased, mouseMoved, mouseWheelのいずれか。
Public Function DispatchMouseEvent(eventType As String, x As Long, y As Long, _
                                                        button As String, clickCount As Long, _
                                                        Optional modifiers As Long = 0, _
                                                        Optional deltaX As Long = 0, Optional deltaY As Long = 0) As String
                                                        
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.DispatchMouseEvent(eventType, x, y, button, clickCount, _
                                                                                    modifiers, deltaX, deltaY)
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function

'*******************************************************************

'DOM操作系************************************************
Public Sub DOMEnable()
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.DOM_Enable()
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Sub

Public Function Focus(objectId As String) As String
    Dim sendMsg  As String
    sendMsg = CDPMethod_.Focus(objectId)
    Focus = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function

Public Function GetBoxModel(objectId As String) As Variant
    Dim sendMsg  As String
    sendMsg = CDPMethod_.GetBoxModel(objectId)
    GetBoxModel = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function

Public Function SetFileInputFiles(objectId As Variant, files As String) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.SetFileInputFiles(objectId, files)
    SetFileInputFiles = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function

Public Sub FindElementByQuerySelector(Selector As String)
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.querySelector(Selector)
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
    
    Dim nodeId As Long
    nodeId = json_.GetValueFrom(responseMsg, Array("result", "nodeId"))
    Dim nodeInfoJson As String
    nodeInfoJson = DescribeNode(nodeId)
End Sub

'************************************************

'Input系************************************************
Public Function Click(objId As String) As String
    CallJavaScriptFunction "function(value){this.click();}", objId, ""
End Function
'************************************************

'Target:Page系************************************************
Public Function PageEnable() As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.PageEnable()
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function

Public Function GetFrameTree() As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.GetFrameTree()
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
    GetFrameTree = responseMsg
End Function

Public Function CreateIsolatedFrameWorld(FrameId As String) As Long
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.CreateIsolatedFrameWorld(FrameId)
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
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
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function

Public Function AttachToTarget(TargetId As String) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.AttachToTarget(TargetId)
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function

Public Function ActivateTarget(TargetId As String) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.ActivateTarget(TargetId)
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function

Public Function DetachFromTarget(TargetId As String) As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.DetachFromTarget(TargetId)
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function

Public Function TargetClose() As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.TargetClose(basicInfo_.TargetId)
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
End Function
Public Function GetTargetInfo() As String
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.GetTargets()
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
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
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
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
    
    Dim Codes As Variant: Codes = GetXpathCodes(xPath)
    Dim objectId As String: objectId = GetObjectIdOfElementBy(xPath, Codes)
    If objectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = MakeFindInfo("xpath", xPath)
    Set IWebDriver_FindElementByXpath = MakeElementObject(objectId, findInfo)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & xPath
Beep
DisplayErrorAndEnd MakeFindErrorMessage("XPath", xPath)
End Function

Public Function GetXpathCodes(xPath As String) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.evaluate('" & xPath & "', document, null, " & _
                        "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;"
    Codes(1) = "function(xpath){return  document.evaluate(xpath, document, null, " & _
                            "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;}"
    GetXpathCodes = Codes
End Function


'指定の方法で要素を探し(最大４秒間探す）、オブジェクトIDを返す
Public Function GetObjectIdOfElementBy(way As String, Codes As Variant) As String
    
    Dim startTime As Date: startTime = Timer()
    Dim responseMsg As String, objectId As String
    Do
        If executionContextId_ = 0 Then
            responseMsg = ExecuteJavaScript_Inner((Codes(0)))
        Else
            responseMsg = CallJavaScriptFunction((Codes(1)), executionContextId_, way)
        End If
        objectId = json_.GetValueFrom(responseMsg, Array("result", "result", "objectId"))
    
        If objectId <> "" Then
            GetObjectIdOfElementBy = objectId
            Exit Function
        ElseIf CInt((Timer() - startTime)) > 4 Then
            GetObjectIdOfElementBy = ""
            Exit Function
        End If
        
        DoEvents: Sleep 500
    Loop
End Function
'elementをFindした情報を配列にして返す
Public Function MakeFindInfo(way_string As String, way As String, Optional Index As Long = -1) As Collection
    Dim findInfo As Collection: Set findInfo = New Collection
    findInfo.add way_string
    findInfo.add way
    findInfo.add Index
    Set MakeFindInfo = findInfo
End Function
'エレメントを作成し、戻り値として返却
'エレメント内容：ObjectId,NodeId,Attributes(辞書),フレームの場合のみframeId追加
Public Function MakeElementObject(objectId As String, findInfo As Collection) As IWebElement
    
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
    element.Init Me, objectId, findInfo, attributes, FrameId
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
'Find系メソッドで要素が見つからなかった場合のエラーメッセージを作成する
Public Function MakeFindErrorMessage(wayString As String, way As String, Optional wait As String = "") As String
    Dim errMsg As String
    errMsg = _
    "＜FindElementBy" & wayString & "＞" & vbCrLf & _
    "指定された以下" & wayString & "の要素は存在しませんでした。" & vbCrLf & _
    way & vbCrLf & "(イミディエイトウィンドウで確認可能)" & vbCrLf & vbCrLf
    
    If wayString = "XPath" Then
        errMsg = errMsg & _
        "※引数のXPathにはChromeの検証ツールで「完全なXPath」が" & vbCrLf & _
        "　 コピペできるので、それを指定してください。" & vbCrLf & vbCrLf
    End If
    
    errMsg = errMsg & _
        "※もし取得したい要素がiframe内の要素である場合、SwitchFrameを" & vbCrLf & _
        "　 使用して、フレームを切り替えてからでないと取得できません。" & vbCrLf & vbCrLf
    
    errMsg = errMsg & _
    "※上記に沿っていても、要素が画面に出現する前だと失敗します。" & vbCrLf
    
    If wait = "" Then
        errMsg = errMsg & _
        "　 必要であればWaitForElementToAppearBy" & wayString & "で" & vbCrLf
    Else
        errMsg = errMsg & _
        "　 必要であればWaitForElementToAppearBy" & wait & "で" & vbCrLf
    End If
    
    errMsg = errMsg & _
    "　 要素の出現を確認する処理を入れましょう。" & vbCrLf
    
    MakeFindErrorMessage = errMsg
End Function

'Public
Private Function IWebDriver_FindElementById(Id As String) As IWebElement
    Dim Codes As Variant: Codes = GetIdCodes(Id)
    Dim objectId As String: objectId = GetObjectIdOfElementBy(Id, Codes)
    If objectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = MakeFindInfo("id", Id)
    Set IWebDriver_FindElementById = MakeElementObject(objectId, findInfo)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & Id
Beep
DisplayErrorAndEnd MakeFindErrorMessage("Id", Id)
End Function
Public Function GetIdCodes(Id As String) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.getElementById('" & Id & "');"
    Codes(1) = "function(id){return  document.getElementById(id);}"
    GetIdCodes = Codes
End Function

'Public
Private Function IWebDriver_FindElementByName(name As String) As IWebElement

    Dim Codes As Variant: Codes = GetNameCodes(name)
    Dim objectId As String: objectId = GetObjectIdOfElementBy(name, Codes)
    If objectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = MakeFindInfo("name", name)
    Set IWebDriver_FindElementByName = MakeElementObject(objectId, findInfo)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & name
Beep
DisplayErrorAndEnd MakeFindErrorMessage("Name", name)
End Function
Public Function GetNameCodes(name As String) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.getElementsByName('" & name & "')[0];"
    Codes(1) = "function(name){return  document.getElementsByName(name)[0];}"
    GetNameCodes = Codes
End Function

'Public
'tag名で要素コレクションを取得（要素グループのLengthを取得し、Length分要素作成ループ）
Private Function IWebDriver_FindElementsByTag(tagName As String) As IWebElements
    Dim countCodes(1) As Variant
    countCodes(0) = "document.getElementsByTagName('" & tagName & "').length;"
    countCodes(1) = "function(tagName){return  document.getElementsByTagName('" & tagName & "').length;}"
    Dim elemsCount As Long: elemsCount = GetElementsCountBy(tagName, countCodes)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    
    Dim objectId As String
    Dim Codes As Variant
    Dim Index As Long
    For Index = 0 To elemsCount - 1
        Codes = GetTagCodes(tagName, Index)
        objectId = GetObjectIdOfElementBy(tagName, Codes)
        Dim findInfo As Collection: Set findInfo = MakeFindInfo("tag", tagName, Index)
        elements.add MakeElementObject(objectId, findInfo)
    Next
    
    Dim elems As WebElements: Set elems = New WebElements
    elems.Init elements
    
    Set IWebDriver_FindElementsByTag = elems
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & tagName
Beep
DisplayErrorAndEnd MakeFindErrorMessage("Tag", tagName, "○○")
End Function
Private Function GetElementsCountBy(way As String, Codes As Variant) As Long
    Dim responseMsg As String
    If executionContextId_ = 0 Then
        responseMsg = ExecuteJavaScript_Inner((Codes(0)))
    Else
        responseMsg = CallJavaScriptFunction((Codes(1)), executionContextId_, way)
    End If
    GetElementsCountBy = json_.GetValueFrom(responseMsg, Array("result", "result", "value"))
End Function
Public Function GetTagCodes(tagName As String, Index As Long) As Variant
    Dim Codes(1) As String
    Codes(0) = "document.getElementsByTagName('" & tagName & "')[" & Index & "];"
    Codes(1) = "function(tagName){return  document.getElementsByTagName(tagName)[" & Index & "];}"
    GetTagCodes = Codes
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
    Dim json As String: json = GetFrameTree
'    collect.add json_.GetObjectFrom(json, Array("result", "frameTree", "frame"))
    
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        'まずchildFramesから一つのFrame情報（オブジェクト）取得
        frmInfo = json_.GetObjectFrom(json, _
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
Public Sub GetChildFrameInfo(collect As Collection, json As String)
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        frmInfo = json_.GetObjectFrom(json, _
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
    Dim name As String, URL As String
    For Each frameInfo In frameInfos
        name = json_.GetValueFrom(frameInfo, Array("name"))
        URL = json_.GetValueFrom(frameInfo, Array("url"))
        If name = value Or URL = value Then
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
'    responseMsg = CallJavaScriptFunction(functionCode, elem.ObjectId, "")
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
    IWebDriver_IsExistsElementByXpath = IsExistsElementBy(xPath, "xpath")
End Function
'Public
'第1引数で指定されるIdの要素有無の確認
Private Function IWebDriver_IsExistsElementById(Id As String) As Boolean
    IWebDriver_IsExistsElementById = IsExistsElementBy(Id, "id")
End Function
'Public
'第1引数で指定されるNameの要素有無の確認
Private Function IWebDriver_IsExistsElementByName(name As String) As Boolean
    IWebDriver_IsExistsElementByName = IsExistsElementBy(name, "name")
End Function
'IsExists共通関数
Private Function IsExistsElementBy(way As String, way_stringSpecify As String) As Boolean
    Dim Codes As Variant
    Select Case way_stringSpecify
      Case "xpath"
        Codes = GetXpathCodes(way)
      Case "id"
        Codes = GetIdCodes(way)
      Case "name"
        Codes = GetNameCodes(way)
    End Select

    Dim responseMsg As String
    If executionContextId_ = 0 Then
        responseMsg = ExecuteJavaScript_Inner((Codes(0)))
    Else
        responseMsg = CallJavaScriptFunction((Codes(1)), executionContextId_, way)
    End If
    Dim objectId As String
    objectId = json_.GetValueFrom(responseMsg, Array("result", "result", "objectId"))

    If objectId = "" Then
        IsExistsElementBy = False
    Else
        IsExistsElementBy = True
    End If
End Function

'Public
'Xpath(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearByXpath(xPath As String, maxWaitTime_sec As Long)
    WaitForElementToAppearBy xPath, maxWaitTime_sec, "xpath"
End Sub
'Public
'Id(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearById(Id As String, maxWaitTime_sec As Long)
    WaitForElementToAppearBy Id, maxWaitTime_sec, "id"
End Sub
'Public
'Name(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearByName(name As String, maxWaitTime_sec As Long)
    WaitForElementToAppearBy name, maxWaitTime_sec, "name"
End Sub
'WaitForElement共通関数
Private Sub WaitForElementToAppearBy(way As String, _
        maxWaitTime_sec As Long, way_stringSpecify As String)
    
    Dim startTime As Date: startTime = Timer()
    Do
        If IsExistsElementBy(way, way_stringSpecify) Then
            Exit Sub
        ElseIf CInt((Timer() - startTime)) > maxWaitTime_sec Then
            DisplayErrorAndEnd "<WaitForElementToAppearBy" & way_stringSpecify & ">" & vbCrLf & _
                                          maxWaitTime_sec & _
                                         "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
                                          way
        End If
        DoEvents: Sleep 500
    Loop
End Sub

'Public
'AlertダイアログのOKボタンを押して閉じる
Private Sub IWebDriver_HandleAlertDialog()
    Sleep 1000
    Dim sendMsg  As String, responseMsg As String
    sendMsg = CDPMethod_.HandleJavaScriptDialog()
    responseMsg = ProcessMessage(CDPMethod_.Id, sendMsg, json_)
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
Private Sub IWebDriver_SwitchTabByIndex(Index As Long)
    Dim newTargetId As String
    newTargetId = events_.PageInfos(Index)("targetId")
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

'Private Sub IWebDriver_SetForeGroundWindow()
'    SetForeGroundWindow chromeHwnd_
'End Sub



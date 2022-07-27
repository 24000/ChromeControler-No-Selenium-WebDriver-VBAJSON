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


Private webSocket_ As a1_WebSocketCommunicator
Private Json_ As a2_JSONHandler
Private basicInfo_ As a3_BasicInfos
Private execCDP_ As a4_ExecuteCDP
Private events_ As a5_CDPEventHandler
Private execHelper_ As a6_ExecuteHelperFunction

Private processId_ As Long
Private chromeHwnd_ As LongPtr

Private Sub Class_Initialize()
    Dim userDataDir As String
    userDataDir = ThisWorkbook.Path & "\マクロ実行専用UserDataフォルダ\"
    
    Dim isFirst As Boolean
    If Dir(userDataDir, vbDirectory) = "" Then
        MkDir userDataDir
        isFirst = True
    End If
    
    processId_ = StartChrome(userDataDir)
    
    '適切に取得ができないケースあり、要改善（よってまだどこでもhwndは使用しない）
    chromeHwnd_ = FindWindow(CLASS_NAME, CAPTION_NAME)
    
    Set Json_ = New a2_JSONHandler
    Set basicInfo_ = New a3_BasicInfos: basicInfo_.Init Json_
    Set webSocket_ = New a1_WebSocketCommunicator
    If webSocket_.Init(basicInfo_.Host, basicInfo_.WebSocketUrlPath, basicInfo_.DebuggingPort) = False Then End
    
    Set events_ = New a5_CDPEventHandler
    events_.Init Json_, basicInfo_
    
    Dim msgGenerator As a4x1_CDPMessageGenerator
    Set msgGenerator = New a4x1_CDPMessageGenerator
    Dim Handler As a4x2_CDPMessageHandler
    Set Handler = New a4x2_CDPMessageHandler
    Handler.Init webSocket_, Json_, events_
    Set execCDP_ = New a4_ExecuteCDP
    execCDP_.Init msgGenerator, Handler, Json_
    
    Set execHelper_ = New a6_ExecuteHelperFunction
    
    If isFirst Then  'この位置で判定必要（CloseWindowするため通信確立後）
        execCDP_.SessionClose
        ChromeFirstLaunch
    End If
    
    execCDP_.SetDiscoverTargets
    execCDP_.AttachToTarget basicInfo_.targetId
    
    execHelper_.SleepByWinAPI 2000
End Sub

Private Function StartChrome(userDataDir As String) As Long
    Const CHROME_EXE_REG_KEY As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\"
    Const CHROME_EXE As String = """chrome.exe"""
    Const DEBUG_MODE_OPTION As String = " --remote-debugging-Port=9222"
    Const USER_DATA_DIR_OPTION As String = " --user-data-dir="
    Const ENABLE_AUTOMATION_OPTION As String = " --enable-automation"
    
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

'クロームの真の初回起動時の画面遷移
Private Sub ChromeFirstLaunch()
    MsgBox "今回、初回起動のため、このファイルと同一フォルダ内に" & vbCrLf & _
                "マクロ実行専用のUserDataフォルダを作成しました。" & vbCrLf & vbCrLf & _
                "お手数ですが、再度マクロを実行してください。" & vbCrLf & _
                "※次回以降、このメッセージは表示されません。", _
                vbInformation, "Chrome初回起動メッセージ"
    End
End Sub



'以下､インターフェイス公開メソッド
'************************************************
'Public
'現在のページのURL
Private Property Get IWebDriver_URL() As String
    execCDP_.SetDiscoverTargets  'これを使ってPage情報最新化
    Dim i As Long
    For i = 1 To events_.PageInfos.Count
        If events_.PageInfos(i)("targetId") = basicInfo_.targetId Then
            IWebDriver_URL = events_.PageInfos(i)("url")
            Exit Property
        End If
    Next
End Property


'Public
'第1引数で指定のURLを開く
Private Sub IWebDriver_OpenURL(targetUrl As String)
    Dim res As String
    res = execCDP_.OpenURL(targetUrl)
    If InStr(res, "invalid URL") > 0 Then
        execHelper_.DisplayErrorAndEnd _
            "＜OpenURL＞" & vbCrLf & _
            "指定された以下URLは無効なURLです。" & vbCrLf & _
            targetUrl & vbCrLf
    End If
    
    execCDP_.WaitingForWebPageDisplay
    
    execCDP_.SetDiscoverTargets
    execCDP_.PageEnable
'    execCDP_.DOMEnable
    basicInfo_.DocumentObjectId = execCDP_.GetDocumentObjectID
    basicInfo_.DocumentNodeId = execCDP_.GetDocumentNodeID
End Sub

'Public
'第1引数のJavaScriptコードの実行
Private Function IWebDriver_ExecuteScript(script As String) As String
    Dim responseMsg As String
    responseMsg = execCDP_.ExecuteJavaScript(script)
    IWebDriver_ExecuteScript = Json_.GetValue(responseMsg, "result")
End Function

'Public
'第1引数で指定されるXpathの要素を取得し、エレメントObjectとして返却
Private Function IWebDriver_FindElementByXpath(xPath As String) As IWebElement
    
    If InStr(xPath, """") > 0 Then xPath = Replace(xPath, """", "\""")
    
    Dim codes As Variant: codes = execHelper_.GetXpathCodes(xPath)
    Dim ObjectId As String: ObjectId = execCDP_.GetObjectIdOfElementBy(xPath, basicInfo_, codes)
    If ObjectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = execHelper_.MakeFindInfo("xpath", xPath)
    Set IWebDriver_FindElementByXpath = _
        execCDP_.MakeElementObject(ObjectId, findInfo, basicInfo_, events_, execHelper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & xPath
Beep
execHelper_.DisplayErrorAndEnd execHelper_.MakeFindErrorMessage("XPath", xPath)
End Function

'Public
'第1引数で指定されるIdの要素を取得し、エレメントObjectとして返却
Private Function IWebDriver_FindElementById(Id As String) As IWebElement
    Dim codes As Variant: codes = execHelper_.GetIdCodes(Id)
    Dim ObjectId As String: ObjectId = execCDP_.GetObjectIdOfElementBy(Id, basicInfo_, codes)
    If ObjectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = execHelper_.MakeFindInfo("id", Id)
    Set IWebDriver_FindElementById = _
        execCDP_.MakeElementObject(ObjectId, findInfo, basicInfo_, events_, execHelper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & Id
Beep
execHelper_.DisplayErrorAndEnd execHelper_.MakeFindErrorMessage("Id", Id)
End Function

'Public
'第1引数で指定されるNameの要素を取得し、エレメントObjectとして返却
Private Function IWebDriver_FindElementByName(name As String) As IWebElement

    Dim codes As Variant: codes = execHelper_.GetNameCodes(name)
    Dim ObjectId As String: ObjectId = execCDP_.GetObjectIdOfElementBy(name, basicInfo_, codes)
    If ObjectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = execHelper_.MakeFindInfo("name", name)
    Set IWebDriver_FindElementByName = _
        execCDP_.MakeElementObject(ObjectId, findInfo, basicInfo_, events_, execHelper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & name
Beep
execHelper_.DisplayErrorAndEnd execHelper_.MakeFindErrorMessage("Name", name)
End Function


'Public
'tag名で要素コレクションを取得（要素グループのLengthを取得し、Length分要素作成ループ）
Private Function IWebDriver_FindElementsByTag(tagName As String) As IWebElements
    Dim countCodes As Variant: countCodes = execHelper_.GetTagCountCodes(tagName)
    Dim elemsCount As Long: elemsCount = execCDP_.GetElementsCountBy(tagName, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    
    Dim ObjectId As String
    Dim codes As Variant
    Dim Index As Long
    For Index = 0 To elemsCount - 1
        codes = execHelper_.GetTagCodes(tagName, Index)
        ObjectId = execCDP_.GetObjectIdOfElementBy(tagName, basicInfo_, codes)
        Dim findInfo As Collection: Set findInfo = execHelper_.MakeFindInfo("tag", tagName, Index)
        elements.add _
            execCDP_.MakeElementObject(ObjectId, findInfo, basicInfo_, events_, execHelper_)
    Next
    
    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements
    
    Set IWebDriver_FindElementsByTag = elems
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & tagName
Beep
execHelper_.DisplayErrorAndEnd execHelper_.MakeFindErrorMessage("Tag", tagName, "○○")
End Function



'Public
'第1引数で指定される値のフレームにフレームを切替
Private Sub IWebDriver_SwitchFrameByNameOrUrlOrIndex(value As Variant)
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim FrameId As String
    If IsNumeric(value) Then
        If frameInfos.Count < value Then GoTo quit
        FrameId = Json_.GetValue(frameInfos(value), "id")
    Else
        FrameId = GetTargetFrameId(frameInfos, value)
    End If
    If FrameId = "" Then GoTo quit
    
    Dim ExecutionContextId As Long
    ExecutionContextId = execCDP_.CreateIsolatedFrameWorld(FrameId)
    basicInfo_.ExecutionContextId = ExecutionContextId
    basicInfo_.IsolatedFrameId = FrameId
    
Exit Sub

quit:
execHelper_.DisplayErrorAndEnd "＜SwitchFrameByNameOrUrlOrIndex＞" & vbCrLf & _
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
    Dim Json As String: Json = execCDP_.GetFrameTree
'    collect.add json_.GetObject(json, "result", "frameTree", "frame"))
    
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        'まずchildFramesから一つのFrame情報（オブジェクト）取得
        frmInfo = Json_.GetObject(Json, "result", "frameTree", "childFrames", Count)
        If frmInfo = "" Then Exit Do
        'そのオブジェクトの中のframeプロパティ(オブジェクト)を取得し、コレクションに加える
        frminfo2 = Json_.GetObject(frmInfo, "frame")
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
Private Sub GetChildFrameInfo(collect As Collection, Json As String)
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        frmInfo = Json_.GetObject(Json, "childFrames", Count)
        If frmInfo = "" Then Exit Do
        frminfo2 = Json_.GetObject(frmInfo, "frame")
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
        name = Json_.GetValue(frameInfo, "name")
        URL = Json_.GetValue(frameInfo, "url")
        If name = value Or URL = value Then
            GetTargetFrameId = Json_.GetValue(frameInfo, "id")
            Exit Function
        End If
    Next
End Function

'Public
'切り替え可能なフレームの情報をイミディエイトウィンドウに表示
Private Sub IWebDriver_DebugPrintAccessibleFrameInfo()
    execCDP_.PageEnable
    execCDP_.RuntimeEnable
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim i As Long
    For i = 1 To frameInfos.Count
        Debug.Print "index：" & i & _
                           "　name：" & Json_.GetValue(frameInfos(i), "name") & _
                           "　url：" & Json_.GetValue(frameInfos(i), "url")
    Next
End Sub

'Public
Private Sub IWebDriver_SwitchFrameByIframeElement(IframeElement As IWebElement)
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim elem As b0x0_WebElement: Set elem = IframeElement
    If elem.FrameId = "" Then
        execHelper_.DisplayErrorAndEnd "＜SwitchFrameByIframeElement＞" & vbCrLf & _
                                      "指定されたElementがIframeではありません" & vbCrLf & vbCrLf & _
                                      "Iframe要素を引数に渡して下さい" & vbCrLf
    End If
    
    execCDP_.GetTargetInfo
    execCDP_.SetDiscoverTargets

    'frameElementのframeIdが正当なframeIdならIsolatedWorldで対応
    Dim ExecutionContextId As Long
    ExecutionContextId = execCDP_.CreateIsolatedFrameWorld(elem.FrameId)
    If ExecutionContextId <> 0 Then
        basicInfo_.ExecutionContextId = ExecutionContextId
        basicInfo_.IsolatedFrameId = elem.FrameId
        Exit Sub
    Else
        execHelper_.DisplayErrorAndEnd "＜SwitchFrameByIframeElement＞" & vbCrLf & _
                                       "このIframeはこのメソッドでは切り替えができません。" & vbCrLf & vbCrLf & _
                                      "※SwitchFrameByNameOrUrlOrIndexの" & vbCrLf & _
                                      "　 使用を検討してください。" & vbCrLf
    End If
    
    
'    テスト
'    Dim docElem As b0x0_WebElement
'    Dim functionCode As String, responseMsg As String, docObjectId As String
'    functionCode = "function(name){return  this.contentDocument;}"
'    responseMsg = CallJavaScriptFunction(functionCode, elem.ObjectId, "")
'    docObjectId = json_.GetValue(responseMsg, "result", "result", "objectId"))
'    Dim info  As String
'    info = DescribeNode(docObjectId)
    
    
'    frameIdがtargetIdなら接続切り替えで対応
'    basicInfo_.Update elem.FrameId, json_, False
'    isolatedFrameId_ = ""
'
'    webSocket_.CloseWebSocket
'    Set webSocket_ = Nothing
'    Set webSocket_ = New a1_WebSocketCommunicator
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
    If basicInfo_.IsolatedFrameId <> "" Then
        basicInfo_.IsolatedFrameId = ""
        basicInfo_.ExecutionContextId = 0
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
'    Set webSocket_ = New a1_WebSocketCommunicator
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
    IWebDriver_IsExistsElementByXpath = _
        execCDP_.IsExistsElementBy(xPath, basicInfo_, execHelper_.GetXpathCodes(xPath))
End Function
'Public
'第1引数で指定されるIdの要素有無の確認
Private Function IWebDriver_IsExistsElementById(Id As String) As Boolean
    IWebDriver_IsExistsElementById = _
        execCDP_.IsExistsElementBy(Id, basicInfo_, execHelper_.GetIdCodes(Id))
End Function
'Public
'第1引数で指定されるNameの要素有無の確認
Private Function IWebDriver_IsExistsElementByName(name As String) As Boolean
    IWebDriver_IsExistsElementByName = _
        execCDP_.IsExistsElementBy(name, basicInfo_, execHelper_.GetNameCodes(name))
End Function


'Public
'Xpath(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearByXpath(xPath As String, maxWaitTime_sec As Long)
    If execCDP_.WaitForElementToAppearBy(xPath, _
        maxWaitTime_sec, basicInfo_, execHelper_.GetXpathCodes(xPath)) = False Then
        
        execHelper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearByXPath>" & vbCrLf & _
             maxWaitTime_sec & "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
             xPath
    End If
End Sub
'Public
'Id(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearById(Id As String, maxWaitTime_sec As Long)
    If execCDP_.WaitForElementToAppearBy(Id, _
        maxWaitTime_sec, basicInfo_, execHelper_.GetIdCodes(Id)) = False Then
        
        execHelper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearById>" & vbCrLf & _
             maxWaitTime_sec & "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
             Id
    End If
End Sub
'Public
'Name(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearByName(name As String, maxWaitTime_sec As Long)
    If execCDP_.WaitForElementToAppearBy(name, _
        maxWaitTime_sec, basicInfo_, execHelper_.GetNameCodes(name)) = False Then
    
        execHelper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearByName>" & vbCrLf & _
             maxWaitTime_sec & "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
             name
    End If
End Sub


'Public
'AlertダイアログのOKボタンを押して閉じる
Private Sub IWebDriver_HandleAlertDialog()
    Dim res As String
    Dim fileSafeCounter As Long
    
    Do
        execHelper_.SleepByWinAPI 500
        res = execCDP_.HandleAlertDialog
        If InStr(res, "No dialog is showing") = 0 Then
            execHelper_.SleepByWinAPI 500
            Exit Sub
        End If
        
        fileSafeCounter = fileSafeCounter + 1
    Loop While fileSafeCounter < 6

quit:
    execHelper_.DisplayErrorAndEnd "<HandleAlertDialog>" & vbCrLf & _
                                   "ダイアログが表示されていません。" & vbCrLf & _
                                   "※このメソッドはAlertが表示される場合に使用します。" & vbCrLf
End Sub

Private Sub IWebDriver_SleepByWinAPI(milliSecond As Long)
    execHelper_.SleepByWinAPI milliSecond
End Sub


'タブの移動（Index指定：一番左のタブを１とする）
Private Sub IWebDriver_SwitchTabByIndex(Index As Long)
    Dim newTargetId As String
    newTargetId = events_.PageInfos(Index)("targetId")
    basicInfo_.Update newTargetId, Json_, True
    
    webSocket_.CloseWebSocket
    Set webSocket_ = Nothing
    Set webSocket_ = New a1_WebSocketCommunicator
    If webSocket_.Init(basicInfo_.Host, basicInfo_.WebSocketUrlPath, basicInfo_.DebuggingPort) = False Then End
    
    execCDP_.ActivateTarget basicInfo_.targetId
    execCDP_.AttachToTarget basicInfo_.targetId
    execCDP_.PageEnable
    execCDP_.DOMEnable
    
    basicInfo_.DocumentObjectId = execCDP_.GetDocumentObjectID
    basicInfo_.DocumentNodeId = execCDP_.GetDocumentNodeID
End Sub

Private Sub IWebDriver_CloseWindow()
    execCDP_.SessionClose
End Sub

Private Sub IWebDriver_ScrollWindow(yPixel As Long, Optional xPixel As Long = 0&)
    execCDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, xPixel, yPixel
End Sub

'Private Sub IWebDriver_SetForeGroundWindow()
'    SetForeGroundWindow chromeHwnd_
'End Sub

Private Sub IWebDriver_AllowPopUpWindowForURL(targetUrl As String)
    Dim chromeExe As String
    Dim objExec As Object
    With CreateObject("WScript.Shell")
        .RegWrite name:="HKEY_LOCAL_MACHINE\SOFTWARE\Policies\{Google}", _
                        value:=targetUrl, _
                        Type:="REG_SZ"
    End With
End Sub

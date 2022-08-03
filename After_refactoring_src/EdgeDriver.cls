VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EdgeDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWebDriver

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Const CLASS_NAME As String = "Chrome_WidgetWin_1"
'Private Const CAPTION_NAME As String = "新しいタブ - Google Chrome" '"data:, - Google Chrome"

'Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
'        ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
'

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "USER32.DLL" ( _
        ByVal hWnd As LongPtr, ByRef ProcessId As Long) As Long

Private Declare PtrSafe Sub SetForeGroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal ms As LongPtr)


Private webSocket_ As a1_WebSocketCommunicator
Private Json_ As a2_JSONHandler
Private basicInfo_ As a3_BasicInfos
Private CDP_ As a4_ExecuteCDP
Private events_ As a5_CDPEventHandler
Private Helper_ As a6_ExecuteHelperFunction

Private processId_ As Long
'Private chromeHwnd_ As LongPtr

Private Sub Class_Initialize()
    Dim userDataDir As String
    userDataDir = ThisWorkbook.Path & "\マクロ実行専用UserDataフォルダForEdge\"
    
    Dim isFirst As Boolean
    If Dir(userDataDir, vbDirectory) = "" Then
        MkDir userDataDir
        isFirst = True
    End If
    
    processId_ = StartEdge(userDataDir)
    
    '適切に取得ができないケースあり、要改善（よってまだどこでもhwndは使用しない）
'    chromeHwnd_ = FindWindow(CLASS_NAME, CAPTION_NAME)
    
    New_InnerObjects
    
    If isFirst Then  'この位置で判定必要（CloseWindowするため通信確立後）
        CDP_.SessionClose
        EdgeFirstLaunch
    End If
    
    CDP_.SetDiscoverTargets
    CDP_.AttachToTarget basicInfo_.targetId
    
    Helper_.SleepByWinAPI 2000
End Sub

Private Function StartEdge(userDataDir As String) As Long
    Const EDGE_EXE_REG_KEY As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe\"
    Const EDGE_EXE As String = """msedge.exe"""
    Const DEBUG_MODE_OPTION As String = " --remote-debugging-Port=9222"
    Const USER_DATA_DIR_OPTION As String = " --user-data-dir="
    Const ENABLE_AUTOMATION_OPTION As String = " --enable-automation"
    
    Dim edgeExe As String
    Dim objExec As Object
    With CreateObject("WScript.Shell")
        edgeExe = """" & .RegRead(EDGE_EXE_REG_KEY) & """"
        Set objExec = .Exec( _
                            edgeExe & DEBUG_MODE_OPTION & ENABLE_AUTOMATION_OPTION _
                            & USER_DATA_DIR_OPTION & "" & userDataDir & "")
        StartEdge = objExec.ProcessId
        Set objExec = Nothing
    End With
End Function

'内部オブジェクトの生成
Private Sub New_InnerObjects()
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
    Set CDP_ = New a4_ExecuteCDP
    CDP_.Init msgGenerator, Handler, Json_
    
    Set Helper_ = New a6_ExecuteHelperFunction
End Sub

'クロームの真の初回起動時の画面遷移
Private Sub EdgeFirstLaunch()
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
    CDP_.SetDiscoverTargets  'これを使ってPage情報最新化
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
    res = CDP_.OpenURL(targetUrl)
    If InStr(res, "invalid URL") > 0 Then
        Helper_.DisplayErrorAndEnd _
            "＜OpenURL＞" & vbCrLf & _
            "指定された以下URLは無効なURLです。" & vbCrLf & _
            targetUrl & vbCrLf
    End If
    
    CDP_.WaitingForWebPageDisplay
    
    CDP_.SetDiscoverTargets
    CDP_.PageEnable
'    CDP_.DOMEnable
    basicInfo_.DocumentObjectId = CDP_.GetDocumentObjectID
    basicInfo_.DocumentNodeId = CDP_.GetDocumentNodeID
End Sub

'Public
'第1引数のJavaScriptコードの実行
Private Function IWebDriver_ExecuteScript(script As String) As String
    Dim responseMsg As String
    responseMsg = CDP_.ExecuteJavaScript(script)
    IWebDriver_ExecuteScript = Json_.GetValue(responseMsg, "result")
End Function

'Public
'第1引数で指定されるXpathの要素を取得し、エレメントObjectとして返却
Private Function IWebDriver_FindElementByXpath(xPath As String) As IWebElement
    
    If InStr(xPath, """") > 0 Then xPath = Replace(xPath, """", "\""")
    
    Dim codes As Variant: codes = Helper_.GetXpathCodes(xPath)
    Dim ObjectId As String: ObjectId = CDP_.GetObjectIdOfElementBy(xPath, basicInfo_, codes)
    If ObjectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("xpath", xPath)
    Set IWebDriver_FindElementByXpath = _
        CDP_.MakeElementObject(ObjectId, findInfo, basicInfo_, events_, Helper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & xPath
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("XPath", xPath)
End Function

'Public
'第1引数で指定されるIdの要素を取得し、エレメントObjectとして返却
Private Function IWebDriver_FindElementById(Id As String) As IWebElement
    Dim codes As Variant: codes = Helper_.GetIdCodes(Id)
    Dim ObjectId As String: ObjectId = CDP_.GetObjectIdOfElementBy(Id, basicInfo_, codes)
    If ObjectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("id", Id)
    Set IWebDriver_FindElementById = _
        CDP_.MakeElementObject(ObjectId, findInfo, basicInfo_, events_, Helper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & Id
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Id", Id)
End Function

'Public
'第1引数で指定されるNameの要素を取得し、エレメントObjectとして返却
Private Function IWebDriver_FindElementByName(name As String) As IWebElement

    Dim codes As Variant: codes = Helper_.GetNameCodes(name)
    Dim ObjectId As String: ObjectId = CDP_.GetObjectIdOfElementBy(name, basicInfo_, codes)
    If ObjectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("name", name)
    Set IWebDriver_FindElementByName = _
        CDP_.MakeElementObject(ObjectId, findInfo, basicInfo_, events_, Helper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & name
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Name", name)
End Function


'Public
'tag名で要素コレクションを取得（要素グループのLengthを取得し、Length分要素作成ループ）
Private Function IWebDriver_FindElementsByTag(tagName As String) As IWebElements
    Dim countCodes As Variant: countCodes = Helper_.GetTagCountCodes(tagName)
    Dim elemsCount As Long: elemsCount = CDP_.GetElementsCountBy(tagName, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    
    Dim ObjectId As String
    Dim codes As Variant
    Dim Index As Long
    For Index = 0 To elemsCount - 1
        codes = Helper_.GetTagCodes(tagName, Index)
        ObjectId = CDP_.GetObjectIdOfElementBy(tagName, basicInfo_, codes)
        Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("tag", tagName, Index)
        elements.add _
            CDP_.MakeElementObject(ObjectId, findInfo, basicInfo_, events_, Helper_)
    Next
    
    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements
    
    Set IWebDriver_FindElementsByTag = elems
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & tagName
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Tag", tagName, "○○")
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
    ExecutionContextId = CDP_.CreateIsolatedFrameWorld(FrameId)
    basicInfo_.ExecutionContextId = ExecutionContextId
    basicInfo_.IsolatedFrameId = FrameId
    
Exit Sub

quit:
Helper_.DisplayErrorAndEnd "＜SwitchFrameByNameOrUrlOrIndex＞" & vbCrLf & _
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
    Dim Json As String: Json = CDP_.GetFrameTree
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
    CDP_.PageEnable
    CDP_.RuntimeEnable
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
        Helper_.DisplayErrorAndEnd "＜SwitchFrameByIframeElement＞" & vbCrLf & _
                                      "指定されたElementがIframeではありません" & vbCrLf & vbCrLf & _
                                      "Iframe要素を引数に渡して下さい" & vbCrLf
    End If
    
    CDP_.GetTargetInfo
    CDP_.SetDiscoverTargets

    'frameElementのframeIdが正当なframeIdならIsolatedWorldで対応
    Dim ExecutionContextId As Long
    ExecutionContextId = CDP_.CreateIsolatedFrameWorld(elem.FrameId)
    If ExecutionContextId <> 0 Then
        basicInfo_.ExecutionContextId = ExecutionContextId
        basicInfo_.IsolatedFrameId = elem.FrameId
        Exit Sub
    Else
        Helper_.DisplayErrorAndEnd "＜SwitchFrameByIframeElement＞" & vbCrLf & _
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
        CDP_.IsExistsElementBy(xPath, basicInfo_, Helper_.GetXpathCodes(xPath))
End Function
'Public
'第1引数で指定されるIdの要素有無の確認
Private Function IWebDriver_IsExistsElementById(Id As String) As Boolean
    IWebDriver_IsExistsElementById = _
        CDP_.IsExistsElementBy(Id, basicInfo_, Helper_.GetIdCodes(Id))
End Function
'Public
'第1引数で指定されるNameの要素有無の確認
Private Function IWebDriver_IsExistsElementByName(name As String) As Boolean
    IWebDriver_IsExistsElementByName = _
        CDP_.IsExistsElementBy(name, basicInfo_, Helper_.GetNameCodes(name))
End Function


'Public
'Xpath(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearByXpath(xPath As String, maxWaitTime_sec As Long)
    If CDP_.WaitForElementToAppearBy(xPath, _
        maxWaitTime_sec, basicInfo_, Helper_.GetXpathCodes(xPath)) = False Then
        
        Helper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearByXPath>" & vbCrLf & _
             maxWaitTime_sec & "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
             xPath
    End If
End Sub
'Public
'Id(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearById(Id As String, maxWaitTime_sec As Long)
    If CDP_.WaitForElementToAppearBy(Id, _
        maxWaitTime_sec, basicInfo_, Helper_.GetIdCodes(Id)) = False Then
        
        Helper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearById>" & vbCrLf & _
             maxWaitTime_sec & "秒経過しても、指定の以下要素は現れませんでした。" & vbCrLf & _
             Id
    End If
End Sub
'Public
'Name(第1引数)で指定されるの要素が出現するのを、最大待機時間(第2引数)待つ
Private Sub IWebDriver_WaitForElementToAppearByName(name As String, maxWaitTime_sec As Long)
    If CDP_.WaitForElementToAppearBy(name, _
        maxWaitTime_sec, basicInfo_, Helper_.GetNameCodes(name)) = False Then
    
        Helper_.DisplayErrorAndEnd _
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
        Helper_.SleepByWinAPI 500
        res = CDP_.HandleAlertDialog
        If InStr(res, "No dialog is showing") = 0 Then
            Helper_.SleepByWinAPI 500
            Exit Sub
        End If
        
        fileSafeCounter = fileSafeCounter + 1
    Loop While fileSafeCounter < 6

quit:
    Helper_.DisplayErrorAndEnd "<HandleAlertDialog>" & vbCrLf & _
                                   "ダイアログが表示されていません。" & vbCrLf & _
                                   "※このメソッドはAlertが表示される場合に使用します。" & vbCrLf
End Sub

Private Sub IWebDriver_SleepByWinAPI(milliSecond As Long)
    Helper_.SleepByWinAPI milliSecond
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
    
    CDP_.ActivateTarget basicInfo_.targetId
    CDP_.AttachToTarget basicInfo_.targetId
    CDP_.PageEnable
    CDP_.DOMEnable
    
    basicInfo_.DocumentObjectId = CDP_.GetDocumentObjectID
    basicInfo_.DocumentNodeId = CDP_.GetDocumentNodeID
End Sub

Private Sub IWebDriver_CloseWindow()
    CDP_.SessionClose
End Sub

Private Sub IWebDriver_ScrollWindow(yPixel As Long, Optional xPixel As Long = 0&)
    CDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, xPixel, yPixel
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



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IWebElement

Private Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        
Private Declare PtrSafe Function FindWindowEx Lib "User32" Alias "FindWindowExA" ( _
        ByVal hWndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        
Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Any) As LongPtr

Private chrome_ As ChromeDriver

Private objectId_ As String
Private nodeId_ As Long
Private attributes_ As Object
Private frameId_ As String


Public Sub Init(chrome As ChromeDriver, _
                        objectId As String, Optional nodeId As Long = 0, _
                        Optional attributes As Object = Nothing, Optional FrameId As String)
    Set chrome_ = chrome
    objectId_ = objectId
    nodeId_ = nodeId
    Set attributes_ = attributes
    frameId_ = FrameId
End Sub

Public Property Get objectId() As String
    objectId = objectId_
End Property
Public Property Get FrameId() As String
    FrameId = frameId_
End Property

'***********************************************
'Public
'チェックボックス要素についてチェックChecked属性の設定
Private Property Let IWebElement_Checked(RHS As Boolean)
    Dim responseJson As String
    'checkedプロパティを持っているか事前にチェック。持っていない場合はじく。
    responseJson = chrome_.JavaScritpCallFunction( _
        "function(value){return this.checked;}", objectId_, "")
    If InStr(responseJson, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim boolString As String
    If RHS Then
        boolString = "true"
    Else
        boolString = "false"
    End If

    responseJson = chrome_.JavaScritpCallFunction( _
        "function(value){ this.checked = value;}", objectId_, boolString)

Exit Property
quit:
chrome_.DisplayErrorAndEnd "このElementはCheckedを持ちません。" & vbCrLf & _
                                            "※CheckedはCheckBoxで使用可能"
End Property
'Public
'チェックボックス要素についてチェック状態取得
Private Property Get IWebElement_Checked() As Boolean
    Dim responseJson As String
    responseJson = chrome_.JavaScritpCallFunction( _
        "function(value){return this.checked;}", objectId_, "")
    If InStr(responseJson, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim Checked As Boolean
    Checked = chrome_.Json.GetValueFrom(responseJson, Array("result", "result", "value"))
    IWebElement_Checked = Checked
Exit Property

quit:
chrome_.DisplayErrorAndEnd "このElementはCheckedを持ちません。" & vbCrLf & _
                                            "※CheckedはCheckBoxで使用可能"
End Property


'Public
'要素の状態(有効or無効)を取得
Private Property Get IWebElement_IsEnabled() As Boolean
    Dim responseJson As String
    responseJson = chrome_.JavaScritpCallFunction( _
        "function(value){return this.disabled;}", objectId_, "")
    If InStr(responseJson, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim Enabled As Boolean
    Enabled = chrome_.Json.GetValueFrom(responseJson, Array("result", "result", "value"))
    IWebElement_IsEnabled = Not Enabled
Exit Property
    
quit:
chrome_.DisplayErrorAndEnd "このElementはEnabledを持ちません。" & vbCrLf & _
                                            "※EnabledはCheckBoxで使用可能"
End Property



'Public
'javaScriptによってクリック実行
Private Sub IWebElement_Click()
    chrome_.Focus (objectId_)
    Dim responseJson As String
    responseJson = chrome_.JavaScritpCallFunction("function(value){this.click();}", objectId_, "")
    
    If InStr(responseJson, "this.click is not a function") Then
        Dim javaScriptCode As String
        javaScriptCode = "function(value){this.dispatchEvent(new MouseEvent( 'click' ));}"
        responseJson = chrome_.JavaScritpCallFunction(javaScriptCode, objectId_, "")
    End If
    
    If InStr(responseJson, "this.dispatchEvent is not a function") Then
        chrome_.Focus (objectId_)
        chrome_.DispatchKeyEvent_DOMKeyCode "keyUp", "Enter"
    End If
    
    If chrome_.CDPEvents.DialogInfoDic("IsExistDialog") Then Exit Sub
    
    If IsPageChanging Then WaitPageChangeToComplete
    If IsIsolatedFrameChanging Then WaitIsolatedFrameChangeToComplete
    
End Sub
Private Function IsPageChanging() As Boolean
    chrome_.PageEnable '一回無害なメソッドを投げて、イベント発生状況更新
    
    If chrome_.CDPEvents.IsCurrentTargetScheduledNavigation _
    Or chrome_.CDPEvents.IsCurrentTargetStartedLoading Then
        IsPageChanging = True
    Else
        IsPageChanging = False
    End If
End Function
Private Sub WaitPageChangeToComplete()
    Dim ChromeInterFace As IWebDriver
    Set ChromeInterFace = chrome_
    
    'TargetNavigatedは返って来ないケースがある？ので繰り返し回数に制限をかけてループ
    Dim failSafeCounter As Long
    Dim isEndChange As Boolean
    Do
        isEndChange = chrome_.CDPEvents.IsCurrentTargetNavigated
'                                Or　chrome_.CDPEvents.IsCurrentTargetClearedScheduledNavigation
        If isEndChange Then Exit Do
        
        chrome_.RuntimeEnable
        ChromeInterFace.SleepByWindowsAPI 1000
        failSafeCounter = failSafeCounter + 1
    Loop While (isEndChange = False And failSafeCounter <= 5)
    
    chrome_.CDPEvents.IsCurrentTargetScheduledNavigation = False
    chrome_.CDPEvents.IsCurrentTargetStartedLoading = False
    chrome_.CDPEvents.IsCurrentTargetClearedScheduledNavigation = False
    chrome_.CDPEvents.IsCurrentTargetNavigated = False
End Sub
Private Function IsIsolatedFrameChanging() As Boolean
    chrome_.DOMEnable '一回無害なメソッド(frameなのでDOM系)を投げて、イベント発生状況更新
    
    If chrome_.CDPEvents.IsIsolatedFrameScheduledNavigation _
    Or chrome_.CDPEvents.IsIsolatedFrameStartedLoading Then
        IsIsolatedFrameChanging = True
    Else
        IsIsolatedFrameChanging = False
    End If
End Function
Private Sub WaitIsolatedFrameChangeToComplete()
    Dim ChromeInterFace As IWebDriver
    Set ChromeInterFace = chrome_
    
    'FrameNavigatedは返って来ないケースがある？ので繰り返し回数に制限をかけてループ
    Dim failSafeCounter As Long
    Dim isEndChange As Boolean
    Do
        isEndChange = chrome_.CDPEvents.IsIsolatedFrameNavigated
'                               Or chrome_.CDPEvents.IsIsolatedFrameClearedScheduledNavigation
        If isEndChange Then Exit Do
        
        chrome_.DOMEnable
        chrome_.RuntimeEnable
        ChromeInterFace.SleepByWindowsAPI 1000
        failSafeCounter = failSafeCounter + 1
    Loop While (isEndChange = False And failSafeCounter <= 5)
    
    chrome_.CDPEvents.IsIsolatedFrameScheduledNavigation = False
    chrome_.CDPEvents.IsIsolatedFrameStartedLoading = False
    chrome_.CDPEvents.IsIsolatedFrameClearedScheduledNavigation = False
    chrome_.CDPEvents.IsIsolatedFrameNavigated = False
End Sub





'Public
'要素のValue属性に文字列を設定
Private Sub IWebElement_SetText(text As String)
    chrome_.Focus (objectId_)
    chrome_.DispatchKeyEvent_DOMKeyCode "keyUp", "Enter"

    Dim encodingText As String: encodingText = chrome_.Json.EncodeURL(text)
    chrome_.JavaScritpCallFunction _
    "function(value){" & _
        "value = decodeURI(value);" & _
        "this.value = value;" & _
    "}", _
    objectId_, encodingText
End Sub

'Public
'要素のTextContentに文字列を設定
Private Sub IWebElement_SetTextContent(text As String)
    Dim encodingText As String: encodingText = chrome_.Json.EncodeURL(text)
    chrome_.JavaScritpCallFunction _
    "function(value){" & _
        "value = decodeURI(value);" & _
        "this.textContent  = value;" & _
    "}", _
    objectId_, encodingText
End Sub

'Public
'SelectBoxの選択肢を文字列で選択
Private Sub IWebElement_SelectItemInSelectBoxByText(text As String)
    Dim encodingText As String: encodingText = chrome_.Json.EncodeURL(text)
    Dim javaScriptCode As String
    javaScriptCode = _
    "function(value){value = decodeURI(value);" & _
        "for(let i = 0; i < this.options.length; i++){" & _
            "let val1 = this.options[i].value.replace(/\\r?\\n/g,'');" & _
            "let val2 = this.options[i].textContent.replace(/\\r?\\n/g,'');" & _
            "console.log(value + ':' + val1 + ':'+ val2);" & _
            "if(val1 === value || val2 === value){" & _
                "this.options[i].selected = true;" & _
                "return true;" & _
            "}" & _
        "}" & _
    "return false;}"
    
    Dim responseJson As String
    responseJson = chrome_.JavaScritpCallFunction(javaScriptCode, objectId_, encodingText)
    
    Dim IsSelect As Boolean
    IsSelect = chrome_.Json.GetValueFrom(responseJson, Array("result", "result", "value"))
    
    If IsSelect = False Then
        MsgBox "「" & text & "」は選択肢の中に存在しません。" & vbCrLf & _
                    "処理を終了します。"
        chrome_.WebSocket.Http.CloseHInternetHandles
        End
    End If
    
    javaScriptCode = _
    "function(value){this.onchange();}"
    responseJson = chrome_.JavaScritpCallFunction(javaScriptCode, objectId_, encodingText)
    
    If InStr(responseJson, "this.onchange is not a function") Then
        javaScriptCode = _
        "function(value){" & _
            "const event = new CustomEvent('change', {bubbles: false, cancelable: false});" & _
            "this.dispatchEvent(event);" & _
        "}"
        responseJson = chrome_.JavaScritpCallFunction(javaScriptCode, objectId_, encodingText)
    End If

End Sub



'Public
Private Function IWebElement_FindElementByName(Name As String) As IWebElement
    Dim FunctionCode As String, responseJson As String
    FunctionCode = "function(name){return  this.getElementsByName(name)[0];}"
    responseJson = chrome_.JavaScritpCallFunction(FunctionCode, objectId_, Name)
    
    Dim findObjectId As String
    findObjectId = chrome_.Json.GetValueFrom( _
                            responseJson, Array("result", "result", "objectId"))
    If findObjectId = "" Then GoTo quit
    Set IWebElement_FindElementByName = chrome_.MakeElementObject(findObjectId)
Exit Function

quit:
chrome_.DisplayErrorAndEnd "指定された以下Nameの要素は存在しませんでした。" & vbCrLf & Name
End Function

'Public
Private Function IWebElement_FindElementsByTag(tagName As String) As IWebElements
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
    
    Set IWebElement_FindElementsByTag = elems
Exit Function
quit:
chrome_.DisplayErrorAndEnd "指定された以下TagNameの要素は存在しませんでした。" & vbCrLf & tagName
End Function
Private Function GetElementsCount(tagName As String) As Long
    Dim javaScriptCode As String, responseJson As String
    javaScriptCode = _
    "function(tagName){return  this.getElementsByTagName('" & tagName & "').length;}"
    responseJson = chrome_.JavaScritpCallFunction(javaScriptCode, objectId_, tagName)
    GetElementsCount = chrome_.Json.GetValueFrom(responseJson, Array("result", "result", "value"))
End Function
Private Function CreateElementsItemByIndex(tagName As String, index As Long) As IWebElement
    Dim javaScriptCode As String, responseJson As String
    Dim objectId As String
    Dim elem As IWebElement

    javaScriptCode = _
    "function(tagName){return  this.getElementsByTagName(tagName)[" & index & "];}"
    responseJson = chrome_.JavaScritpCallFunction(javaScriptCode, objectId_, tagName)

    objectId = chrome_.Json.GetValueFrom(responseJson, Array("result", "result", "objectId"))
    Set CreateElementsItemByIndex = chrome_.MakeElementObject(objectId)
End Function

'Public
Private Function IWebElement_GetTextContent() As String
    Dim javaScriptCode As String, responseJson As String
    javaScriptCode = "function(value){return this.textContent.replace(/\\s/g, ' ');}"
    responseJson = chrome_.JavaScritpCallFunction(javaScriptCode, objectId_, "")
    
    Dim text As String
    text = chrome_.Json.GetValueFrom(responseJson, Array("result", "result", "value"))
    IWebElement_GetTextContent = text
End Function


'要素中心位置のxy座標をCDPMethodで左クリック発行
Private Sub IWebElement_DispatchLeftClickEvent()
    Dim centerXY As Variant
    centerXY = chrome_.GetBoxModel(objectId_)
    Dim x As Long, y As Long
    x = centerXY(0)
    y = centerXY(1)
    
    Dim ChromeInterFace As IWebDriver: Set ChromeInterFace = chrome_
    ChromeInterFace.ScrollWindow y - 50
    
    Dim result As String
    result = chrome_.DispatchMouseEvent("mousePressed", x, y, "left", 1)
    result = chrome_.DispatchMouseEvent("mouseReleased", x, y, "left", 1)
End Sub

'Public
'type=fileのInputBoxへファイル設定
Private Sub IWebElement_SetFileToFileSelectInputButton(filePath As String)
'    filePath = Replace(filePath, "\", "\\")
'    filePath = chrome_.Json.EncodeURL(filePath)
'    filePath = Replace(filePath, "%5C", "\")
'    chrome_.SetFileInputFiles objectId_,
'CDPMethodで設定しようとしても、ファイルパスに日本語が含まれる場合、
'エンコードして送信はできるが、クローム側でデコードしてくれない為、没
    Dim centerXY As Variant
    centerXY = chrome_.GetBoxModel(objectId_)
    
    Dim ChromeInterFace As IWebDriver: Set ChromeInterFace = chrome_
    ChromeInterFace.ScrollWindow CLng(centerXY(1)) - 50
    
    IWebElement_DispatchLeftClickEvent
    chrome_.SleepByWindowsAPIInner 1000
    ControlFileSelectDialog filePath
End Sub
Private Sub ControlFileSelectDialog(filePath As String)
    Dim DialogHwnd As LongPtr: DialogHwnd = FindWindow(vbNullString, "開く")
    Dim InputHwnd As Long
    InputHwnd = FindWindowEx(DialogHwnd, 0&, "ComboBoxEx32", "")
    InputHwnd = FindWindowEx(InputHwnd, 0&, "ComboBox", "")
    InputHwnd = FindWindowEx(InputHwnd, 0&, "Edit", "")
    Call SendMessage(InputHwnd, &HC, 0, ByVal filePath)
 
    Dim ButtonHwnd As Long
    ButtonHwnd = FindWindowEx(DialogHwnd, 0&, "Button", "開く(&O)")
    Call SendMessage(ButtonHwnd, &HF5, 0, 0&)
End Sub

Private Sub IWebElement_ClickAndThenAlertDialogErase()
    chrome_.Focus (objectId_)
    Dim responseJson As String
    responseJson = chrome_.JavaScritpCallFunction("function(value){this.click();}", objectId_, "")
    chrome_.SleepByWindowsAPIInner 1000
    
    Dim ChromeInterFace As IWebDriver
    Set ChromeInterFace = chrome_
    ChromeInterFace.HandleAlertDialog
End Sub



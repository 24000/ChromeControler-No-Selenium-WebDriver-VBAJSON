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

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hWndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Any) As LongPtr

Private chrome_ As ChromeDriver

Private objectId_ As String
Private findInfo_ As Collection
Private nodeId_ As Long
Private attributes_ As Object
Private frameId_ As String

Private xpath_ As String
Private name_


Public Sub Init(chrome As ChromeDriver, _
                        objectId As String, findInfo As Collection, _
                        Optional attributes As Object = Nothing, Optional FrameId As String)
    Set chrome_ = chrome
    objectId_ = objectId
    Set findInfo_ = findInfo
    Set attributes_ = attributes
    frameId_ = FrameId
End Sub

Public Property Get objectId() As String
    objectId = objectId_
End Property
Public Property Get FrameId() As String
    FrameId = frameId_
End Property

'CallJavaScriptFunction呼出しの共通関数
'戻り値が"Cannot find context with specified id"となる場合、ObjectIdの再取得が必要なため共通化
Private Function CallFunction(functionCode As String, args As String) As String
    Dim responseMsg As String
    responseMsg = chrome_.CallJavaScriptFunctionForElement(functionCode, objectId_, args)
    If InStr(responseMsg, "Cannot find context with specified id") > 0 Then
        chrome_.ReCreateIsolatedFrameWorld
        ReGetObjectId
        responseMsg = chrome_.CallJavaScriptFunctionForElement(functionCode, objectId_, args)
    End If
    CallFunction = responseMsg
End Function
Public Function ReGetObjectId()
    Dim Codes As Variant
    Select Case findInfo_(1)
      Case "id"
        Codes = chrome_.GetIdCodes(findInfo_(2))
      Case "xpath"
        Codes = chrome_.GetXpathCodes(findInfo_(2))
      Case "name"
        Codes = chrome_.GetNameCodes(findInfo_(2))
      Case "tag"
        Codes = chrome_.GetTagCodes(findInfo_(2), findInfo_(3))
    End Select
    objectId_ = chrome_.GetObjectIdOfElementBy(findInfo_(2), Codes)
End Function

'ChromeDriverのCDPMethodを呼出す場合に、引数にObjectIdを使用して呼出すMethodの共通管理のための関数
'戻り値が"Cannot find context with specified id"となる場合に、ObjectIdの再取得が必要になるため共通化
'現在FocusとGetGetBoxModelの呼び出しのみ
Private Function CallObjectIdUsingMethod(methodName As String) As String
    Dim responseMsg As String
    responseMsg = CallByName(chrome_, methodName, VbMethod, objectId_)

    If InStr(responseMsg, "Cannot find context with specified id") > 0 Then
        chrome_.ReCreateIsolatedFrameWorld
        ReGetObjectId
        responseMsg = CallByName(chrome_, methodName, VbMethod, objectId_)
    End If
    CallObjectIdUsingMethod = responseMsg
End Function

'***********************************************
'Public
'チェックボックス要素についてChecked属性の設定
Private Property Let IWebElement_Checked(RHS As Boolean)
    Dim responseJson As String
    'checkedプロパティを持っているか事前にチェック。持っていない場合はじく。
    responseJson = CallFunction("function(){return this.checked;}", "")
    If InStr(responseJson, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim boolString As String
    If RHS Then
        boolString = "true"
    Else
        boolString = "false"
    End If

    responseJson = CallFunction("function(value){ this.checked = value;}", boolString)

Exit Property
quit:
chrome_.DisplayErrorAndEnd "このElementはCheckedプロパティを持ちません。" & vbCrLf & _
                                            "※CheckedはCheckBoxで使用可能" & vbCrLf
End Property
'Public
'チェックボックス要素についてチェック状態取得
Private Property Get IWebElement_Checked() As Boolean
    Dim responseJson As String
    responseJson = CallFunction("function(){return this.checked;}", "")
    If InStr(responseJson, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim Checked As Boolean
    Checked = chrome_.json.GetValueFrom(responseJson, Array("result", "result", "value"))
    IWebElement_Checked = Checked
Exit Property

quit:
chrome_.DisplayErrorAndEnd "このElementはCheckedプロパティを持ちません。" & vbCrLf & _
                                            "※CheckedはCheckBoxで使用可能" & vbCrLf
End Property







'Public
'要素の状態(有効or無効)を取得
Private Property Get IWebElement_IsEnabled() As Boolean
    Dim responseJson As String
    responseJson = CallFunction("function(){return this.disabled;}", "")
    If InStr(responseJson, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim Enabled As Boolean
    Enabled = chrome_.json.GetValueFrom(responseJson, Array("result", "result", "value"))
    IWebElement_IsEnabled = Not Enabled
Exit Property
    
quit:
chrome_.DisplayErrorAndEnd "このElementはEnabledプロパティを持ちません。" & vbCrLf
End Property



'Public
'javaScriptによってクリック実行
Private Sub IWebElement_Click()
    CallObjectIdUsingMethod "Focus"

    Dim responseJson As String
    responseJson = CallFunction("function(){this.click();}", "")
    
    If InStr(responseJson, "this.click is not a function") Then
        responseJson = CallFunction("function(){this.dispatchEvent(new MouseEvent( 'click' ));}", "")
    End If
    
    If InStr(responseJson, "this.dispatchEvent is not a function") Then
        CallObjectIdUsingMethod "Focus"
        chrome_.DispatchKeyEvent_DOMKeyCode "keyDown", "Enter"
        chrome_.DispatchKeyEvent_DOMKeyCode "keyUp", "Enter"
    End If
    
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
    Loop While (isEndChange = False And failSafeCounter <= 4)
    
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
    Loop While (isEndChange = False And failSafeCounter <= 4)
    
    chrome_.CDPEvents.IsIsolatedFrameScheduledNavigation = False
    chrome_.CDPEvents.IsIsolatedFrameStartedLoading = False
    chrome_.CDPEvents.IsIsolatedFrameClearedScheduledNavigation = False
    chrome_.CDPEvents.IsIsolatedFrameNavigated = False
End Sub





'Public
'要素のValue属性に文字列を設定
Private Sub IWebElement_SetText(text As String)
    CallObjectIdUsingMethod "Focus"
    chrome_.DispatchKeyEvent_DOMKeyCode "keyDown", "Enter"
    chrome_.DispatchKeyEvent_DOMKeyCode "keyUp", "Enter"
    
    Dim encodingText As String: encodingText = chrome_.json.EncodeURL(text)
    CallFunction _
        "function(value){value = decodeURI(value);this.value = value;}", encodingText
    
    CallFunction _
    "function(value){" & _
        "const event = new CustomEvent('change', {bubbles: false, cancelable: false});" & _
        "this.dispatchEvent(event);" & _
    "}", encodingText

End Sub

'Public
'要素のTextContentに文字列を設定
Private Sub IWebElement_SetTextContent(text As String)
    Dim encodingText As String: encodingText = chrome_.json.EncodeURL(text)
    CallFunction _
        "function(value){value = decodeURI(value);this.textContent  = value;}", encodingText
    
     CallFunction _
    "function(value){" & _
        "const event = new CustomEvent('change', {bubbles: false, cancelable: false});" & _
        "this.dispatchEvent(event);" & _
    "}", encodingText
End Sub

'Public
'SelectBoxの選択肢を文字列で選択
Private Sub IWebElement_SelectItemInSelectBoxByText(text As String)
    Dim encodingText As String: encodingText = chrome_.json.EncodeURL(text)
    Dim javaScriptCode As String
    javaScriptCode = _
    "function(value){value = decodeURI(value);" & _
        "for(let i = 0; i < this.options.length; i++){" & _
            "let val1 = this.options[i].value.replace(/\\r?\\n/g,'');" & _
            "let val2 = this.options[i].textContent.replace(/\\r?\\n/g,'');" & _
            "console.log('(you)' + value + ':(value)' + val1 + ':(textContent)'+ val2);" & _
            "if(val1 === value || val2 === value){" & _
                "this.options[i].selected = true;" & _
                "return true;" & _
            "}" & _
        "}" & _
    "return false;}"
    
    Dim responseJson As String
    responseJson = CallFunction(javaScriptCode, encodingText)
    
    Dim IsSelect As Boolean
    IsSelect = chrome_.json.GetValueFrom(responseJson, Array("result", "result", "value"))
    
    If IsSelect = False Then
        MsgBox "「" & text & "」は選択肢の中に存在しません。" & vbCrLf & _
                    "処理を終了します。"
        chrome_.WebSocket.Http.CloseHInternetHandles
        End
    End If
    
    javaScriptCode = _
    "function(value){this.onchange();}"
    responseJson = CallFunction(javaScriptCode, encodingText)
    
    If InStr(responseJson, "this.onchange is not a function") Then
        javaScriptCode = _
        "function(value){" & _
            "const event = new CustomEvent('change', {bubbles: false, cancelable: false});" & _
            "this.dispatchEvent(event);" & _
        "}"
        responseJson = CallFunction(javaScriptCode, encodingText)
    End If

End Sub

'Public
Private Function IWebElement_FindElementByName(name As String) As IWebElement
    Dim functionCode As String, responseJson As String
    functionCode = "function(name){return  this.getElementsByName(name)[0];}"
    responseJson = CallFunction(functionCode, name)
    
    Dim findObjectId As String
    findObjectId = chrome_.json.GetValueFrom( _
                            responseJson, Array("result", "result", "objectId"))
    If findObjectId = "" Then GoTo quit
    Dim findInfoInElem As Collection: Set findInfoInElem = chrome_.MakeFindInfo("name", name)
    Set IWebElement_FindElementByName = chrome_.MakeElementObject(findObjectId, findInfoInElem)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & name
Beep
chrome_.DisplayErrorAndEnd chrome_.MakeFindErrorMessage("Name", name)
End Function

'Public
Private Function IWebElement_FindElementsByTag(tagName As String) As IWebElements
    Dim elemsCount As Long
    elemsCount = GetElementsCount(tagName)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    Dim Index As Long
    For Index = 0 To elemsCount - 1
        elements.add CreateElementsItemByIndex(tagName, Index)
    Next
    
    Dim elems As WebElements: Set elems = New WebElements
    elems.Init elements
    
    Set IWebElement_FindElementsByTag = elems
Exit Function
quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & tagName
Beep
chrome_.DisplayErrorAndEnd chrome_.MakeFindErrorMessage("Tag", tagName, "○○")
End Function
Private Function GetElementsCount(tagName As String) As Long
    Dim javaScriptCode As String, responseJson As String
    javaScriptCode = _
    "function(tagName){return  this.getElementsByTagName('" & tagName & "').length;}"
    responseJson = CallFunction(javaScriptCode, tagName)
    GetElementsCount = chrome_.json.GetValueFrom(responseJson, Array("result", "result", "value"))
End Function
Private Function CreateElementsItemByIndex(tagName As String, Index As Long) As IWebElement
    Dim javaScriptCode As String, responseJson As String
    Dim objectId As String
    Dim elem As IWebElement

    javaScriptCode = _
    "function(tagName){return  this.getElementsByTagName(tagName)[" & Index & "];}"
    responseJson = CallFunction(javaScriptCode, tagName)

    objectId = chrome_.json.GetValueFrom(responseJson, Array("result", "result", "objectId"))
    Dim findInfoInElem As Collection: Set findInfoInElem = chrome_.MakeFindInfo("tag", tagName, Index)
    Set CreateElementsItemByIndex = chrome_.MakeElementObject(objectId, findInfoInElem)
End Function

'Public
Private Function IWebElement_GetTextContent() As String
    Dim javaScriptCode As String, responseJson As String
    javaScriptCode = "function(){return this.textContent.replace(/\\s/g, ' ');}"
    responseJson = CallFunction(javaScriptCode, "")
    
    Dim text As String
    text = chrome_.json.GetValueFrom(responseJson, Array("result", "result", "value"))
    IWebElement_GetTextContent = text
End Function

'Public
Private Function IWebElement_GetSelectedValue() As String
    Dim javaScriptCode As String, responseJson As String
    javaScriptCode = "function(){return this.value.replace(/\\s/g, ' ');}"
    responseJson = CallFunction(javaScriptCode, "")
    
    Dim text As String
    text = chrome_.json.GetValueFrom(responseJson, Array("result", "result", "value"))
    IWebElement_GetSelectedValue = text
End Function

'Public
Private Function IWebElement_GetSelectedTextContent() As String
    Dim javaScriptCode As String, responseJson As String
    javaScriptCode = _
    "function(){" & _
        "for(let i = 0; i < this.options.length; i++){" & _
            "if(this.options[i].selected){" & _
                "return this.options[i].textContent.replace(/\\s/g, ' ');" & _
            "}" & _
        "}" & _
    "}"
    responseJson = CallFunction(javaScriptCode, "")
    
    Dim text As String
    text = chrome_.json.GetValueFrom(responseJson, Array("result", "result", "value"))
    IWebElement_GetSelectedTextContent = text
End Function

'要素中心位置のxy座標をCDPMethodで左クリック発行
Private Sub IWebElement_DispatchLeftClickEvent()
    Dim responseMsg As String: responseMsg = CallObjectIdUsingMethod("GetBoxModel")
    Dim centerXY As Variant: centerXY = GetCenterPositionXY(responseMsg)
    Dim x As Long: x = centerXY(0)
    Dim y As Long: y = centerXY(1)
    
    Dim ChromeInterFace As IWebDriver: Set ChromeInterFace = chrome_
    ChromeInterFace.ScrollWindow y - 50
    
    Dim Result As String
    Result = chrome_.DispatchMouseEvent("mousePressed", x, y, "left", 1)
    Result = chrome_.DispatchMouseEvent("mouseReleased", x, y, "left", 1)
End Sub

Public Function GetCenterPositionXY(responseMsg As String)
    Dim Result As Variant
    Result = chrome_.json.GetValueFrom(responseMsg, Array("result", "model", "content"))
    Result = Split(Result, ",")
    Dim width As Long, height As Long
    width = chrome_.json.GetValueFrom(responseMsg, Array("result", "model", "width"))
    height = chrome_.json.GetValueFrom(responseMsg, Array("result", "model", "height"))
    Dim xy(1) As Variant
    xy(0) = Result(0) + (width / 2)
    xy(1) = Result(1) + (height / 2)
    GetCenterPositionXY = xy
End Function

'Public
'type=fileのInputBoxへファイル設定
Private Sub IWebElement_SetFileToFileSelectInputButton(filePath As String)
'    filePath = Replace(filePath, "\", "\\")
'    filePath = chrome_.Json.EncodeURL(filePath)
'    filePath = Replace(filePath, "%5C", "\")
'    chrome_.SetFileInputFiles objectId_,
'CDPMethodで設定しようとしても、ファイルパスに日本語が含まれる場合、
'エンコードして送信はできるが、クローム側でデコードしてくれない為、没
    Dim responseMsg As String: responseMsg = CallObjectIdUsingMethod("GetBoxModel")
    Dim centerXY As Variant: centerXY = GetCenterPositionXY(responseMsg)
    
    Dim ChromeInterFace As IWebDriver: Set ChromeInterFace = chrome_
    ChromeInterFace.ScrollWindow CLng(centerXY(1)) - 50
    
    IWebElement_DispatchLeftClickEvent
    chrome_.SleepByWindowsAPIInner 1000
    ControlFileSelectDialog filePath
    chrome_.SleepByWindowsAPIInner 1000
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

'Public
Private Sub IWebElement_ClickAndThenAlertDialogErase()
    CallObjectIdUsingMethod "Focus"
'    chrome_.Focus (objectId_)
    Dim responseJson As String
    responseJson = CallFunction("function(){this.click();}", "")
    chrome_.SleepByWindowsAPIInner 1000
    
    Dim ChromeInterFace As IWebDriver
    Set ChromeInterFace = chrome_
    ChromeInterFace.HandleAlertDialog
End Sub



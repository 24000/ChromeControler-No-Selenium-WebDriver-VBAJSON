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

Private chrome_ As ChromeDriver

Private objectId_ As String
Private nodeId_ As Long
Private attributes_ As Object
Private frameId_ As String


Public Sub Init(chrome As ChromeDriver, _
                        objectId As String, Optional nodeId As Long = 0, _
                        Optional attributes As Object = Nothing, Optional frameId As String)
    Set chrome_ = chrome
    objectId_ = objectId
    nodeId_ = nodeId
    Set attributes_ = attributes
    frameId_ = frameId
End Sub



'***********************************************
'Public
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
Private Property Get IWebElement_Checked() As Boolean
    Dim responseJson As String
    responseJson = chrome_.JavaScritpCallFunction( _
        "function(value){return this.checked;}", objectId_, "")
    If InStr(responseJson, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim Checked As Boolean
    Checked = chrome_.jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "value"))
    IWebElement_Checked = Checked
Exit Property

quit:
chrome_.DisplayErrorAndEnd "このElementはCheckedを持ちません。" & vbCrLf & _
                                            "※CheckedはCheckBoxで使用可能"
End Property

'Public
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
    
    If IsPageChanging Then WaitPageChangeToComplete
    If IsIsolatedFrameChanging Then WaitIsolatedFrameChangeToComplete
    
End Sub
Private Function IsPageChanging() As Boolean
    '一回無害なメソッドを投げて、イベント発生状況更新
    chrome_.PageEnable
    
    If chrome_.CDPEvents.IsCurrentTargetScheduledNavigation _
    Or chrome_.CDPEvents.IsCurrentTargetStartedLoading Then
        IsPageChanging = True
    Else
        IsPageChanging = False
    End If
End Function
Private Sub WaitPageChangeToComplete()
    Dim chromeInterFace As IWebDriver
    Set chromeInterFace = chrome_
    
    Dim isEndChane As Boolean
    Do
        '一回無害なメソッドを投げて、イベント発生状況更新
        chrome_.PageEnable
        isEndChane = _
            chrome_.CDPEvents.IsCurrentTargetClearedScheduledNavigation _
            Or chrome_.CDPEvents.IsCurrentTargetNavigated
            
        If isEndChane Then
            chrome_.CDPEvents.IsCurrentTargetScheduledNavigation = False
            chrome_.CDPEvents.IsCurrentTargetStartedLoading = False
            chrome_.CDPEvents.IsCurrentTargetClearedScheduledNavigation = False
            chrome_.CDPEvents.IsCurrentTargetNavigated = False
            Exit Do
        End If
        
        chromeInterFace.SleepByWindowsAPI 1000
    Loop While (isEndChane = False)
End Sub

Private Function IsIsolatedFrameChanging() As Boolean
    '一回無害なメソッドを投げて、イベント発生状況更新
    chrome_.PageEnable
    
    If chrome_.CDPEvents.IsIsolatedFrameScheduledNavigation _
    Or chrome_.CDPEvents.IsIsolatedFrameStartedLoading Then
        IsIsolatedFrameChanging = True
    Else
        IsIsolatedFrameChanging = False
    End If
End Function
Private Sub WaitIsolatedFrameChangeToComplete()
    Dim chromeInterFace As IWebDriver
    Set chromeInterFace = chrome_
    
    Dim isEndChane As Boolean
    Do
        '一回無害なメソッドを投げて、イベント発生状況更新
        chrome_.PageEnable
        isEndChane = _
            chrome_.CDPEvents.IsIsolatedFrameClearedScheduledNavigation _
            Or chrome_.CDPEvents.IsIsolatedFrameNavigated
            
        If isEndChane Then
            chrome_.CDPEvents.IsIsolatedFrameScheduledNavigation = False
            chrome_.CDPEvents.IsIsolatedFrameStartedLoading = False
            chrome_.CDPEvents.IsIsolatedFrameClearedScheduledNavigation = False
            chrome_.CDPEvents.IsIsolatedFrameNavigated = False
            Exit Do
        End If
        
        chromeInterFace.SleepByWindowsAPI 1000
    Loop While (isEndChane = False)
End Sub


'Public
Private Sub IWebElement_SetText(text As String)
    chrome_.Focus (objectId_)
    chrome_.DispatchKeyEvent_DOMKeyCode "keyUp", "Enter"

    Dim encodingText As String: encodingText = chrome_.jsonhandler.EncodeURL(text)
    chrome_.JavaScritpCallFunction _
    "function(value){" & _
        "value = decodeURI(value);" & _
        "this.value = value;" & _
    "}", _
    objectId_, encodingText
End Sub
'Public
Private Sub IWebElement_SetTextContent(text As String)
    Dim encodingText As String: encodingText = chrome_.jsonhandler.EncodeURL(text)
    chrome_.JavaScritpCallFunction _
    "function(value){" & _
        "value = decodeURI(value);" & _
        "this.textContent  = value;" & _
    "}", _
    objectId_, encodingText
End Sub
'Public
Private Sub IWebElement_SelectItemByText(text As String)
    Dim encodingText As String: encodingText = chrome_.jsonhandler.EncodeURL(text)
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
    IsSelect = chrome_.jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "value"))
    
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
    
    Dim chromeInterFace As IWebDriver
    Set chromeInterFace = chrome_
    chromeInterFace.SleepByWindowsAPI 2000
End Sub

'Public
Private Function IWebElement_GetElementByName(Name As String) As IWebElement
    Dim functionCode As String, responseJson As String
    functionCode = "function(name){return  this.getElementsByName(name)[0];}"
    responseJson = chrome_.JavaScritpCallFunction(functionCode, objectId_, Name)
    
    Dim findObjectId As String
    findObjectId = chrome_.jsonhandler.GetValueFromJson( _
                            responseJson, Array("result", "result", "objectId"))
    If findObjectId = "" Then GoTo quit
    Set IWebElement_GetElementByName = chrome_.MakeElementObject(findObjectId)
Exit Function

quit:
chrome_.DisplayErrorAndEnd "指定された以下Nameの要素は存在しませんでした。" & vbCrLf & Name
End Function

'Public
Private Function IWebElement_GetElementsByTagName(tagName As String) As IWebElements
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
    
    Set IWebElement_GetElementsByTagName = elems
Exit Function
quit:
chrome_.DisplayErrorAndEnd "指定された以下TagNameの要素は存在しませんでした。" & vbCrLf & tagName
End Function
Private Function GetElementsCount(tagName As String) As Long
    Dim javaScriptCode As String, responseJson As String
    javaScriptCode = _
    "function(tagName){return  this.getElementsByTagName('" & tagName & "').length;}"
    responseJson = chrome_.JavaScritpCallFunction(javaScriptCode, objectId_, tagName)
    GetElementsCount = chrome_.jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "value"))
End Function
Private Function CreateElementsItemByIndex(tagName As String, index As Long) As IWebElement
    Dim javaScriptCode As String, responseJson As String
    Dim objectId As String
    Dim elem As IWebElement

    javaScriptCode = _
    "function(tagName){return  this.getElementsByTagName(tagName)[" & index & "];}"
    responseJson = chrome_.JavaScritpCallFunction(javaScriptCode, objectId_, tagName)

    objectId = chrome_.jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "objectId"))
    Set CreateElementsItemByIndex = chrome_.MakeElementObject(objectId)
End Function

'Public
Private Function IWebElement_GetTextContent() As String
    Dim javaScriptCode As String, responseJson As String
    javaScriptCode = "function(value){return this.textContent.replace(/\\s/g, ' ');}"
    responseJson = chrome_.JavaScritpCallFunction(javaScriptCode, objectId_, "")
    
    Dim text As String
    text = chrome_.jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "value"))
    IWebElement_GetTextContent = text
End Function

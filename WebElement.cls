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

Private chrome_ As ChromeControler
Private webSocket_ As a_0_x1_WebSocketCommunicator
Private CDPMethod_ As a_0_x2_CDPMethods
Private jsonHandler_ As a_0_x4_JSONHandler

Private objectId_ As String
Private nodeId_ As Long

Public Sub Init(chrome As ChromeControler, objectId As String, _
                        Optional nodeId As Long = 0)
    Set chrome_ = chrome
    Set webSocket_ = chrome.WebSocket
    Set CDPMethod_ = chrome.CDPMethod
    Set jsonHandler_ = chrome.jsonhandler
    
    objectId_ = objectId
    nodeId_ = nodeId
End Sub


'***********************************************
'Public
Private Sub IWebElement_Click()
    chrome_.JavaScritpCallFunction "function(value){this.click();}", objectId_, ""
End Sub

'Public
Private Sub IWebElement_SetText(text As String)
    Dim encodingText As String: encodingText = jsonHandler_.EncodeURL(text)
    chrome_.JavaScritpCallFunction _
    "function(value){" & _
        "value = decodeURI(value);" & _
        "this.value = value;" & _
    "}", _
    objectId_, encodingText
End Sub
'Public
Private Sub IWebElement_SetTextContent(text As String)
    Dim encodingText As String: encodingText = jsonHandler_.EncodeURL(text)
    chrome_.JavaScritpCallFunction _
    "function(value){" & _
        "value = decodeURI(value);" & _
        "this.textContent  = value;" & _
    "}", _
    objectId_, encodingText
End Sub
'Public
Private Sub IWebElement_SelectItemByText(text As String)
    Dim encodingText As String: encodingText = jsonHandler_.EncodeURL(text)
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
    IsSelect = jsonHandler_.GetValueFromJson(responseJson, Array("result", "result", "value"))
    
    If IsSelect = False Then
        MsgBox "「" & text & "」は選択肢の中に存在しません。" & vbCrLf & _
                    "処理を終了します。"
        webSocket_.Http.CloseHInternetHandles
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
    
    chrome_.SleepByWindowsAPI 2000
End Sub
'Public
Private Function IWebElement_GetElementByXpath(xpath As String) As IWebElement

End Function
'Public
Private Function IWebElement_GetElementsByTagName(tagName As String) As IWebElements

End Function
'Public
Private Function IWebElement_GetTextContent() As String
    Dim javaScriptCode As String, responseJson As String
    javaScriptCode = "function(value){return this.textContent;}"
    responseJson = chrome_.JavaScritpCallFunction(javaScriptCode, objectId_, "")
    
    Dim text As String
    text = jsonHandler_.GetValueFromJson(responseJson, Array("result", "result", "value"))
'    text = jsonHandler_.DecodeURL(text)
    IWebElement_GetTextContent = text
End Function

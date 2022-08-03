VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "b0x0_WebElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IWebElement

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hWndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Any) As LongPtr

Private Json_ As a2_JSONHandler
Private basicInfo_ As a3_BasicInfos
Private CDP_ As a4_ExecuteCDP
Private events_ As a5_CDPEventHandler
Private Helper_ As a6_ExecuteHelperFunction

Private objectId_ As String
Private findInfo_ As Collection
Private nodeId_ As Long
Private attributes_ As Object
Private frameId_ As String

Private xpath_ As String
Private name_


Public Sub Init(Json As a2_JSONHandler, _
                       basicInfo As a3_BasicInfos, _
                       ExecCDP As a4_ExecuteCDP, _
                       events As a5_CDPEventHandler, _
                       execHelper As a6_ExecuteHelperFunction, _
                       ObjectId As String, _
                       findInfo As Collection, _
                       Optional attributes As Object = Nothing, _
                       Optional FrameId As String)
                       
    Set Json_ = Json
    Set basicInfo_ = basicInfo
    Set CDP_ = ExecCDP
    Set events_ = events
    Set Helper_ = execHelper
    
    objectId_ = ObjectId
    Set findInfo_ = findInfo
    Set attributes_ = attributes
    frameId_ = FrameId
End Sub

Public Property Get ObjectId() As String
    ObjectId = objectId_
End Property
Public Property Get FrameId() As String
    FrameId = frameId_
End Property

'CallJavaScriptFunction呼出しの共通関数
'戻り値が"Cannot find context with specified id"となる場合、ObjectIdの再取得が必要なため共通化
Private Function CallFunction(functionCode As String, args As String) As String
    Dim res As String
    res = CDP_.CallJavascriptFunction(functionCode, objectId_, args)
    If InStr(res, "Cannot find context with specified id") > 0 Then
        basicInfo_.UpdateExecutionContextId CDP_
        ReGetObjectId
        res = CDP_.CallJavascriptFunction(functionCode, objectId_, args)
    End If
    CallFunction = res
End Function
Public Function ReGetObjectId()
    Dim codes As Variant
    Select Case findInfo_(1)
      Case "id"
        codes = Helper_.GetIdCodes(findInfo_(2))
      Case "xpath"
        codes = Helper_.GetXpathCodes(findInfo_(2))
      Case "name"
        codes = Helper_.GetNameCodes(findInfo_(2))
      Case "tag"
        codes = Helper_.GetTagCodes(findInfo_(2), findInfo_(3))
    End Select
    objectId_ = CDP_.GetObjectIdOfElementBy(findInfo_(2), basicInfo_, codes)
End Function

'CDPMethodを呼出す場合に、引数にObjectIdを使用して呼出すMethodの共通管理のための関数
'戻り値が"Cannot find context with specified id"となる場合に、ObjectIdの再取得が必要になるため共通化
'現在FocusとGetGetBoxModelの呼び出しのみ
Private Function CallObjectIdUsingMethod(methodName As String) As String
    Dim res As String
    res = CallByName(CDP_, methodName, VbMethod, objectId_)

    If InStr(res, "Cannot find context with specified id") > 0 Then
        basicInfo_.UpdateExecutionContextId CDP_
        ReGetObjectId
        res = CallByName(CDP_, methodName, VbMethod, objectId_)
    End If
    CallObjectIdUsingMethod = res
End Function

'***********************************************
'Public
'チェックボックス要素についてChecked属性の設定
Private Property Let IWebElement_Checked(RHS As Boolean)
    Dim res As String
    'checkedプロパティを持っているか事前にチェック。持っていない場合はじく。
    res = CallFunction("function(){return this.checked;}", "")
    If InStr(res, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim boolString As String
    If RHS Then
        boolString = "true"
    Else
        boolString = "false"
    End If

    res = CallFunction("function(value){ this.checked = value;}", boolString)
Exit Property

quit:
Helper_.DisplayErrorAndEnd _
    "<WebElement：Checkedプロパティ>" & vbCrLf & _
    "このElementはCheckedプロパティを持ちません。" & vbCrLf & _
    "※CheckedはCheckBoxで使用可能" & vbCrLf
End Property
'Public
'チェックボックス要素についてチェック状態取得
Private Property Get IWebElement_Checked() As Boolean
    Dim res As String
    res = CallFunction("function(){return this.checked;}", "")
    If InStr(res, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim Checked As Boolean
    Checked = Json_.GetValue(res, "result", "result", "value")
    IWebElement_Checked = Checked
Exit Property

quit:
Helper_.DisplayErrorAndEnd "このElementはCheckedプロパティを持ちません。" & vbCrLf & _
                                            "※CheckedはCheckBoxで使用可能" & vbCrLf
End Property









'Public
'要素の状態(有効or無効)を取得
Private Property Get IWebElement_IsEnabled() As Boolean
    Dim res As String
    res = CallFunction("function(){return this.disabled;}", "")
    If InStr(res, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim Enabled As Boolean
    Enabled = Json_.GetValue(res, "result", "result", "value")
    IWebElement_IsEnabled = Not Enabled
Exit Property
    
quit:
Helper_.DisplayErrorAndEnd "このElementはEnabledプロパティを持ちません。" & vbCrLf
End Property



'Public
'javaScriptによってクリック実行
Private Sub IWebElement_Click()
    CallObjectIdUsingMethod "Focus"

    Dim res As String
    res = CallFunction("function(){this.click();}", "")
    
    If InStr(res, "this.click is not a function") Then
        res = CallFunction("function(){this.dispatchEvent(new MouseEvent( 'click' ));}", "")
    End If
    
    If InStr(res, "this.dispatchEvent is not a function") Then
        FocusAndEnterKey
    End If
    
    If IsPageChanging Then WaitPageChangeToComplete
    If IsIsolatedFrameChanging Then WaitIsolatedFrameChangeToComplete
End Sub
Private Sub FocusAndEnterKey()
    CallObjectIdUsingMethod "Focus"
    CDP_.DispatchKeyEvent_DOMKeyCode "keyDown", "Enter"
    CDP_.DispatchKeyEvent_DOMKeyCode "keyUp", "Enter"
End Sub


Private Function IsPageChanging() As Boolean
    CDP_.PageEnable '一回無害なメソッドを投げて、イベント発生状況更新
    
    If events_.IsCurrentTargetScheduledNavigation _
    Or events_.IsCurrentTargetStartedLoading Then
        IsPageChanging = True
    Else
        IsPageChanging = False
    End If
End Function
Private Sub WaitPageChangeToComplete()
    
    'TargetNavigatedは返って来ないケースがある？ので繰り返し回数に制限をかけてループ
    Dim failSafeCounter As Long
    Dim isEndChange As Boolean
    Do
        isEndChange = events_.IsCurrentTargetNavigated
'                                Or　chrome_.CDPEvents.IsCurrentTargetClearedScheduledNavigation
        If isEndChange Then Exit Do
        
        CDP_.RuntimeEnable
        Sleep 1000
        failSafeCounter = failSafeCounter + 1
    Loop While (isEndChange = False And failSafeCounter <= 4)
    
    events_.IsCurrentTargetScheduledNavigation = False
    events_.IsCurrentTargetStartedLoading = False
    events_.IsCurrentTargetClearedScheduledNavigation = False
    events_.IsCurrentTargetNavigated = False
End Sub
Private Function IsIsolatedFrameChanging() As Boolean
    CDP_.DOMEnable '一回無害なメソッド(frameなのでDOM系)を投げて、イベント発生状況更新
    
    If events_.IsIsolatedFrameScheduledNavigation _
    Or events_.IsIsolatedFrameStartedLoading Then
        IsIsolatedFrameChanging = True
    Else
        IsIsolatedFrameChanging = False
    End If
End Function
Private Sub WaitIsolatedFrameChangeToComplete()
    
    'FrameNavigatedは返って来ないケースがある？ので繰り返し回数に制限をかけてループ
    Dim failSafeCounter As Long
    Dim isEndChange As Boolean
    Do
        isEndChange = events_.IsIsolatedFrameNavigated
'                               Or chrome_.CDPEvents.IsIsolatedFrameClearedScheduledNavigation
        If isEndChange Then Exit Do
        
        CDP_.DOMEnable
        CDP_.RuntimeEnable
        Sleep 1000
        failSafeCounter = failSafeCounter + 1
    Loop While (isEndChange = False And failSafeCounter <= 4)
    
    events_.IsIsolatedFrameScheduledNavigation = False
    events_.IsIsolatedFrameStartedLoading = False
    events_.IsIsolatedFrameClearedScheduledNavigation = False
    events_.IsIsolatedFrameNavigated = False
End Sub





'Public
'要素のValue属性に文字列を設定
Private Sub IWebElement_SetText(text As String)
    FocusAndEnterKey
    Dim encodingText As String: encodingText = Json_.EncodeURL(text)
    CallFunction "function(value){value = decodeURI(value);this.value = value;}", encodingText
    
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText
End Sub

'Public
'要素のTextContentに文字列を設定
Private Sub IWebElement_SetTextContent(text As String)
    FocusAndEnterKey
    Dim encodingText As String: encodingText = Json_.EncodeURL(text)
    CallFunction "function(value){value = decodeURI(value);this.textContent  = value;}", encodingText
    
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText
End Sub

'Public
'SelectBoxの選択肢を文字列で選択
Private Sub IWebElement_SelectItemInSelectBoxByText(text As String)
    Dim encodingText As String: encodingText = Json_.EncodeURL(text)
    
    Dim res As String: res = CallFunction(Helper_.GetSelectItemCode, encodingText)
    
    Dim IsSelect As Boolean: IsSelect = Json_.GetValue(res, "result", "result", "value")
    If IsSelect = False Then
        MsgBox "<SelectItemInSelectBoxByText>" & vbCrLf & _
                    "「" & text & "」は選択肢の中に存在しません。" & vbCrLf & "処理を終了します。"
        End
    End If
    
'    CallFunction "function(value){this.onchange();}", encodingText
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText
End Sub

'Public
Private Function IWebElement_FindElementByName(name As String) As IWebElement
    Dim funcCode As String, res As String
    funcCode = "function(name){return  this.getElementsByName(name)[0];}"
    res = CallFunction(funcCode, name)
    
    Dim ObjectId As String
    ObjectId = Json_.GetValue(res, "result", "result", "objectId")
    If ObjectId = "" Then GoTo quit
    
    Dim findInfoInElem As Collection
    Set findInfoInElem = Helper_.MakeFindInfo("name", name)
    
    Set IWebElement_FindElementByName = _
        CDP_.MakeElementObject(ObjectId, findInfoInElem, basicInfo_, events_, Helper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & name
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Name", name)
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
    
    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements
    
    Set IWebElement_FindElementsByTag = elems
Exit Function
quit:
Debug.Print vbCrLf & ">>見つからなかった要素" & vbCrLf & tagName
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Tag", tagName, "○○")
End Function
Private Function GetElementsCount(tagName As String) As Long
    Dim funcCode As String, res As String
    funcCode = "function(tagName){return  this.getElementsByTagName('" & tagName & "').length;}"
    res = CallFunction(funcCode, tagName)
    GetElementsCount = Json_.GetValue(res, "result", "result", "value")
End Function
Private Function CreateElementsItemByIndex(tagName As String, Index As Long) As IWebElement
    Dim funcCode As String, res As String
    funcCode = "function(tagName){return  this.getElementsByTagName(tagName)[" & Index & "];}"
    res = CallFunction(funcCode, tagName)
    
    Dim ObjectId As String
    ObjectId = Json_.GetValue(res, "result", "result", "objectId")
    Dim findInfoInElem As Collection: Set findInfoInElem = Helper_.MakeFindInfo("tag", tagName, Index)

    Set CreateElementsItemByIndex = _
        CDP_.MakeElementObject(ObjectId, findInfoInElem, basicInfo_, events_, Helper_)
End Function

'Public
'エレメントのTextContentを取得
Private Function IWebElement_GetTextContent() As String
    Dim funcCode As String, res As String
    funcCode = "function(){return this.textContent.replace(/\\s/g, ' ');}"
    res = CallFunction(funcCode, "")
    
    IWebElement_GetTextContent = Json_.GetValue(res, "result", "result", "value")
End Function

'Public
'セレクトボックスの選択中オプションのValueを取得
Private Function IWebElement_GetSelectedValue() As String
    Dim funcCode As String, res As String
    funcCode = "function(){return this.value.replace(/\\s/g, ' ');}"
    res = CallFunction(funcCode, "")

    IWebElement_GetSelectedValue = Json_.GetValue(res, "result", "result", "value")
End Function

'Public
'セレクトボックスの選択中オプションのTextContentを取得
Private Function IWebElement_GetSelectedTextContent() As String
    Dim funcCode As String, res As String
    funcCode = Helper_.GetSelectedTextContentCode
    res = CallFunction(funcCode, "")

    IWebElement_GetSelectedTextContent = Json_.GetValue(res, "result", "result", "value")
End Function

'エレメントの中心位置のxy座標に左クリックイベント発行
Private Sub IWebElement_DispatchLeftClickEvent()
    Dim res As String: res = CallObjectIdUsingMethod("GetBoxModel")
    Dim centerXY As Variant: centerXY = GetCenterPositionXY(res)
    Dim x As Long: x = centerXY(0)
    Dim y As Long: y = centerXY(1)
    
    CDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, 0, y - 50
    
    Dim Result As String
    Result = CDP_.DispatchMouseEvent("mousePressed", x, y, "left", 1)
    Result = CDP_.DispatchMouseEvent("mouseReleased", x, y, "left", 1)
End Sub
'エレメントの中心位置のxy座標を取得
Private Function GetCenterPositionXY(res As String)
    Dim Result As Variant
    Result = Json_.GetValue(res, "result", "model", "content")
    Result = Split(Result, ",")
    Dim width As Long, height As Long
    width = Json_.GetValue(res, "result", "model", "width")
    height = Json_.GetValue(res, "result", "model", "height")
    Dim xy(1) As Variant
    xy(0) = Result(0) + (width / 2)
    xy(1) = Result(1) + (height / 2)
    GetCenterPositionXY = xy
End Function

'Public
'type=fileのInputBoxへファイル設定
Private Sub IWebElement_SetFileToFileSelectInputButton(filePath As String)
'    filePath = Replace(filePath, "\", "\\")
'    filePath = Json_.EncodeURL(filePath)
'    filePath = Replace(filePath, "%5C", "\")
'    chrome_.SetFileInputFiles objectId_,
'CDPMethodで設定しようとしても、ファイルパスに日本語が含まれる場合、
'エンコードして送信はできるが、クローム側でデコードしてくれない為、没
    Dim res As String: res = CallObjectIdUsingMethod("GetBoxModel")
    Dim centerXY As Variant: centerXY = GetCenterPositionXY(res)
    
    CDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, 0, CLng(centerXY(1)) - 50
    
    IWebElement_DispatchLeftClickEvent
    Sleep 1000
    ControlFileSelectDialog filePath
    Sleep 1000
End Sub
'ファイル選択ダイアログ操作
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
'クリック後に制御を返さないアラートダイアログ出現時に使用するアラート処理メソッド
Private Sub IWebElement_ClickAndThenAlertDialogErase()
    CallObjectIdUsingMethod "Focus"
    Dim res As String
    res = CallFunction("function(){this.click();}", "")
    Sleep 1000
    
    CDP_.HandleAlertDialog
End Sub

'Public
'elementの属性値を取得（属性が存在しない場合は””を返却）
Private Function IWebElement_GetAttribute(AttributeName As String) As String
    Dim res As String
    res = CallFunction("function(name){return this.getAttribute(name);}", AttributeName)
    If InStr(res, """value"":null") > 0 Then
        IWebElement_GetAttribute = ""
    Else
        IWebElement_GetAttribute = Json_.GetValue(res, "result", "result", "value")
    End If
End Function

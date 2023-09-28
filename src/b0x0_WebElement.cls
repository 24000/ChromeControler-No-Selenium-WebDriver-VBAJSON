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

Private browser_ As a3_BrowserInfo
Private CDP_ As a4_ExecuteCDP
Private events_ As a5_CDPEventNotifier
Private Helper_ As a6_ExecuteHelperFunction

Private objectId_ As String
Private findInfo_ As Collection
Private nodeId_ As Long
Private attributes_ As Object
Private frameId_ As String

Private xpath_ As String
Private name_


Public Sub Init(browser As a3_BrowserInfo, _
                       ExecCDP As a4_ExecuteCDP, _
                       events As a5_CDPEventNotifier, _
                       execHelper As a6_ExecuteHelperFunction, _
                       objectId As String, _
                       findInfo As Collection, _
                       Optional attributes As Object = Nothing, _
                       Optional frameId As String)
                       
    Set browser_ = browser
    Set CDP_ = ExecCDP
    Set events_ = events
    Set Helper_ = execHelper
    
    objectId_ = objectId
    Set findInfo_ = findInfo
    Set attributes_ = attributes
    frameId_ = frameId
End Sub

Public Property Get objectId() As String
    objectId = objectId_
End Property
Public Property Get frameId() As String
    frameId = frameId_
End Property

'CallJavaScriptFunction�ďo���̋��ʊ֐�
'�߂�l��"Cannot find context with specified id"�ƂȂ�ꍇ�AObjectId�̍Ď擾���K�v�Ȃ��ߋ��ʉ�
Private Function CallFunction(functionCode As String, args As String) As String
    Dim res As String
    res = CDP_.CallJavaScriptFunction(functionCode, objectId_, args)
    If InStr(res, "Cannot find context with specified id") > 0 Then
        browser_.UpdateExecutionContextId CDP_
        ReGetObjectId
        res = CDP_.CallJavaScriptFunction(functionCode, objectId_, args)
    End If
    CallFunction = res
End Function
Public Function ReGetObjectId()
    Dim Codes As Variant
    Select Case findInfo_(1)
      Case "id"
        Codes = Helper_.GetIdCodes(findInfo_(2))
      Case "xpath"
        Codes = Helper_.GetXpathCodes(findInfo_(2))
      Case "name"
        Codes = Helper_.GetNameCodes(findInfo_(2))
      Case "tag"
        Codes = Helper_.GetTagCodes(findInfo_(2), findInfo_(3))
    End Select
    objectId_ = CDP_.GetObjectIdOfElementBy(findInfo_(2), browser_, Codes)
End Function

'CDPMethod���ďo���ꍇ�ɁA������ObjectId���g�p���Čďo��Method�̋��ʊǗ��̂��߂̊֐�
'�߂�l��"Cannot find context with specified id"�ƂȂ�ꍇ�ɁAObjectId�̍Ď擾���K�v�ɂȂ邽�ߋ��ʉ�
'����Focus��GetGetBoxModel�̌Ăяo���̂�
Private Function CallObjectIdUsingMethod(methodName As String) As String
    Dim res As String
    res = CallByName(CDP_, methodName, VbMethod, objectId_)

    If InStr(res, "Cannot find context with specified id") > 0 Then
        browser_.UpdateExecutionContextId CDP_
        ReGetObjectId
        res = CallByName(CDP_, methodName, VbMethod, objectId_)
    End If
    CallObjectIdUsingMethod = res
End Function

'***********************************************
'Public
'�`�F�b�N�{�b�N�X�v�f�ɂ���Checked�����̐ݒ�
Private Property Let IWebElement_Checked(RHS As Boolean)
    Dim res As String
    'checked�v���p�e�B�������Ă��邩���O�Ƀ`�F�b�N�B�����Ă��Ȃ��ꍇ�͂����B
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
    "<WebElement�FChecked�v���p�e�B>" & vbCrLf & _
    "����Element��Checked�v���p�e�B�������܂���B" & vbCrLf & _
    "��Checked��CheckBox�Ŏg�p�\" & vbCrLf
End Property
'Public
'�`�F�b�N�{�b�N�X�v�f�ɂ��ă`�F�b�N��Ԏ擾
Private Property Get IWebElement_Checked() As Boolean
    Dim res As String
    res = CallFunction("function(){return this.checked;}", "")
    If InStr(res, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim Checked As Boolean
    Checked = xJson.GetValue(res, "result", "result", "value")
    IWebElement_Checked = Checked
Exit Property

quit:
Helper_.DisplayErrorAndEnd "����Element��Checked�v���p�e�B�������܂���B" & vbCrLf & _
                                            "��Checked��CheckBox�Ŏg�p�\" & vbCrLf
End Property











'Public
'�v�f�̏��(�L��or����)���擾
Private Property Get IWebElement_IsEnabled() As Boolean
    Dim res As String
    res = CallFunction("function(){return this.disabled;}", "")
    If InStr(res, """type"":""boolean""") = 0 Then GoTo quit
    
    Dim Enabled As Boolean
    Enabled = xJson.GetValue(res, "result", "result", "value")
    IWebElement_IsEnabled = Not Enabled
Exit Property
    
quit:
Helper_.DisplayErrorAndEnd "����Element��Enabled�v���p�e�B�������܂���B" & vbCrLf
End Property



'Public
'javaScript�ɂ���ăN���b�N���s
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
    CDP_.PageEnable '��񖳊Q�ȃ��\�b�h�𓊂��āA�C�x���g�����󋵍X�V
    
    If browser_.GetPageStatus(IsTargetNavigated) _
    Or browser_.GetPageStatus(IsStartedLoading) Then
        IsPageChanging = True
    Else
        IsPageChanging = False
    End If
End Function
Private Sub WaitPageChangeToComplete()
    
    'TargetNavigated�͕Ԃ��ė��Ȃ��P�[�X������H�̂ŌJ��Ԃ��񐔂ɐ����������ă��[�v
    Dim failSafeCounter As Long
    Dim isEndChange As Boolean
    Do
        isEndChange = browser_.GetPageStatus(IsTargetNavigated)
'                                Or�@chrome_.CDPEvents.IsCurrentTargetClearedScheduledNavigation
        If isEndChange Then Exit Do
        
        CDP_.RuntimeEnable
        Sleep 1000
        failSafeCounter = failSafeCounter + 1
    Loop While (isEndChange = False And failSafeCounter <= 4)
    
    browser_.SetPageStatus IsScheduledNavigation, False
    browser_.SetPageStatus IsStartedLoading, False
    browser_.SetPageStatus IsScheduledNavigation, False
    browser_.SetPageStatus IsTargetNavigated, False
End Sub
Private Function IsIsolatedFrameChanging() As Boolean
    CDP_.DOMEnable '��񖳊Q�ȃ��\�b�h(frame�Ȃ̂�DOM�n)�𓊂��āA�C�x���g�����󋵍X�V
    
    
    If browser_.GetIframeStatus(IsScheduledNavigation) _
    Or browser_.GetIframeStatus(IsStartedLoading) Then
        IsIsolatedFrameChanging = True
    Else
        IsIsolatedFrameChanging = False
    End If
End Function
Private Sub WaitIsolatedFrameChangeToComplete()
    
    'FrameNavigated�͕Ԃ��ė��Ȃ��P�[�X������H�̂ŌJ��Ԃ��񐔂ɐ����������ă��[�v
    Dim failSafeCounter As Long
    Dim isEndChange As Boolean
    Do
        isEndChange = browser_.GetIframeStatus(IsTargetNavigated)
'                               Or chrome_.CDPEvents.IsIsolatedFrameClearedScheduledNavigation
        If isEndChange Then Exit Do
        
        CDP_.DOMEnable
        CDP_.RuntimeEnable
        Sleep 1000
        failSafeCounter = failSafeCounter + 1
    Loop While (isEndChange = False And failSafeCounter <= 4)
    
    browser_.SetIframeStatus IsScheduledNavigation, False
    browser_.SetIframeStatus IsStartedLoading, False
    browser_.SetIframeStatus IsClearedScheduledNavigation, False
    browser_.SetIframeStatus IsTargetNavigated, False

End Sub





'Public
'�v�f��Value�����ɕ������ݒ�
Private Sub IWebElement_SetText(text As String)
    FocusAndEnterKey
    Dim encodingText As String: encodingText = xJson.EncodeURL(text)
    CallFunction "function(value){value = decodeURI(value);this.value = value;}", encodingText
    
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText
End Sub

'Public
'�v�f��TextContent�ɕ������ݒ�
Private Sub IWebElement_SetTextContent(text As String)
    FocusAndEnterKey
    Dim encodingText As String: encodingText = xJson.EncodeURL(text)
    CallFunction "function(value){value = decodeURI(value);this.textContent  = value;}", encodingText
    
    CallFunction Helper_.GetDispatchChangeEventCode, encodingText
End Sub

'Public
'SelectBox�̑I�����𕶎���őI��
Private Sub IWebElement_SelectItemInSelectBoxByText(text As String)
    Dim encodingText As String: encodingText = xJson.EncodeURL(text)
    
    Dim res As String: res = CallFunction(Helper_.GetSelectItemCode, encodingText)
    
    Dim IsSelect As Boolean: IsSelect = xJson.GetValue(res, "result", "result", "value")
    If IsSelect = False Then
        MsgBox "<SelectItemInSelectBoxByText>" & vbCrLf & _
                    "�u" & text & "�v�͑I�����̒��ɑ��݂��܂���B" & vbCrLf & "�������I�����܂��B"
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
    
    Dim objectId As String
    objectId = xJson.GetValue(res, "result", "result", "objectId")
    If objectId = "" Then GoTo quit
    
    Dim findInfoInElem As Collection
    Set findInfoInElem = Helper_.MakeFindInfo("name", name)
    
    Set IWebElement_FindElementByName = _
        CDP_.MakeElementObject(objectId, findInfoInElem, browser_, events_, Helper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>������Ȃ������v�f" & vbCrLf & name
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
Debug.Print vbCrLf & ">>������Ȃ������v�f" & vbCrLf & tagName
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Tag", tagName, "����")
End Function
Private Function GetElementsCount(tagName As String) As Long
    Dim funcCode As String, res As String
    funcCode = "function(tagName){return  this.getElementsByTagName('" & tagName & "').length;}"
    res = CallFunction(funcCode, tagName)
    GetElementsCount = xJson.GetValue(res, "result", "result", "value")
End Function
Private Function CreateElementsItemByIndex(tagName As String, Index As Long) As IWebElement
    Dim funcCode As String, res As String
    funcCode = "function(tagName){return  this.getElementsByTagName(tagName)[" & Index & "];}"
    res = CallFunction(funcCode, tagName)
    
    Dim objectId As String
    objectId = xJson.GetValue(res, "result", "result", "objectId")
    Dim findInfoInElem As Collection: Set findInfoInElem = Helper_.MakeFindInfo("tag", tagName, Index)

    Set CreateElementsItemByIndex = _
        CDP_.MakeElementObject(objectId, findInfoInElem, browser_, events_, Helper_)
End Function

'Public
'�G�������g��TextContent���擾
Private Function IWebElement_GetTextContent() As String
    Dim funcCode As String, res As String
    funcCode = "function(){return this.textContent.replace(/\\s/g, ' ');}"
    res = CallFunction(funcCode, "")
    
    IWebElement_GetTextContent = xJson.GetValue(res, "result", "result", "value")
End Function

'Public
'�Z���N�g�{�b�N�X�̑I�𒆃I�v�V������Value���擾
Private Function IWebElement_GetSelectedValue() As String
    Dim funcCode As String, res As String
    funcCode = "function(){return this.value.replace(/\\s/g, ' ');}"
    res = CallFunction(funcCode, "")

    IWebElement_GetSelectedValue = xJson.GetValue(res, "result", "result", "value")
End Function

'Public
'�Z���N�g�{�b�N�X�̑I�𒆃I�v�V������TextContent���擾
Private Function IWebElement_GetSelectedTextContent() As String
    Dim funcCode As String, res As String
    funcCode = Helper_.GetSelectedTextContentCode
    res = CallFunction(funcCode, "")

    IWebElement_GetSelectedTextContent = xJson.GetValue(res, "result", "result", "value")
End Function

'�G�������g�̒��S�ʒu��xy���W�ɍ��N���b�N�C�x���g���s
Private Sub IWebElement_DispatchLeftClickEvent()
    Dim res As String: res = CallObjectIdUsingMethod("GetBoxModel")
    Dim centerXY As Variant: centerXY = GetCenterPositionXY(res)
    Dim x As Long: x = centerXY(0)
    Dim y As Long: y = centerXY(1)
    
    CDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, 0, y - 50
    
    Dim result As String
    result = CDP_.DispatchMouseEvent("mousePressed", x, y, "left", 1)
    result = CDP_.DispatchMouseEvent("mouseReleased", x, y, "left", 1)
End Sub
'�G�������g�̒��S�ʒu��xy���W���擾
Private Function GetCenterPositionXY(res As String)
    Dim result As Variant
    result = xJson.GetValue(res, "result", "model", "content")
    result = Split(result, ",")
    Dim width As Long, height As Long
    width = xJson.GetValue(res, "result", "model", "width")
    height = xJson.GetValue(res, "result", "model", "height")
    Dim xy(1) As Variant
    xy(0) = result(0) + (width / 2)
    xy(1) = result(1) + (height / 2)
    GetCenterPositionXY = xy
End Function

'Public
'type=file��InputBox�փt�@�C���ݒ�
Private Sub IWebElement_SetFileToFileSelectInputButton(filePath As String)
'    filePath = Replace(filePath, "\", "\\")
'    filePath = xJson.EncodeURL(filePath)
'    filePath = Replace(filePath, "%5C", "\")
'    chrome_.SetFileInputFiles objectId_,
'CDPMethod�Őݒ肵�悤�Ƃ��Ă��A�t�@�C���p�X�ɓ��{�ꂪ�܂܂��ꍇ�A
'�G���R�[�h���đ��M�͂ł��邪�A�N���[�����Ńf�R�[�h���Ă���Ȃ��ׁA�v
    Dim res As String: res = CallObjectIdUsingMethod("GetBoxModel")
    Dim centerXY As Variant: centerXY = GetCenterPositionXY(res)
    
    CDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, 0, CLng(centerXY(1)) - 50
    
    IWebElement_DispatchLeftClickEvent
    Sleep 1000
    ControlFileSelectDialog filePath
    Sleep 1000
End Sub
'�t�@�C���I���_�C�A���O����
Private Sub ControlFileSelectDialog(filePath As String)
    Dim DialogHwnd As LongPtr: DialogHwnd = FindWindow(vbNullString, "�J��")
    Dim InputHwnd As Long
    InputHwnd = FindWindowEx(DialogHwnd, 0&, "ComboBoxEx32", "")
    InputHwnd = FindWindowEx(InputHwnd, 0&, "ComboBox", "")
    InputHwnd = FindWindowEx(InputHwnd, 0&, "Edit", "")
    Call SendMessage(InputHwnd, &HC, 0, ByVal filePath)
 
    Dim ButtonHwnd As Long
    ButtonHwnd = FindWindowEx(DialogHwnd, 0&, "Button", "�J��(&O)")
    Call SendMessage(ButtonHwnd, &HF5, 0, 0&)
End Sub

'Public
'�N���b�N��ɐ����Ԃ��Ȃ��A���[�g�_�C�A���O�o�����Ɏg�p����A���[�g�������\�b�h
Private Sub IWebElement_ClickAndThenAlertDialogErase()
    CallObjectIdUsingMethod "Focus"
    Dim res As String
    res = CallFunction("function(){this.click();}", "")
    Sleep 1000
    
    CDP_.HandleAlertDialog
End Sub

'Public
'element�̑����l���擾�i���������݂��Ȃ��ꍇ�́h�h��ԋp�j
Private Function IWebElement_GetAttribute(AttributeName As String) As String
    Dim res As String
    res = CallFunction("function(name){return this.getAttribute(name);}", AttributeName)
    If InStr(res, """value"":null") > 0 Then
        IWebElement_GetAttribute = ""
    Else
        IWebElement_GetAttribute = xJson.GetValue(res, "result", "result", "value")
    End If
End Function
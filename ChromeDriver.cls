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

Private Const CHROME_EXE As String = """chrome.exe"""
Private Const DEBUG_MODE_OPTION As String = " --remote-debugging-Port=9222"
Private Const USER_DATA_DIR_OPTION As String = " --user-data-dir="
Private Const ENABLE_AUTOMATION_OPTION As String = " --enable-automation"

Private userDataDir As String

Private webSocket_ As a_0_x1_WebSocketCommunicator
Private CDPMethod_ As a_0_x2_CDPMethods
Private jsonHandler_ As a_0_x4_JSONHandler
Private basicInfo_ As a_0_x3_BasicInfos
Private events_ As a_0_x1_x2_CDPEventHandler

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

Public Property Get WebSocket() As a_0_x1_WebSocketCommunicator
    Set WebSocket = webSocket_
End Property
Public Property Get CDPMethod() As a_0_x2_CDPMethods
    Set CDPMethod = CDPMethod_
End Property
Public Property Get jsonhandler() As a_0_x4_JSONHandler
    Set jsonhandler = jsonHandler_
End Property
Public Property Get CDPEvents() As a_0_x1_x2_CDPEventHandler
    Set CDPEvents = events_
End Property


Private Sub Class_Initialize()
    
    userDataDir = ThisWorkbook.Path & "\�}�N�����s�ō쐬�����t�H���__�폜OK\"
    If Dir(userDataDir, vbDirectory) = "" Then
        MkDir userDataDir
    End If
    
    CreateObject("WScript.Shell").Run _
        CHROME_EXE & DEBUG_MODE_OPTION & ENABLE_AUTOMATION_OPTION _
                              & USER_DATA_DIR_OPTION & "" & userDataDir & ""
    Sleep 2000
    
    Set jsonHandler_ = New a_0_x4_JSONHandler
    Set basicInfo_ = New a_0_x3_BasicInfos: basicInfo_.Init jsonhandler
    Set webSocket_ = New a_0_x1_WebSocketCommunicator
    If webSocket_.Init(basicInfo_.Host, basicInfo_.WebSocketUrlPath, basicInfo_.DebuggingPort) = False Then End
    
    Set CDPMethod_ = New a_0_x2_CDPMethods
    Set events_ = New a_0_x1_x2_CDPEventHandler
    
    SetDiscoverTargets
End Sub

'message�������M����M
'��M���ɂ̓N���[������C�x���g�������ʒm�����ꍇ������̂ŁAid�ő��M�ɑ΂����M�𔻒f
'���C�x���g�ʒm�ɂ�id���܂܂�Ȃ�
Public Function ProcessMessage( _
    sendId As Long, message As String, jsonhandler As a_0_x4_JSONHandler) As String

    If webSocket_.SendMessage(message) = False Then GoTo quit

    Dim responseJson As String, responseId As Variant
    Do
        responseJson = webSocket_.GetMessage
        If responseJson = "��M�G���[����" Then GoTo quit
            
        responseId = jsonhandler.GetValueFromJson(responseJson, Array("id"))
        If responseId = "" Then
            If InStr(responseJson, "method") > 0 Then
                events_.GetInfo responseJson, Me
            End If
        End If
    Loop While (sendId <> responseId)
    
    ProcessMessage = responseJson
Exit Function

quit:
Dim errMsg As String
If responseJson = "" Then
    errMsg = "���M�G���[�����F�C�~�f�B�G�C�g�Q��"
Else
    errMsg = "��M�G���[����:�C�~�f�B�G�C�g�Q��"
End If
MsgBox errMsg, vbCritical, "�G���["
webSocket_.Http.CloseHInternetHandles
End
End Function

'Browser�n************************************************
Private Function SessionClose() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.SessionClose
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

'************************************************

'ElementId�n************************************************
'CDPMethod�ɂȂ��̂�Javascript���s
Private Function GetDocumentObjectID() As String
    Dim responseJson  As String
    responseJson = ExecuteJavaScript_Inner("document;")
    documentObjectId_ = jsonHandler_.GetValueFromJson(responseJson, _
                                                        Array("result", "result", "objectId"))
    
    GetDocumentObjectID = documentObjectId_
End Function

Private Function GetDocumentNodeID() As Long
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.GetDocumentID
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    GetDocumentNodeID = jsonHandler_.GetValueFromJson(responseJson, Array("result", "root", "nodeId"))
End Function

Private Function ChangeObjectIdToNodeId(objectId As String) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.RequestNode(objectId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    
    Dim nodeId As String
    nodeId = jsonHandler_.GetValueFromJson(responseJson, Array("result", "nodeId"))
    ChangeObjectIdToNodeId = nodeId
End Function

Private Function ChangeNodeIdToObjectId(nodeId As Long) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.ResolveNode(nodeId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    
    Dim objectId As String
    objectId = jsonHandler_.GetValueFromJson(responseJson, Array("result", "objectId"))
    ChangeNodeIdToObjectId = objectId
End Function
'************************************************

'Element���n************************************************
Private Function GetAttributes(nodeId As Long) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.GetAttributes(nodeId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    GetAttributes = responseJson
End Function

Private Function DescribeNode(IdName As String, id As Variant) As Variant
    If IdName <> "nodeId" And IdName <> "objectId" Then
        Err.Raise 10000, , "�������ɂ�""nodeId""��""objectId""�̂ǂ��炩�w�肵�Ă��������B"
    End If
    
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.DescribeNodeBy_nodeId_or_objectId(IdName, id)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    Dim attributes As String
    attributes = jsonHandler_.GetValueFromJson(responseJson, Array("result", "node", "attributes"))
    DescribeNode = attributes
End Function
'************************************************

'Runtime************************************************
'executionContextCreated �C�x���g�ɂ����s�R���e�L�X�g�����̕񍐂�L���ɂ��܂��B
'���|�[�g���L���ɂȂ�ƁA�C�x���g�͊����̎��s�R���e�L�X�g���Ƃɒ����ɑ��M����܂��B
Public Sub RuntimeEnable()
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.RunTime_Enable()
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Sub
'JavaScript�R�[�h���s
Private Function ExecuteJavaScript_Inner(script As String) As String
    Dim jsonForSending  As String
    jsonForSending = CDPMethod_.JavaScritpEvaluate(script)
    ExecuteJavaScript_Inner = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

'JavaScript_Function���s
Public Function JavaScritpCallFunction(funcName As String, objectId As Variant, args As Variant) As String
    Dim jsonForSending  As String
    jsonForSending = CDPMethod_.JavaScritpCallFunction(funcName, objectId, args)
    JavaScritpCallFunction = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function
'************************************************
'Input Domain*******************************************************************
'�L�[�C�x���g���y�[�W�Ƀf�B�X�p�b�`����BkeyDown , keyUp, rawKeyDown, char�̂����ꂩ�B
Public Function DispatchKeyEvent_DOMKeyCode(eventType As String, code As String) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.DispatchKeyEvent_DOMKeyCode(eventType, code)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

'*******************************************************************

'DOM����n************************************************
Public Sub DOMEnable()
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.DOM_Enable()
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Sub

Public Sub Focus(objectId As String)
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.Focus(objectId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Sub

Public Sub FindElementByQuerySelector(Selector As String)
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.querySelector(Selector)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    
    Dim nodeId As Long
    nodeId = jsonHandler_.GetValueFromJson(responseJson, Array("result", "nodeId"))
    Dim nodeInfoJson As String
    nodeInfoJson = DescribeNode("nodeId", nodeId)
End Sub

Public Function FindElementByID(id As String) As String
    
    
End Function
'************************************************

'Input�n************************************************
Public Function Click(objId As String) As String
    JavaScritpCallFunction "function(value){this.click();}", objId, ""
End Function
'************************************************

'Target:Page�n************************************************
Public Function PageEnable() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.PageEnable()
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

Public Function GetFrameTree() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.GetFrameTree()
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    GetFrameTree = responseJson
End Function

Public Function CreateIsolatedFrameWorld(frameId As String) As Long
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.CreateIsolatedFrameWorld(frameId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    CreateIsolatedFrameWorld = _
        jsonHandler_.GetValueFromJson(responseJson, Array("result", "executionContextId"))
End Function

Public Function SetDiscoverTargets() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.SetDiscoverTargets(basicInfo_.targetId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

Public Function AttachToTarget(targetId As String) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.AttachToTarget(targetId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

Public Function ActivateTarget(targetId As String) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.ActivateTarget(targetId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

Public Function DetachFromTarget(targetId As String) As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.DetachFromTarget(targetId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function

Public Function TargetClose() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.TargetClose(basicInfo_.targetId)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function
Public Function GetTargetInfo() As String
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.GetTargets()
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
End Function





















'�ȉ���C���^�[�t�F�C�X���J���\�b�h����т��̊֘A���\�b�h
'************************************************
'Public
'���݂̃y�[�W��URL
Private Property Get IWebDriver_CurrentURL() As String
    SetDiscoverTargets  '������g����Page���ŐV��
    Dim i As Long
    For i = 1 To events_.PageInfos.Count
        If events_.PageInfos(i)("targetId") = basicInfo_.targetId Then
            IWebDriver_CurrentURL = events_.PageInfos(i)("url")
            Exit Property
        End If
    Next
End Property


'Public
'��1�����Ŏw���URL���J��
Private Sub IWebDriver_URLOpen(url As String, Optional maxWaitTime_ms As Long = 30000)
    Dim jsonForSending  As String, responseJson As String
    jsonForSending = CDPMethod_.URLOpen(url)
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    If InStr(responseJson, "invalid URL") > 0 Then GoTo quit
    
    WaitingForWebPageDisplay (maxWaitTime_ms)
    
    SetDiscoverTargets
    PageEnable
    DOMEnable
    documentObjectId_ = GetDocumentObjectID
    documentNodeId_ = GetDocumentNodeID
Exit Sub

quit:
DisplayErrorAndEnd "�w�肳�ꂽ�ȉ�URL�͖�����URL�ł��B" & vbCrLf & url
End Sub

'�y�[�W�\���҂����[�v
Private Function WaitingForWebPageDisplay(maxWaitTime_ms As Long)
    Dim javaScriptCode As String, responseJson As String, state As String
    Do
        javaScriptCode = "document.readyState;"
        responseJson = ExecuteJavaScript_Inner(javaScriptCode)
        state = jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "value"))
        DoEvents
        Sleep 800
    Loop Until (state = "complete")
End Function

'Public
'��1������JavaScript�R�[�h�̎��s
Private Function IWebDriver_ExecuteJavaScript(script As String) As String
    Dim responseJson As String
    responseJson = ExecuteJavaScript_Inner(script)
    IWebDriver_ExecuteJavaScript = jsonHandler_.GetValueFromJson(responseJson, Array("result"))
End Function

'Public
'��1�����Ŏw�肳���Xpath�̗v�f���擾���A�G�������gObject�Ƃ��ĕԋp
Private Function IWebDriver_GetElementByXpath(xPath As String) As IWebElement
    ReCreateIsolatedFrameWorld_IfFlagOn
    
    Dim scriptCode As String, functionCode As String
    scriptCode = "document.evaluate('" & xPath & "', document, null, " & _
                        "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;"
    functionCode = "function(xpath){return  document.evaluate(xpath, document, null, " & _
                            "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;}"
            
    Dim objectId As String
    objectId = GetObjectIdOfElementBy(xPath, scriptCode, functionCode)
    If objectId = "" Then GoTo quit
    Set IWebDriver_GetElementByXpath = MakeElementObject(objectId)
Exit Function

quit:
DisplayErrorAndEnd "�w�肳�ꂽ�ȉ�Xpath�̗v�f�͑��݂��܂���ł����B" & vbCrLf & xPath
End Function
'Xpath�ŗv�f��T���A�I�u�W�F�N�gID��Ԃ�
Private Function GetObjectIdOfElementBy(way As String, scriptCode As String, functionCode As String) As String
    Dim responseJson As String
    If executionContextId_ = 0 Then
        responseJson = ExecuteJavaScript_Inner(scriptCode)
    Else
        responseJson = JavaScritpCallFunction(functionCode, executionContextId_, way)
    End If
    
    GetObjectIdOfElementBy = _
        jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "objectId"))
End Function
'�G�������g���쐬���A�߂�l�Ƃ��ĕԋp
Public Function MakeElementObject(objectId As String) As IWebElement
'    DOMEnable
    Dim nodeId As Long
'    nodeId = ChangeObjectIdToNodeId(objectId)
'    Dim attributes As String
'    attributes = GetAttributes(nodeId)
    
    Dim element As WebElement: Set element = New WebElement
    element.Init Me, objectId, nodeId
    Set MakeElementObject = element
End Function

'Public
Private Function IWebDriver_GetElementByName(Name As String) As IWebElement
    ReCreateIsolatedFrameWorld_IfFlagOn
    
    Dim scriptCode As String, functionCode As String
    scriptCode = "document.getElementsByName('" & Name & "')[0];"
    functionCode = "function(name){return  document.getElementsByName(name)[0];}"
    
    Dim objectId As String
    objectId = GetObjectIdOfElementBy(Name, scriptCode, functionCode)
    If objectId = "" Then GoTo quit
    Set IWebDriver_GetElementByName = MakeElementObject(objectId)
Exit Function

quit:
DisplayErrorAndEnd "�w�肳�ꂽ�ȉ�Name�̗v�f�͑��݂��܂���ł����B" & vbCrLf & Name
End Function

'Public
Private Function IWebDriver_GetElementsByTagName(tagName As String) As IWebElements
    
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
    
    Set IWebDriver_GetElementsByTagName = elems
Exit Function

quit:
DisplayErrorAndEnd "�w�肳�ꂽ�ȉ�TagName�̗v�f�͑��݂��܂���ł����B" & vbCrLf & tagName
End Function

Private Function GetElementsCount(tagName As String) As Long
    Dim javaScriptCode As String, responseJson As String
    If executionContextId_ = 0 Then
        javaScriptCode = "document.getElementsByTagName('" & tagName & "').length;"
        responseJson = ExecuteJavaScript_Inner(javaScriptCode)
    Else
        javaScriptCode = _
        "function(tagName){return  document.getElementsByTagName('" & tagName & "').length;}"
        responseJson = JavaScritpCallFunction(javaScriptCode, executionContextId_, tagName)
    End If
    GetElementsCount = jsonHandler_.GetValueFromJson(responseJson, Array("result", "result", "value"))
End Function

Private Function CreateElementsItemByIndex(tagName As String, Index As Long) As IWebElement
    Dim javaScriptCode As String, responseJson As String
    Dim objectId As String
    Dim elem As IWebElement
    
    If executionContextId_ = 0 Then
        javaScriptCode = "document.getElementsByTagName('" & tagName & "')[" & Index & "];"
        responseJson = ExecuteJavaScript_Inner(javaScriptCode)
    Else
        javaScriptCode = _
        "function(tagName){return  document.getElementsByTagName(tagName)[" & Index & "];}"
        responseJson = JavaScritpCallFunction(javaScriptCode, executionContextId_, tagName)
    End If
    objectId = jsonhandler.GetValueFromJson(responseJson, Array("result", "result", "objectId"))
    Set CreateElementsItemByIndex = MakeElementObject(objectId)
End Function


'Public
'��1�����Ŏw�肳���l�̃t���[���Ƀt���[����ؑ�
Private Sub IWebDriver_SwitchFrameByNameOrUrlOrIndex(value As Variant)
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim frameId As String
    If IsNumeric(value) Then
        If frameInfos.Count < value Then GoTo quit
        frameId = jsonHandler_.GetValueFromJson(frameInfos(value), Array("id"))
    Else
        frameId = GetTargetFrameId(frameInfos, value)
    End If
    If frameId = "" Then GoTo quit
    executionContextId_ = CreateIsolatedFrameWorld(frameId)
    isolatedFrameId_ = frameId
    
Exit Sub

quit:
DisplayErrorAndEnd "�w�肳�ꂽframe�͑��݂��܂���B" & vbCrLf & value & vbCrLf & vbCrLf & _
                              "������Url�w��̏ꍇ�AChrome������Ԃ��Ȃ��P�[�X������܂��B" & vbCrLf & _
                              "   �Ƃ�Ȃ��悤�ł����Index�A�܂��͂����Name�w��𐄏����܂��B"
End Sub

'frame�c���[���擾���A�e�t���[���̏��iJSON�I�u�W�F�N�g������j���R���N�V�����Ƃ��ďW��
'frameTree��result��frameTree
'��frame(Page�̐e�t���[��)��ChildFrames(�z��̒���Page�̎qframe���(�I�u�W�F�N�g)�j�̍\��
Private Function GetFrameInfos() As Collection
    Dim collect As Collection: Set collect = New Collection
    Dim json As String: json = GetFrameTree
'    collect.add jsonHandler_.GetObjectFromJson(json, Array("result", "frameTree", "frame"))
    
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        '�܂�childFrames������Frame���i�I�u�W�F�N�g�j�擾
        frmInfo = jsonHandler_.GetObjectFromJson(json, _
                        Array("result", "frameTree", "childFrames", Count))
        If frmInfo = "" Then Exit Do
        '���̃I�u�W�F�N�g�̒���frame�v���p�e�B(�I�u�W�F�N�g)���擾���A�R���N�V�����ɉ�����
        frminfo2 = jsonHandler_.GetObjectFromJson(frmInfo, Array("frame"))
        collect.add frminfo2
        'childFrames�v���p�e�B������ꍇ�A�ċA�\���Ȃ̂ōċA�֐��œ��l�̏�������
        If InStr(frmInfo, "childFrames") > 0 Then
            GetChildFrameInfo collect, frmInfo
        End If
        Count = Count + 1
    Loop
    Set GetFrameInfos = collect
End Function
'childFrame������ꍇ�A�ċA����
Public Sub GetChildFrameInfo(collect As Collection, json As String)
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        frmInfo = jsonHandler_.GetObjectFromJson(json, _
                        Array("childFrames", Count))
        If frmInfo = "" Then Exit Do
        frminfo2 = jsonHandler_.GetObjectFromJson(frmInfo, Array("frame"))
        collect.add frminfo2
        If InStr(frmInfo, "childFrames") > 0 Then
            GetChildFrameInfo collect, frmInfo
        End If
        Count = Count + 1
    Loop
End Sub
'�e�t���[����name,url�Ǝw�肳�ꂽ�l���r�B��v����t���[����Id��Ԃ��B
Private Function GetTargetFrameId(frameInfos As Collection, value As Variant) As String
    Dim frameInfo As Variant
    Dim Name As String, url As String
    For Each frameInfo In frameInfos
        Name = jsonHandler_.GetValueFromJson(frameInfo, Array("name"))
        url = jsonHandler_.GetValueFromJson(frameInfo, Array("url"))
        If Name = value Or url = value Then
            GetTargetFrameId = jsonHandler_.GetValueFromJson(frameInfo, Array("id"))
            Exit Function
        End If
    Next
End Function

'Public
'�؂�ւ��\�ȃt���[���̏����C�~�f�B�G�C�g�E�B���h�E�ɕ\��
Private Sub IWebDriver_DebugPrintAccessibleFrameInfo()
    PageEnable
    RuntimeEnable
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim i As Long
    For i = 1 To frameInfos.Count
        Debug.Print "index�F" & i & _
                           "�@name�F" & jsonHandler_.GetValueFromJson(frameInfos(i), Array("name")) & _
                           "�@url�F" & jsonHandler_.GetValueFromJson(frameInfos(i), Array("url"))
    Next
End Sub

'Public
'�f�t�H���g�̃t���[���ɖ߂�
Private Sub IWebDriver_SwitchFrameToDefault()
    isolatedFrameId_ = ""
    executionContextId_ = 0
End Sub

'Public
'��1�����Ŏw�肳���Xpath�̗v�f�L���̊m�F
Private Function IWebDriver_IsExistsElementByXpath(xPath As String) As Boolean
    ReCreateIsolatedFrameWorld_IfFlagOn
    
    Dim scriptCode As String, functionCode As String
    scriptCode = "document.evaluate('" & xPath & "', document, null, " & _
                        "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;"
    functionCode = "function(xpath){return  document.evaluate(xpath, document, null, " & _
                            "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;}"
    Dim objectId As String
    objectId = GetObjectIdOfElementBy(xPath, scriptCode, functionCode)
    If objectId = "" Then
        IWebDriver_IsExistsElementByXpath = False
    Else
        IWebDriver_IsExistsElementByXpath = True
    End If
End Function

'Public
'Xpath(��1����)�Ŏw�肳���̗v�f���o������̂��A�ő�ҋ@����(��2����)�҂�
Private Sub IWebDriver_WaitForElementToAppearByXpath(xPath As String, maxWaitTime_ms As Long)
    
    Dim startTime As Date: startTime = Timer()
    Do
        If IWebDriver_IsExistsElementByXpath(xPath) Then
            Exit Sub
        ElseIf CInt((Timer() - startTime) * 1000) > maxWaitTime_ms Then
            DisplayErrorAndEnd maxWaitTime_ms & _
                                         "�~���b�o�߂��Ă��A�w��̈ȉ��v�f�͌���܂���ł����B" & vbCrLf & _
                                          xPath
        End If
        ReCreateIsolatedFrameWorld_IfFlagOn
        DoEvents: Sleep 500
    Loop
End Sub

'Public
'Alert�_�C�A���O��OK�{�^���������ĕ���
Private Sub IWebDriver_PushOkButtonOnDialog()
    Dim jsonForSending  As String, responseJson As String
    '�C�x���g��M���邽�߂Ɉ�񖳊Q�ȃ��\�b�h�𑗎�M����B
    jsonForSending = CDPMethod_.GetTargets
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    
    jsonForSending = CDPMethod_.handleJavaScriptDialog()
    responseJson = ProcessMessage(CDPMethod_.id, jsonForSending, jsonHandler_)
    If InStr(responseJson, "No dialog is showing") > 0 Then GoTo quit
Exit Sub

quit:
    DisplayErrorAndEnd "�_�C�A���O���\������Ă��܂���" & vbCrLf & _
                                   "���\�������^�C�~���O�������K�v��������܂���B"
End Sub

Private Sub IWebDriver_SleepByWindowsAPI(milliSecond As Long)
    Sleep milliSecond
End Sub

'���ʃG���[�Ή��B���b�Z�[�W�\���A���C���^�[�l�b�g�n���h����������ċ����I���Ƃ���
'���G���[�������Ƀf�o�b�O���[�h�ɓ���A���[�U�[�ɔC����Ɖ������Ȃ��\���������̂ŋ����I���Ƃ���B
Public Sub DisplayErrorAndEnd(displayText As String)
    displayText = displayText & vbCrLf & "�������I�����܂��B"
    MsgBox displayText, vbCritical, "�G���["
    webSocket_.Http.CloseHInternetHandles
    End
End Sub

'���ʑΉ��B
'IsolatedWorld���Ă���t���[���ɍX�V������ꍇ�A�C�x���g�Ō��m���ăt���O�����̂ŁA�ēxIsolatedWorld���鏈���B
'�y�g�p���Ă��郁�\�b�h�z
'IWebDriver_GetElementByXpath
'IWebDriver_WaitForElementToAppearByXpath
'IWebDriver_IsExistsElementByXpath
'IWebDriver_GetElementsByTagName
Public Sub ReCreateIsolatedFrameWorld_IfFlagOn()
'event���m�����܂������Ȃ����߁A��UDefaultFrame�ȊO�̑���p�����ɂ͖���CreateIsolatedFrameWorld����
    If isolatedFrameId_ <> "" Then
        GetFrameTree '��U���Q�ȃ��\�b�h�𓊂��ăC�x���g�����`�F�b�N
        executionContextId_ = CreateIsolatedFrameWorld(isolatedFrameId_)
    End If
    
'    DOMEnable '�O�̂��߈�񖳊Q�ȃ��\�b�h�𓊂��ăt���[���X�V�C�x���g�̔����`�F�b�N�B�����Ȃ����I
'    If events_.IsIsolatedFrameScheduledNavigation = False Then Exit Sub
'
'    '�X�V���I���܂Ń��[�v
'    If events_.IsIsolatedFrameScheduledNavigationEnd = False Then
'        Dim Counter As Long
'        Do
'            Sleep 1000
'            DOMEnable
'            Counter = Counter + 1
'            If Counter > 30 Then GoTo quit '�ő��30�b�҂��Ă��_���������狭���I��
'        Loop Until events_.IsIsolatedFrameScheduledNavigationEnd
'    End If
'
'    executionContextId_ = CreateIsolatedFrameWorld(isolatedFrameId_)
    events_.IsIsolatedFrameScheduledNavigation = False
    events_.IsIsolatedFrameScheduledNavigationEnd = False
Exit Sub

'quit:
'DisplayErrorAndEnd "�t���[���̍X�V�����m���܂������A30�b�o�߂��Ă��X�V���I���܂���ł����B"
End Sub

'�E�ׂ̃^�u�ֈړ�
Private Sub IWebDriver_SwitchTabByIndex(Index As Long)
    Dim currentTargetId As String
    currentTargetId = events_.PageInfos(Index)("targetId")
    basicInfo_.Update currentTargetId, jsonHandler_
    
    webSocket_.CloseWebSocket
    Set webSocket_ = Nothing
    Set webSocket_ = New a_0_x1_WebSocketCommunicator
    If webSocket_.Init(basicInfo_.Host, basicInfo_.WebSocketUrlPath, basicInfo_.DebuggingPort) = False Then End
    
    ActivateTarget basicInfo_.targetId
    PageEnable
    DOMEnable
    
    documentObjectId_ = GetDocumentObjectID
    documentNodeId_ = GetDocumentNodeID
End Sub

Private Sub IWebDriver_CloseWindow()
    SessionClose
End Sub
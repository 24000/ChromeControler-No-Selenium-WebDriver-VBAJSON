VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a4_ExecuteCDP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private MsgGenerate_ As a4x1_CDPMessageGenerator
Private MsgHandler_ As a4x2_CDPMessageHandler


Public Sub Init(MsgGenerate As a4x1_CDPMessageGenerator, _
                       Handler As a4x2_CDPMessageHandler)
                       
    Set MsgGenerate_ = MsgGenerate
    Set MsgHandler_ = Handler
End Sub

'Browser�n************************************************
'�u���E�U�����
Public Function SessionClose() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.SessionClose
    SessionClose = MsgHandler_.SendAndReceive(sendMsg)
End Function

'************************************************

'DOM(ElementId)�n************************************************
'CDPMethod�ɂȂ��̂�Javascript���s
Public Function GetDocumentObjectID() As String
    Dim resMsg  As String
    resMsg = ExecuteJavaScript("document;")
    GetDocumentObjectID = xJson.GetValue(resMsg, "result", "result", "objectId")
End Function

Public Function GetDocumentNodeID() As Long
    Dim sendMsg  As String, resMsg As String
    sendMsg = MsgGenerate_.GetDocumentID
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    GetDocumentNodeID = xJson.GetValue(resMsg, "result", "root", "nodeId")
End Function

Private Function ChangeObjectIdToNodeId(objectId As String) As String
    Dim sendMsg  As String, resMsg As String
    sendMsg = MsgGenerate_.RequestNode(objectId)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    ChangeObjectIdToNodeId = xJson.GetValue(resMsg, "result", "nodeId")
End Function

Private Function ChangeNodeIdToObjectId(nodeId As Long) As String
    Dim sendMsg  As String, resMsg As String
    sendMsg = MsgGenerate_.ResolveNode(nodeId)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    ChangeNodeIdToObjectId = xJson.GetValue(resMsg, "result", "objectId")
End Function
'************************************************

'DOM(Element���)�n************************************************
Private Function GetAttributes(nodeId As Long) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.GetAttributes(nodeId)
    GetAttributes = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function DescribeNode(id As Variant) As Variant
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.DescribeNode(id)
    DescribeNode = MsgHandler_.SendAndReceive(sendMsg)
End Function
'************************************************

'Runtime************************************************
'executionContextCreated �C�x���g�ɂ����s�R���e�L�X�g�����C�x���g�̕񍐂�L������B
'���L���ɂ������T�ŁA�C�x���g�͊����̎��s�R���e�L�X�g���Ƃɒ����ɑ��M�����B
Public Function RuntimeEnable() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.RunTime_Enable()
    RuntimeEnable = MsgHandler_.SendAndReceive(sendMsg)
End Function
'JavaScript�R�[�h���s
Public Function ExecuteJavaScript(script As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.JavaScritpEvaluate(script)
    ExecuteJavaScript = MsgHandler_.SendAndReceive(sendMsg)
End Function

'JavaScript_Function���s
Public Function CallJavaScriptFunction(funcName As String, id As Variant, args As Variant) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.CallJavaScriptFunction(funcName, id, args)
    CallJavaScriptFunction = MsgHandler_.SendAndReceive(sendMsg)
End Function

'************************************************
'Input Domain*******************************************************************
'�L�[�C�x���g���y�[�W�Ƀf�B�X�p�b�`����BkeyDown , keyUp, rawKeyDown, char�̂����ꂩ�B
Public Function DispatchKeyEvent_DOMKeyCode(eventType As String, code As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.DispatchKeyEvent_DOMKeyCode(eventType, code)
    DispatchKeyEvent_DOMKeyCode = MsgHandler_.SendAndReceive(sendMsg)
End Function

'�}�E�X�C�x���g��xy���W�Ƀf�B�X�p�b�`����BmousePressed, mouseReleased, mouseMoved, mouseWheel�̂����ꂩ�B
Public Function DispatchMouseEvent(eventType As String, x As Long, y As Long, _
                                                        button As String, clickCount As Long, _
                                                        Optional modifiers As Long = 0, _
                                                        Optional deltaX As Long = 0, _
                                                        Optional deltaY As Long = 0) As String
                                                        
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.DispatchMouseEvent( _
                        eventType, x, y, button, clickCount, _
                        modifiers, deltaX, deltaY)
    DispatchMouseEvent = MsgHandler_.SendAndReceive(sendMsg)
End Function

'*******************************************************************

'DOM����n************************************************
Public Function DOMEnable() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.DOM_Enable()
    DOMEnable = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function Focus(objectId As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.Focus(objectId)
    Focus = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function GetBoxModel(objectId As String) As Variant
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.GetBoxModel(objectId)
    GetBoxModel = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function SetFileInputFiles(objectId As Variant, files As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.SetFileInputFiles(objectId, files)
    SetFileInputFiles = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Sub FindElementByQuerySelector(Selector As String)
    Dim sendMsg  As String, resMsg As String
    sendMsg = MsgGenerate_.querySelector(Selector)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    
    Dim nodeId As Long
    nodeId = xJson.GetValue(resMsg, "result", "nodeId")
    Dim nodeInfoJson As String
    nodeInfoJson = DescribeNode(nodeId)
End Sub

'************************************************

'Input�n************************************************
Public Function Click(objId As String) As String
    CallJavaScriptFunction "function(value){this.click();}", objId, ""
End Function
'************************************************

'Page�n************************************************
Public Function OpenURL(targetURL As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.URLOpen(targetURL)
    OpenURL = MsgHandler_.SendAndReceive(sendMsg)
End Function

'CDP�ɂ�����PageDomain�̃C�x���g�ʒm��L��������
Public Function PageEnable() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.PageEnable()
    PageEnable = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function GetFrameTree() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.GetFrameTree()
    GetFrameTree = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function CreateIsolatedFrameWorld(frameId As String) As Long
    Dim sendMsg  As String, resMsg As String
    sendMsg = MsgGenerate_.CreateIsolatedFrameWorld(frameId)
    resMsg = MsgHandler_.SendAndReceive(sendMsg)
    If InStr(resMsg, "error") > 0 Then
        CreateIsolatedFrameWorld = 0
    Else
        CreateIsolatedFrameWorld = xJson.GetValue(resMsg, "result", "executionContextId")
        Sleep 200
    End If
End Function

'Target�n************************************************
'targetCreated/targetInfoChanged/targetDestroyed �C�x���g�̒ʒm��L����
Public Function SetDiscoverTargets() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.SetDiscoverTargets()
    SetDiscoverTargets = MsgHandler_.SendAndReceive(sendMsg)
End Function

'�w�肳�ꂽ id �̃^�[�Q�b�g�ɃA�^�b�`����B�߂�l��JSON����SessionID���܂܂��B
Public Function AttachToTarget(targetId As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.AttachToTarget(targetId)
    Dim res As String
    res = MsgHandler_.SendAndReceive(sendMsg)
    Dim sessionId As String
    sessionId = xJson.GetValue(res, "result", "sessionId")
    AttachToTarget = sessionId
End Function

'�^�[�Q�b�g���������i�t�H�[�J�X�j����B
Public Function ActivateTarget(targetId As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.ActivateTarget(targetId)
    ActivateTarget = MsgHandler_.SendAndReceive(sendMsg)
End Function

'�w�肳�ꂽ ID �̃Z�b�V�������폜�B
Public Function DetachFromTarget(targetId As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.DetachFromTarget(targetId)
    DetachFromTarget = MsgHandler_.SendAndReceive(sendMsg)
End Function

'�^�[�Q�b�g�����B�^�[�Q�b�g���y�[�W�̏ꍇ�A���̃y�[�W��������B
Public Function TargetClose(targetId As String) As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.TargetClose(targetId)
    TargetClose = MsgHandler_.SendAndReceive(sendMsg)
End Function
Public Function GetTargetInfo() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.GetTargets()
    GetTargetInfo = MsgHandler_.SendAndReceive(sendMsg)
End Function

Public Function HandleAlertDialog() As String
    Dim sendMsg  As String
    sendMsg = MsgGenerate_.HandleJavaScriptDialog()
    HandleAlertDialog = MsgHandler_.SendAndReceive(sendMsg)
End Function

    

'����Function�n************************************************

'�y�[�W�\���҂����[�v
Public Function WaitingForWebPageDisplay()
    Dim javaScriptCode As String, responseMsg As String, state As String
    Do
        Sleep 800
        javaScriptCode = "document.readyState;"
        responseMsg = ExecuteJavaScript(javaScriptCode)
        state = xJson.GetValue(responseMsg, "result", "result", "value")
        DoEvents
    Loop Until (state = "complete")
End Function



'�w��̕��@�ŗv�f��T��(�ő�S�b�ԒT���j�A�I�u�W�F�N�gID��Ԃ�
Public Function GetObjectIdOfElementBy(way As String, basicInfo As a3_BasicInfos, Codes As Variant) As String
    
    Dim startTime As Date: startTime = Timer()
    Dim res As String, objectId As String
    Do
        If basicInfo.ExecutionContextId = 0 Then
            res = ExecuteJavaScript((Codes(0)))
        Else
            res = CallJavascriptFunctionWrapper((Codes(1)), basicInfo, way)
        End If

        objectId = xJson.GetValue(res, "result", "result", "objectId")
    
        If objectId <> "" Then Exit Do
        
        If CInt((Timer() - startTime)) > 4 Then Exit Do
        
        DoEvents:  Sleep 500
    Loop
    
    GetObjectIdOfElementBy = objectId
End Function

'CallJavaScriptFunction���Ăԏꍇ�̃��b�p�\�֐��B�G���[�Ή����K�{�̂���
Public Function CallJavascriptFunctionWrapper(funcCode As String, basicInfo As a3_BasicInfos, args As Variant) As String
    Dim res As String
    res = CallJavaScriptFunction(funcCode, basicInfo.ExecutionContextId, args)
    If InStr(res, "Cannot find context with specified id") > 0 Then
        basicInfo.UpdateExecutionContextId Me
        res = CallJavaScriptFunction(funcCode, basicInfo.ExecutionContextId, args)
    End If
    CallJavascriptFunctionWrapper = res
End Function



'�G�������g���쐬���A�߂�l�Ƃ��ĕԋp
'�G�������g���e�FObjectId,NodeId,Attributes(����),�t���[���̏ꍇ�̂�frameId�ǉ�
Public Function MakeElementObject( _
        objectId As String, _
        findInfo As Collection, _
        basicInfo As a3_BasicInfos, _
        events As a5_CDPEventNotifier, _
        Helper As a6_ExecuteHelperFunction) As IWebElement
    
    Dim info  As String
    info = DescribeNode(objectId)
    If InStr(info, "attributes") > 0 Then
        Dim attributes As Object
        Set attributes = MakeAttributesDic(xJson.GetValue(info, "result", "node", "attributes"))
    End If
    If InStr(info, "frameId") > 0 Then
        Dim frameId As String
        frameId = xJson.GetValue(info, "result", "node", "frameId")
    End If
    
    Dim element As b0x0_WebElement: Set element = New b0x0_WebElement
    element.Init basicInfo, Me, events, Helper, _
                       objectId, findInfo, attributes, frameId
    Set MakeElementObject = element
End Function
'element�̑������������ɂ��ĕԂ��B
Private Function MakeAttributesDic(attributes As String) As Object
    Dim ary As Variant: ary = Split(attributes, ",")
    Dim attributesDic As Object: Set attributesDic = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(ary) To UBound(ary) - 1 Step 2
        If attributesDic.Exists(ary(i)) = False Then
            attributesDic.add ary(i), ary(i + 1)
        End If
    Next
    Set MakeAttributesDic = attributesDic
End Function


Public Function GetElementsCountBy(way As String, basicInfo As a3_BasicInfos, Codes As Variant) As Long
    Dim responseMsg As String
    If basicInfo.ExecutionContextId = 0 Then
        responseMsg = ExecuteJavaScript((Codes(0)))
    Else
        responseMsg = CallJavascriptFunctionWrapper((Codes(1)), basicInfo, way)
    End If
    GetElementsCountBy = xJson.GetValue(responseMsg, "result", "result", "value")
End Function


'IsExists���ʊ֐�
Public Function IsExistsElementBy(way As String, basicInfo As a3_BasicInfos, Codes As Variant) As Boolean

    Dim responseMsg As String
    If basicInfo.ExecutionContextId = 0 Then
        responseMsg = ExecuteJavaScript((Codes(0)))
    Else
        responseMsg = CallJavascriptFunctionWrapper((Codes(1)), basicInfo, way)
    End If
    Dim objectId As String
    objectId = xJson.GetValue(responseMsg, "result", "result", "objectId")

    If objectId = "" Then
        IsExistsElementBy = False
    Else
        IsExistsElementBy = True
    End If
End Function

'WaitForElement���ʊ֐�
Public Function WaitForElementToAppearBy( _
    way As String, maxWaitTime_sec As Long, basicInfo As a3_BasicInfos, Codes As Variant) As Boolean
    
    Dim startTime As Date: startTime = Timer()
    Do
        If IsExistsElementBy(way, basicInfo, Codes) Then
            WaitForElementToAppearBy = True
            Exit Function
        ElseIf CInt((Timer() - startTime)) > maxWaitTime_sec Then
            WaitForElementToAppearBy = False
            Exit Function
        End If
        DoEvents: Sleep 500
    Loop
End Function
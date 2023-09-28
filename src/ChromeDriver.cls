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
Private Const CAPTION_NAME As String = "�V�����^�u - Google Chrome" '"data:, - Google Chrome"

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "USER32.DLL" ( _
        ByVal hWnd As LongPtr, ByRef ProcessId As Long) As Long

Private Declare PtrSafe Sub SetForeGroundwindow Lib "user32" Alias "SetForegroundWindow" (ByVal ms As LongPtr)




Private http_ As a1x1_HTTPCommunicator
Private webSocket_ As a1_WebSocketCommunicator
Private basicInfo_ As a3_BasicInfos
Private event_ As a5_CDPEventNotifier
Private browser_ As a3_BrowserInfo
Private CDP_ As a4_ExecuteCDP
Private Helper_ As a6_ExecuteHelperFunction

Private pageNum_ As Long
Private sessionId_ As String

Private processId_ As Long
Private chromeHwnd_ As LongPtr

Private Sub Class_Initialize()
    Dim userDataDir As String
    userDataDir = CreateObject("Shell.Application").Namespace("shell:Downloads").Self.path & "\�}�N�����s��pUserData�t�H���_ForChrome\"
    
    Dim isFirst As Boolean
    If Dir(userDataDir, vbDirectory) = "" Then
        MkDir userDataDir
        isFirst = True
    End If
    
    processId_ = StartChrome(userDataDir)
    New_InnerObjects
    
    If isFirst Then  '���̈ʒu�Ŕ���K�v�iCloseWindow���邽�ߒʐM�m����j
        CDP_.SessionClose
        ChromeFirstLaunch
    End If
    
'    CDP_.SetDiscoverTargets
    pageNum_ = 1
'    sessionId_ = CDP_.AttachToTarget(browser_.targetId(pageNum_))
    
    WaitForActivation
    '�K�؂Ɏ擾���ł��Ȃ��P�[�X����A�v���P�i����Ă܂��ǂ��ł�hwnd�͎g�p���Ȃ��j
    chromeHwnd_ = FindWindow(CLASS_NAME, CAPTION_NAME)
    
End Sub

Private Function StartChrome(userDataDir As String) As Long

    Dim chromeExePath As String
    Dim objExec As Object
    With CreateObject("WScript.Shell")
        chromeExePath = """" & .RegRead(xG.CHROME_EXE_REG_KEY) & """"
        Set objExec = .Exec( _
                            chromeExePath & xG.DEBUG_MODE_OPTION & xG.ENABLE_AUTOMATION_OPTION _
                           & xG.HOME_PAGE & xG.USER_DATA_DIR_OPTION & "" & userDataDir & "")
        StartChrome = objExec.ProcessId
        Set objExec = Nothing
    End With

End Function

'�u���E�U�N���҂���SetForeGroundWindow�ɂ��g�p
Private Sub WaitForActivation()
On Error Resume Next
    
    Dim counter As Long
recheck:
    AppActivate processId_
    If Err.Number <> 0 Then
        counter = counter + 1
        If counter > 20 Then
            MsgBox "10�b�o�߂��܂������u���E�U���L���ɂȂ�܂���ł����B" & vbCrLf & _
                        "�������I�����܂�"
            End
        End If
        Debug.Print Err.Number & ":" & Err.Description
        Helper_.SleepByWinAPI 500
        Err.Number = 0
        GoTo recheck
    End If
        
On Error GoTo 0
End Sub

'�����I�u�W�F�N�g�̐���
Private Sub New_InnerObjects()
    Set basicInfo_ = New a3_BasicInfos
    
    Set http_ = New a1x1_HTTPCommunicator
    
    'BrowserInfo����
    Dim json As String: json = http_.GetTargetList()
    Set event_ = New a5_CDPEventNotifier
    Set browser_ = New a3_BrowserInfo: browser_.Init json, event_
    
    Set webSocket_ = New a1_WebSocketCommunicator
    If webSocket_.Init(http_, browser_.WebSocketPath(1)) = False Then End
    
    
'    ExecuteCDP_�����i�ˑ��������j
    Dim msgGenerator As a4x1_CDPMessageGenerator
    Set msgGenerator = New a4x1_CDPMessageGenerator
    Dim Handler As a4x2_CDPMessageHandler
    Set Handler = New a4x2_CDPMessageHandler
    Handler.Init webSocket_, event_
    Set CDP_ = New a4_ExecuteCDP
    CDP_.Init msgGenerator, Handler
    
    Set Helper_ = New a6_ExecuteHelperFunction
End Sub




'�N���[���̐^�̏���N�����̉�ʑJ��
Private Sub ChromeFirstLaunch()
    MsgBox "����A����N���̂��߁ADownLoad�t�H���_����" & vbCrLf & _
                "�}�N�����s��p��UserData�t�H���_���쐬���܂����B" & vbCrLf & vbCrLf & _
                "���萔�ł����A�ēx�}�N�������s���Ă��������B" & vbCrLf & _
                "������ȍ~�A���̃��b�Z�[�W�͕\������܂���B" & vbCrLf & _
                "���쐬���ꂽ�t�H���_�͍폜���Ă���肠��܂��񂪁A" & vbCrLf & _
                "   �}�N���Ď��s���ɂ͍č쐬����܂��B", _
                vbInformation, "Chrome����N�����b�Z�[�W"
    End
End Sub












'�ȉ���C���^�[�t�F�C�X���J���\�b�h
'************************************************
'Public
'���݂̃y�[�W��URL
Private Property Get IWebDriver_URL() As String
    CDP_.SetDiscoverTargets  '������g����Page���ŐV��
    IWebDriver_URL = browser_.URL
'    Dim i As Long
'    For i = 1 To browser_.PageInfosCount
'        If event_.pageInfos(i)("targetId") = basicInfo_.TargetId Then
'            IWebDriver_URL = events_.pageInfos(i)("url")
'            Exit Property
'        End If
'    Next
End Property


'Public
'��1�����Ŏw���URL���J��
Private Sub IWebDriver_OpenURL(targetURL As String)
    Dim res As String
    res = CDP_.OpenURL(targetURL)
    If InStr(res, "invalid URL") > 0 Then
        Helper_.DisplayErrorAndEnd _
            "��OpenURL��" & vbCrLf & _
            "�w�肳�ꂽ�ȉ�URL�͖�����URL�ł��B" & vbCrLf & _
            targetURL & vbCrLf
    End If
    
    CDP_.WaitingForWebPageDisplay
    
    CDP_.SetDiscoverTargets
    CDP_.PageEnable
'    CDP_.DOMEnable
    basicInfo_.DocumentObjectId = CDP_.GetDocumentObjectID
    basicInfo_.DocumentNodeId = CDP_.GetDocumentNodeID
End Sub

'Public
'��1������JavaScript�R�[�h�̎��s
Private Function IWebDriver_ExecuteScript(script As String) As String
    Dim responseMsg As String
    responseMsg = CDP_.ExecuteJavaScript(script)
    IWebDriver_ExecuteScript = xJson.GetValue(responseMsg, "result")
End Function

'Public
'��1�����Ŏw�肳���Xpath�̗v�f���擾���A�G�������gObject�Ƃ��ĕԋp
Private Function IWebDriver_FindElementByXpath(xPath As String) As IWebElement
    
    If InStr(xPath, """") > 0 Then xPath = Replace(xPath, """", "\""")
    
    Dim Codes As Variant: Codes = Helper_.GetXpathCodes(xPath)
    Dim objectId As String: objectId = CDP_.GetObjectIdOfElementBy(xPath, basicInfo_, Codes)
    If objectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("xpath", xPath)
    Set IWebDriver_FindElementByXpath = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, event_, Helper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>������Ȃ������v�f" & vbCrLf & xPath
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("XPath", xPath)
End Function

'Public
'��1�����Ŏw�肳���Id�̗v�f���擾���A�G�������gObject�Ƃ��ĕԋp
Private Function IWebDriver_FindElementById(id As String) As IWebElement
    Dim Codes As Variant: Codes = Helper_.GetIdCodes(id)
    Dim objectId As String: objectId = CDP_.GetObjectIdOfElementBy(id, basicInfo_, Codes)
    If objectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("id", id)
    Set IWebDriver_FindElementById = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, event_, Helper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>������Ȃ������v�f" & vbCrLf & id
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Id", id)
End Function

'Public
'��1�����Ŏw�肳���Name�̗v�f���擾���A�G�������gObject�Ƃ��ĕԋp
Private Function IWebDriver_FindElementByName(name As String) As IWebElement

    Dim Codes As Variant: Codes = Helper_.GetNameCodes(name)
    Dim objectId As String: objectId = CDP_.GetObjectIdOfElementBy(name, basicInfo_, Codes)
    If objectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("name", name)
    Set IWebDriver_FindElementByName = _
        CDP_.MakeElementObject(objectId, findInfo, basicInfo_, event_, Helper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>������Ȃ������v�f" & vbCrLf & name
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Name", name)
End Function


'Public
'tag���ŗv�f�R���N�V�������擾�i�v�f�O���[�v��Length���擾���ALength���v�f�쐬���[�v�j
Private Function IWebDriver_FindElementsByTag(tagName As String) As IWebElements
    Dim countCodes As Variant: countCodes = Helper_.GetTagCountCodes(tagName)
    Dim elemsCount As Long: elemsCount = CDP_.GetElementsCountBy(tagName, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    
    Dim objectId As String
    Dim Codes As Variant
    Dim Index As Long
    For Index = 0 To elemsCount - 1
        Codes = Helper_.GetTagCodes(tagName, Index)
        objectId = CDP_.GetObjectIdOfElementBy(tagName, basicInfo_, Codes)
        Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("tag", tagName, Index)
        elements.add _
            CDP_.MakeElementObject(objectId, findInfo, basicInfo_, event_, Helper_)
    Next
    
    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements
    
    Set IWebDriver_FindElementsByTag = elems
Exit Function

quit:
Debug.Print vbCrLf & ">>������Ȃ������v�f" & vbCrLf & tagName
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("Tag", tagName, "����")
End Function

'Public
'class���ŗv�f�R���N�V�������擾�i�v�f�O���[�v��Length���擾���ALength���v�f�쐬���[�v�j
Private Function IWebDriver_FindElementsByClassName(className As String) As IWebElements
    Dim countCodes As Variant: countCodes = Helper_.GetClassNameCountCodes(className)
    Dim elemsCount As Long: elemsCount = CDP_.GetElementsCountBy(className, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    
    Dim objectId As String
    Dim Codes As Variant
    Dim Index As Long
    For Index = 0 To elemsCount - 1
        Codes = Helper_.GetClassNameCodes(className, Index)
        objectId = CDP_.GetObjectIdOfElementBy(className, basicInfo_, Codes)
        Dim findInfo As Collection: Set findInfo = Helper_.MakeFindInfo("className", className, Index)
        elements.add _
            CDP_.MakeElementObject(objectId, findInfo, basicInfo_, event_, Helper_)
    Next
    
    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements
    
    Set IWebDriver_FindElementsByClassName = elems
Exit Function

quit:
Debug.Print vbCrLf & ">>������Ȃ������v�f" & vbCrLf & className
Beep
Helper_.DisplayErrorAndEnd Helper_.MakeFindErrorMessage("className", className, "����")
End Function


'Public
'��1�����Ŏw�肳���l�̃t���[���Ƀt���[����ؑ�
Private Sub IWebDriver_SwitchFrameByNameOrUrlOrIndex(value As Variant)
    Dim frameInfos As Collection
    Dim frameId As String
    Dim Count As Long
    
    Do 'frame�֘A�C�x���g�����^�C�~���O�ɂ���ď��擾�Ɏ��s����P�[�X����B�O�̂���3��܂Ń��[�v�B
        Set frameInfos = GetFrameInfos
        If IsNumeric(value) Then
            If frameInfos.Count < value Then GoTo quit
            frameId = xJson.GetValue(frameInfos(value), "id")
        Else
            frameId = GetTargetFrameId(frameInfos, value)
        End If
        Count = Count + 1
    Loop While frameId = "" And Count < 3
    
    If frameId = "" Then GoTo quit
    
    Dim ExecutionContextId As Long
    ExecutionContextId = CDP_.CreateIsolatedFrameWorld(frameId)
    basicInfo_.ExecutionContextId = ExecutionContextId
    basicInfo_.IsolatedFrameId = frameId
    
Exit Sub

quit:
Helper_.DisplayErrorAndEnd "��SwitchFrameByNameOrUrlOrIndex��" & vbCrLf & _
                                "�w�肳�ꂽ�ȉ���frame�͑��݂��܂���B" & vbCrLf & _
                                value & vbCrLf & vbCrLf & _
                                "��iframe��Name����������΁AName�ł̎w��𐄏����܂��B" & vbCrLf & vbCrLf & _
                                "���t���[����Index�ԍ��AName�AURL���m�F�������ꍇ�A" & vbCrLf & _
                                "�@ DebugPrintAccessibleFrameInfo���g�p���Ă��������B" & vbCrLf & _
                                "�@ �N���[������ʒm�����t���[����񂪃C�~�f�B�G�C�g�ɏo�͂���܂��B" & vbCrLf & vbCrLf & _
                                "�����̃��\�b�h�ł��܂������Ȃ��ꍇ�A" & vbCrLf & _
                                "�@ SwitchFrameByIframeElement��" & vbCrLf & _
                                "�@ �g�p���������Ă��������B" & vbCrLf
End Sub

'frame�c���[���擾���A�e�t���[���̏��iJSON�I�u�W�F�N�g������j���R���N�V�����Ƃ��ďW��
'frame�c���[��result��frameTree
'��frame(Page�̐e�t���[��)��ChildFrames(�z��̒���Page�̎qframe���(�I�u�W�F�N�g)�j�̍\��
Private Function GetFrameInfos() As Collection
    Dim collect As Collection: Set collect = New Collection
    Dim json As String: json = CDP_.GetFrameTree
'    collect.add xJson.GetObject(json, "result", "frameTree", "frame"))
    
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        '�܂�childFrames������Frame���i�I�u�W�F�N�g�j�擾
        frmInfo = xJson.GetObject(json, "result", "frameTree", "childFrames", Count)
        If frmInfo = "" Then Exit Do
        '���̃I�u�W�F�N�g�̒���frame�v���p�e�B(�I�u�W�F�N�g)���擾���A�R���N�V�����ɉ�����
        frminfo2 = xJson.GetObject(frmInfo, "frame")
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
Private Sub GetChildFrameInfo(collect As Collection, json As String)
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        frmInfo = xJson.GetObject(json, "childFrames", Count)
        If frmInfo = "" Then Exit Do
        frminfo2 = xJson.GetObject(frmInfo, "frame")
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
    Dim name As String, URL As String
    For Each frameInfo In frameInfos
        name = xJson.GetValue(frameInfo, "name")
        URL = xJson.GetValue(frameInfo, "url")
        If name = value Or URL = value Then
            GetTargetFrameId = xJson.GetValue(frameInfo, "id")
            Exit Function
        End If
    Next
End Function

'Public
'�؂�ւ��\�ȃt���[���̏����C�~�f�B�G�C�g�E�B���h�E�ɕ\��
Private Sub IWebDriver_DebugPrintAccessibleFrameInfo()
    CDP_.PageEnable
    CDP_.RuntimeEnable
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim i As Long
    For i = 1 To frameInfos.Count
        Debug.Print "index�F" & i & _
                           "�@name�F" & xJson.GetValue(frameInfos(i), "name") & _
                           "�@url�F" & xJson.GetValue(frameInfos(i), "url")
    Next
End Sub

'Public
Private Sub IWebDriver_SwitchFrameByIframeElement(IframeElement As IWebElement)
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim elem As b0x0_WebElement: Set elem = IframeElement
    If elem.frameId = "" Then
        Helper_.DisplayErrorAndEnd "��SwitchFrameByIframeElement��" & vbCrLf & _
                                      "�w�肳�ꂽElement��Iframe�ł͂���܂���" & vbCrLf & vbCrLf & _
                                      "Iframe�v�f�������ɓn���ĉ�����" & vbCrLf
    End If
    
    CDP_.GetTargetInfo
    CDP_.SetDiscoverTargets

    'frameElement��frameId��������frameId�Ȃ�IsolatedWorld�őΉ�
    Dim ExecutionContextId As Long
    ExecutionContextId = CDP_.CreateIsolatedFrameWorld(elem.frameId)
    If ExecutionContextId <> 0 Then
        basicInfo_.ExecutionContextId = ExecutionContextId
        basicInfo_.IsolatedFrameId = elem.frameId
        Exit Sub
    Else
        Helper_.DisplayErrorAndEnd "��SwitchFrameByIframeElement��" & vbCrLf & _
                                       "����Iframe�͂��̃��\�b�h�ł͐؂�ւ����ł��܂���B" & vbCrLf & vbCrLf & _
                                      "��SwitchFrameByNameOrUrlOrIndex��" & vbCrLf & _
                                      "�@ �g�p���������Ă��������B" & vbCrLf
    End If
    
    
'    �e�X�g
'    Dim docElem As b0x0_WebElement
'    Dim functionCode As String, responseMsg As String, docObjectId As String
'    functionCode = "function(name){return  this.contentDocument;}"
'    responseMsg = CallJavaScriptFunction(functionCode, elem.ObjectId, "")
'    docObjectId = xJson.GetValue(responseMsg, "result", "result", "objectId"))
'    Dim info  As String
'    info = DescribeNode(docObjectId)
    
    
'    frameId��targetId�Ȃ�ڑ��؂�ւ��őΉ�
'    basicInfo_.Update elem.FrameId, xJson, False
'    isolatedFrameId_ = ""
'
'    webSocket_.CloseWebSocket
'    Set webSocket_ = Nothing
'    Set webSocket_ = New a1_WebSocketCommunicator
'    If webSocket_.Init(basicInfo_.Host, basicInfo_.WebSocketUrlPath, _
'                                basicInfo_.DebuggingPort) = False Then
'        DisplayErrorAndEnd "SwitchIframe�Ɏ��s���܂���" & vbCrLf
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
'�f�t�H���g�̃t���[���ɖ߂�
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
'    basicInfo_.Update newTargetId, xJson, False
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
'��1�����Ŏw�肳���Xpath�̗v�f�L���̊m�F
Private Function IWebDriver_IsExistsElementByXpath(xPath As String) As Boolean
    IWebDriver_IsExistsElementByXpath = _
        CDP_.IsExistsElementBy(xPath, basicInfo_, Helper_.GetXpathCodes(xPath))
End Function
'Public
'��1�����Ŏw�肳���Id�̗v�f�L���̊m�F
Private Function IWebDriver_IsExistsElementById(id As String) As Boolean
    IWebDriver_IsExistsElementById = _
        CDP_.IsExistsElementBy(id, basicInfo_, Helper_.GetIdCodes(id))
End Function
'Public
'��1�����Ŏw�肳���Name�̗v�f�L���̊m�F
Private Function IWebDriver_IsExistsElementByName(name As String) As Boolean
    IWebDriver_IsExistsElementByName = _
        CDP_.IsExistsElementBy(name, basicInfo_, Helper_.GetNameCodes(name))
End Function


'Public
'Xpath(��1����)�Ŏw�肳���̗v�f���o������̂��A�ő�ҋ@����(��2����)�҂�
Private Sub IWebDriver_WaitForElementToAppearByXpath(xPath As String, maxWaitTime_sec As Long)
    If CDP_.WaitForElementToAppearBy(xPath, _
        maxWaitTime_sec, basicInfo_, Helper_.GetXpathCodes(xPath)) = False Then
        
        Helper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearByXPath>" & vbCrLf & _
             maxWaitTime_sec & "�b�o�߂��Ă��A�w��̈ȉ��v�f�͌���܂���ł����B" & vbCrLf & _
             xPath
    End If
End Sub
'Public
'Id(��1����)�Ŏw�肳���̗v�f���o������̂��A�ő�ҋ@����(��2����)�҂�
Private Sub IWebDriver_WaitForElementToAppearById(id As String, maxWaitTime_sec As Long)
    If CDP_.WaitForElementToAppearBy(id, _
        maxWaitTime_sec, basicInfo_, Helper_.GetIdCodes(id)) = False Then
        
        Helper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearById>" & vbCrLf & _
             maxWaitTime_sec & "�b�o�߂��Ă��A�w��̈ȉ��v�f�͌���܂���ł����B" & vbCrLf & _
             id
    End If
End Sub
'Public
'Name(��1����)�Ŏw�肳���̗v�f���o������̂��A�ő�ҋ@����(��2����)�҂�
Private Sub IWebDriver_WaitForElementToAppearByName(name As String, maxWaitTime_sec As Long)
    If CDP_.WaitForElementToAppearBy(name, _
        maxWaitTime_sec, basicInfo_, Helper_.GetNameCodes(name)) = False Then
    
        Helper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearByName>" & vbCrLf & _
             maxWaitTime_sec & "�b�o�߂��Ă��A�w��̈ȉ��v�f�͌���܂���ł����B" & vbCrLf & _
             name
    End If
End Sub


'Public
'Alert�_�C�A���O��OK�{�^���������ĕ���
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
                                   "�_�C�A���O���\������Ă��܂���B" & vbCrLf & _
                                   "�����̃��\�b�h��Alert���\�������ꍇ�Ɏg�p���܂��B" & vbCrLf
End Sub

Private Sub IWebDriver_SleepByWinAPI(milliSecond As Long)
    Helper_.SleepByWinAPI milliSecond
End Sub


'�^�u�̈ړ��iIndex�w��F��ԍ��̃^�u���P�Ƃ���j
Private Sub IWebDriver_SwitchTabByIndex(Index As Long)
    Dim newTargetId As String
    newTargetId = browser_.targetId(Index)
    basicInfo_.Update newTargetId, True
    
    webSocket_.CloseWebSocket
    Set webSocket_ = Nothing
    Set webSocket_ = New a1_WebSocketCommunicator
    If webSocket_.Init(http_, basicInfo_.WebSocketUrlPath) = False Then End
    
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

Private Sub IWebDriver_SetForeGroundWindow()
    WaitForActivation
End Sub

Private Sub IWebDriver_AllowPopUpWindowForURL(targetURL As String)
    Dim chromeExe As String
    Dim objExec As Object
    With CreateObject("WScript.Shell")
        .RegWrite name:="HKEY_LOCAL_MACHINE\SOFTWARE\Policies\{Google}", _
                        value:=targetURL, _
                        Type:="REG_SZ"
    End With
End Sub
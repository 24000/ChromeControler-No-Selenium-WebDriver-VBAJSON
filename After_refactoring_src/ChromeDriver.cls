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

'Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
'        ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
'

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "USER32.DLL" ( _
        ByVal hWnd As LongPtr, ByRef ProcessId As Long) As Long

Private Declare PtrSafe Sub SetForeGroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal ms As LongPtr)


Private webSocket_ As a1_WebSocketCommunicator
Private Json_ As a2_JSONHandler
Private basicInfo_ As a3_BasicInfos
Private execCDP_ As a4_ExecuteCDP
Private events_ As a5_CDPEventHandler
Private execHelper_ As a6_ExecuteHelperFunction

Private processId_ As Long
Private chromeHwnd_ As LongPtr

Private Sub Class_Initialize()
    Dim userDataDir As String
    userDataDir = ThisWorkbook.Path & "\�}�N�����s��pUserData�t�H���_\"
    
    Dim isFirst As Boolean
    If Dir(userDataDir, vbDirectory) = "" Then
        MkDir userDataDir
        isFirst = True
    End If
    
    processId_ = StartChrome(userDataDir)
    
    '�K�؂Ɏ擾���ł��Ȃ��P�[�X����A�v���P�i����Ă܂��ǂ��ł�hwnd�͎g�p���Ȃ��j
    chromeHwnd_ = FindWindow(CLASS_NAME, CAPTION_NAME)
    
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
    Set execCDP_ = New a4_ExecuteCDP
    execCDP_.Init msgGenerator, Handler, Json_
    
    Set execHelper_ = New a6_ExecuteHelperFunction
    
    If isFirst Then  '���̈ʒu�Ŕ���K�v�iCloseWindow���邽�ߒʐM�m����j
        execCDP_.SessionClose
        ChromeFirstLaunch
    End If
    
    execCDP_.SetDiscoverTargets
    execCDP_.AttachToTarget basicInfo_.targetId
    
    execHelper_.SleepByWinAPI 2000
End Sub

Private Function StartChrome(userDataDir As String) As Long
    Const CHROME_EXE_REG_KEY As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\"
    Const CHROME_EXE As String = """chrome.exe"""
    Const DEBUG_MODE_OPTION As String = " --remote-debugging-Port=9222"
    Const USER_DATA_DIR_OPTION As String = " --user-data-dir="
    Const ENABLE_AUTOMATION_OPTION As String = " --enable-automation"
    
    Dim chromeExe As String
    Dim objExec As Object
    With CreateObject("WScript.Shell")
        chromeExe = """" & .RegRead(CHROME_EXE_REG_KEY) & """"
        Set objExec = .Exec( _
                            chromeExe & DEBUG_MODE_OPTION & ENABLE_AUTOMATION_OPTION _
                            & USER_DATA_DIR_OPTION & "" & userDataDir & "")
        StartChrome = objExec.ProcessId
        Set objExec = Nothing
    End With
End Function

'�N���[���̐^�̏���N�����̉�ʑJ��
Private Sub ChromeFirstLaunch()
    MsgBox "����A����N���̂��߁A���̃t�@�C���Ɠ���t�H���_����" & vbCrLf & _
                "�}�N�����s��p��UserData�t�H���_���쐬���܂����B" & vbCrLf & vbCrLf & _
                "���萔�ł����A�ēx�}�N�������s���Ă��������B" & vbCrLf & _
                "������ȍ~�A���̃��b�Z�[�W�͕\������܂���B", _
                vbInformation, "Chrome����N�����b�Z�[�W"
    End
End Sub



'�ȉ���C���^�[�t�F�C�X���J���\�b�h
'************************************************
'Public
'���݂̃y�[�W��URL
Private Property Get IWebDriver_URL() As String
    execCDP_.SetDiscoverTargets  '������g����Page���ŐV��
    Dim i As Long
    For i = 1 To events_.PageInfos.Count
        If events_.PageInfos(i)("targetId") = basicInfo_.targetId Then
            IWebDriver_URL = events_.PageInfos(i)("url")
            Exit Property
        End If
    Next
End Property


'Public
'��1�����Ŏw���URL���J��
Private Sub IWebDriver_OpenURL(targetUrl As String)
    Dim res As String
    res = execCDP_.OpenURL(targetUrl)
    If InStr(res, "invalid URL") > 0 Then
        execHelper_.DisplayErrorAndEnd _
            "��OpenURL��" & vbCrLf & _
            "�w�肳�ꂽ�ȉ�URL�͖�����URL�ł��B" & vbCrLf & _
            targetUrl & vbCrLf
    End If
    
    execCDP_.WaitingForWebPageDisplay
    
    execCDP_.SetDiscoverTargets
    execCDP_.PageEnable
'    execCDP_.DOMEnable
    basicInfo_.DocumentObjectId = execCDP_.GetDocumentObjectID
    basicInfo_.DocumentNodeId = execCDP_.GetDocumentNodeID
End Sub

'Public
'��1������JavaScript�R�[�h�̎��s
Private Function IWebDriver_ExecuteScript(script As String) As String
    Dim responseMsg As String
    responseMsg = execCDP_.ExecuteJavaScript(script)
    IWebDriver_ExecuteScript = Json_.GetValue(responseMsg, "result")
End Function

'Public
'��1�����Ŏw�肳���Xpath�̗v�f���擾���A�G�������gObject�Ƃ��ĕԋp
Private Function IWebDriver_FindElementByXpath(xPath As String) As IWebElement
    
    If InStr(xPath, """") > 0 Then xPath = Replace(xPath, """", "\""")
    
    Dim codes As Variant: codes = execHelper_.GetXpathCodes(xPath)
    Dim ObjectId As String: ObjectId = execCDP_.GetObjectIdOfElementBy(xPath, basicInfo_, codes)
    If ObjectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = execHelper_.MakeFindInfo("xpath", xPath)
    Set IWebDriver_FindElementByXpath = _
        execCDP_.MakeElementObject(ObjectId, findInfo, basicInfo_, events_, execHelper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>������Ȃ������v�f" & vbCrLf & xPath
Beep
execHelper_.DisplayErrorAndEnd execHelper_.MakeFindErrorMessage("XPath", xPath)
End Function

'Public
'��1�����Ŏw�肳���Id�̗v�f���擾���A�G�������gObject�Ƃ��ĕԋp
Private Function IWebDriver_FindElementById(Id As String) As IWebElement
    Dim codes As Variant: codes = execHelper_.GetIdCodes(Id)
    Dim ObjectId As String: ObjectId = execCDP_.GetObjectIdOfElementBy(Id, basicInfo_, codes)
    If ObjectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = execHelper_.MakeFindInfo("id", Id)
    Set IWebDriver_FindElementById = _
        execCDP_.MakeElementObject(ObjectId, findInfo, basicInfo_, events_, execHelper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>������Ȃ������v�f" & vbCrLf & Id
Beep
execHelper_.DisplayErrorAndEnd execHelper_.MakeFindErrorMessage("Id", Id)
End Function

'Public
'��1�����Ŏw�肳���Name�̗v�f���擾���A�G�������gObject�Ƃ��ĕԋp
Private Function IWebDriver_FindElementByName(name As String) As IWebElement

    Dim codes As Variant: codes = execHelper_.GetNameCodes(name)
    Dim ObjectId As String: ObjectId = execCDP_.GetObjectIdOfElementBy(name, basicInfo_, codes)
    If ObjectId = "" Then GoTo quit
    Dim findInfo As Collection: Set findInfo = execHelper_.MakeFindInfo("name", name)
    Set IWebDriver_FindElementByName = _
        execCDP_.MakeElementObject(ObjectId, findInfo, basicInfo_, events_, execHelper_)
Exit Function

quit:
Debug.Print vbCrLf & ">>������Ȃ������v�f" & vbCrLf & name
Beep
execHelper_.DisplayErrorAndEnd execHelper_.MakeFindErrorMessage("Name", name)
End Function


'Public
'tag���ŗv�f�R���N�V�������擾�i�v�f�O���[�v��Length���擾���ALength���v�f�쐬���[�v�j
Private Function IWebDriver_FindElementsByTag(tagName As String) As IWebElements
    Dim countCodes As Variant: countCodes = execHelper_.GetTagCountCodes(tagName)
    Dim elemsCount As Long: elemsCount = execCDP_.GetElementsCountBy(tagName, basicInfo_, countCodes)
    If elemsCount = 0 Then GoTo quit
    
    Dim elements As Collection: Set elements = New Collection
    
    Dim ObjectId As String
    Dim codes As Variant
    Dim Index As Long
    For Index = 0 To elemsCount - 1
        codes = execHelper_.GetTagCodes(tagName, Index)
        ObjectId = execCDP_.GetObjectIdOfElementBy(tagName, basicInfo_, codes)
        Dim findInfo As Collection: Set findInfo = execHelper_.MakeFindInfo("tag", tagName, Index)
        elements.add _
            execCDP_.MakeElementObject(ObjectId, findInfo, basicInfo_, events_, execHelper_)
    Next
    
    Dim elems As c0x0_WebElements: Set elems = New c0x0_WebElements
    elems.Init elements
    
    Set IWebDriver_FindElementsByTag = elems
Exit Function

quit:
Debug.Print vbCrLf & ">>������Ȃ������v�f" & vbCrLf & tagName
Beep
execHelper_.DisplayErrorAndEnd execHelper_.MakeFindErrorMessage("Tag", tagName, "����")
End Function



'Public
'��1�����Ŏw�肳���l�̃t���[���Ƀt���[����ؑ�
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
    ExecutionContextId = execCDP_.CreateIsolatedFrameWorld(FrameId)
    basicInfo_.ExecutionContextId = ExecutionContextId
    basicInfo_.IsolatedFrameId = FrameId
    
Exit Sub

quit:
execHelper_.DisplayErrorAndEnd "��SwitchFrameByNameOrUrlOrIndex��" & vbCrLf & _
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
'frameTree��result��frameTree
'��frame(Page�̐e�t���[��)��ChildFrames(�z��̒���Page�̎qframe���(�I�u�W�F�N�g)�j�̍\��
Private Function GetFrameInfos() As Collection
    Dim collect As Collection: Set collect = New Collection
    Dim Json As String: Json = execCDP_.GetFrameTree
'    collect.add json_.GetObject(json, "result", "frameTree", "frame"))
    
    Dim frmInfo As String, frminfo2 As String
    Dim Count As Long
    Do
        '�܂�childFrames������Frame���i�I�u�W�F�N�g�j�擾
        frmInfo = Json_.GetObject(Json, "result", "frameTree", "childFrames", Count)
        If frmInfo = "" Then Exit Do
        '���̃I�u�W�F�N�g�̒���frame�v���p�e�B(�I�u�W�F�N�g)���擾���A�R���N�V�����ɉ�����
        frminfo2 = Json_.GetObject(frmInfo, "frame")
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
'�e�t���[����name,url�Ǝw�肳�ꂽ�l���r�B��v����t���[����Id��Ԃ��B
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
'�؂�ւ��\�ȃt���[���̏����C�~�f�B�G�C�g�E�B���h�E�ɕ\��
Private Sub IWebDriver_DebugPrintAccessibleFrameInfo()
    execCDP_.PageEnable
    execCDP_.RuntimeEnable
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim i As Long
    For i = 1 To frameInfos.Count
        Debug.Print "index�F" & i & _
                           "�@name�F" & Json_.GetValue(frameInfos(i), "name") & _
                           "�@url�F" & Json_.GetValue(frameInfos(i), "url")
    Next
End Sub

'Public
Private Sub IWebDriver_SwitchFrameByIframeElement(IframeElement As IWebElement)
    Dim frameInfos As Collection: Set frameInfos = GetFrameInfos
    Dim elem As b0x0_WebElement: Set elem = IframeElement
    If elem.FrameId = "" Then
        execHelper_.DisplayErrorAndEnd "��SwitchFrameByIframeElement��" & vbCrLf & _
                                      "�w�肳�ꂽElement��Iframe�ł͂���܂���" & vbCrLf & vbCrLf & _
                                      "Iframe�v�f�������ɓn���ĉ�����" & vbCrLf
    End If
    
    execCDP_.GetTargetInfo
    execCDP_.SetDiscoverTargets

    'frameElement��frameId��������frameId�Ȃ�IsolatedWorld�őΉ�
    Dim ExecutionContextId As Long
    ExecutionContextId = execCDP_.CreateIsolatedFrameWorld(elem.FrameId)
    If ExecutionContextId <> 0 Then
        basicInfo_.ExecutionContextId = ExecutionContextId
        basicInfo_.IsolatedFrameId = elem.FrameId
        Exit Sub
    Else
        execHelper_.DisplayErrorAndEnd "��SwitchFrameByIframeElement��" & vbCrLf & _
                                       "����Iframe�͂��̃��\�b�h�ł͐؂�ւ����ł��܂���B" & vbCrLf & vbCrLf & _
                                      "��SwitchFrameByNameOrUrlOrIndex��" & vbCrLf & _
                                      "�@ �g�p���������Ă��������B" & vbCrLf
    End If
    
    
'    �e�X�g
'    Dim docElem As b0x0_WebElement
'    Dim functionCode As String, responseMsg As String, docObjectId As String
'    functionCode = "function(name){return  this.contentDocument;}"
'    responseMsg = CallJavaScriptFunction(functionCode, elem.ObjectId, "")
'    docObjectId = json_.GetValue(responseMsg, "result", "result", "objectId"))
'    Dim info  As String
'    info = DescribeNode(docObjectId)
    
    
'    frameId��targetId�Ȃ�ڑ��؂�ւ��őΉ�
'    basicInfo_.Update elem.FrameId, json_, False
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
'��1�����Ŏw�肳���Xpath�̗v�f�L���̊m�F
Private Function IWebDriver_IsExistsElementByXpath(xPath As String) As Boolean
    IWebDriver_IsExistsElementByXpath = _
        execCDP_.IsExistsElementBy(xPath, basicInfo_, execHelper_.GetXpathCodes(xPath))
End Function
'Public
'��1�����Ŏw�肳���Id�̗v�f�L���̊m�F
Private Function IWebDriver_IsExistsElementById(Id As String) As Boolean
    IWebDriver_IsExistsElementById = _
        execCDP_.IsExistsElementBy(Id, basicInfo_, execHelper_.GetIdCodes(Id))
End Function
'Public
'��1�����Ŏw�肳���Name�̗v�f�L���̊m�F
Private Function IWebDriver_IsExistsElementByName(name As String) As Boolean
    IWebDriver_IsExistsElementByName = _
        execCDP_.IsExistsElementBy(name, basicInfo_, execHelper_.GetNameCodes(name))
End Function


'Public
'Xpath(��1����)�Ŏw�肳���̗v�f���o������̂��A�ő�ҋ@����(��2����)�҂�
Private Sub IWebDriver_WaitForElementToAppearByXpath(xPath As String, maxWaitTime_sec As Long)
    If execCDP_.WaitForElementToAppearBy(xPath, _
        maxWaitTime_sec, basicInfo_, execHelper_.GetXpathCodes(xPath)) = False Then
        
        execHelper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearByXPath>" & vbCrLf & _
             maxWaitTime_sec & "�b�o�߂��Ă��A�w��̈ȉ��v�f�͌���܂���ł����B" & vbCrLf & _
             xPath
    End If
End Sub
'Public
'Id(��1����)�Ŏw�肳���̗v�f���o������̂��A�ő�ҋ@����(��2����)�҂�
Private Sub IWebDriver_WaitForElementToAppearById(Id As String, maxWaitTime_sec As Long)
    If execCDP_.WaitForElementToAppearBy(Id, _
        maxWaitTime_sec, basicInfo_, execHelper_.GetIdCodes(Id)) = False Then
        
        execHelper_.DisplayErrorAndEnd _
            "<WaitForElementToAppearById>" & vbCrLf & _
             maxWaitTime_sec & "�b�o�߂��Ă��A�w��̈ȉ��v�f�͌���܂���ł����B" & vbCrLf & _
             Id
    End If
End Sub
'Public
'Name(��1����)�Ŏw�肳���̗v�f���o������̂��A�ő�ҋ@����(��2����)�҂�
Private Sub IWebDriver_WaitForElementToAppearByName(name As String, maxWaitTime_sec As Long)
    If execCDP_.WaitForElementToAppearBy(name, _
        maxWaitTime_sec, basicInfo_, execHelper_.GetNameCodes(name)) = False Then
    
        execHelper_.DisplayErrorAndEnd _
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
        execHelper_.SleepByWinAPI 500
        res = execCDP_.HandleAlertDialog
        If InStr(res, "No dialog is showing") = 0 Then
            execHelper_.SleepByWinAPI 500
            Exit Sub
        End If
        
        fileSafeCounter = fileSafeCounter + 1
    Loop While fileSafeCounter < 6

quit:
    execHelper_.DisplayErrorAndEnd "<HandleAlertDialog>" & vbCrLf & _
                                   "�_�C�A���O���\������Ă��܂���B" & vbCrLf & _
                                   "�����̃��\�b�h��Alert���\�������ꍇ�Ɏg�p���܂��B" & vbCrLf
End Sub

Private Sub IWebDriver_SleepByWinAPI(milliSecond As Long)
    execHelper_.SleepByWinAPI milliSecond
End Sub


'�^�u�̈ړ��iIndex�w��F��ԍ��̃^�u���P�Ƃ���j
Private Sub IWebDriver_SwitchTabByIndex(Index As Long)
    Dim newTargetId As String
    newTargetId = events_.PageInfos(Index)("targetId")
    basicInfo_.Update newTargetId, Json_, True
    
    webSocket_.CloseWebSocket
    Set webSocket_ = Nothing
    Set webSocket_ = New a1_WebSocketCommunicator
    If webSocket_.Init(basicInfo_.Host, basicInfo_.WebSocketUrlPath, basicInfo_.DebuggingPort) = False Then End
    
    execCDP_.ActivateTarget basicInfo_.targetId
    execCDP_.AttachToTarget basicInfo_.targetId
    execCDP_.PageEnable
    execCDP_.DOMEnable
    
    basicInfo_.DocumentObjectId = execCDP_.GetDocumentObjectID
    basicInfo_.DocumentNodeId = execCDP_.GetDocumentNodeID
End Sub

Private Sub IWebDriver_CloseWindow()
    execCDP_.SessionClose
End Sub

Private Sub IWebDriver_ScrollWindow(yPixel As Long, Optional xPixel As Long = 0&)
    execCDP_.DispatchMouseEvent "mouseWheel", 0, 0, "none", 0, 0, xPixel, yPixel
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
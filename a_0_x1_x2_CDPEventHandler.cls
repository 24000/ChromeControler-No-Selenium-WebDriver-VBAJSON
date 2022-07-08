VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a_0_x1_x2_CDPEventHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ダイアログ情報
Public DialogInfoDic As Object
Public PageInfos As Collection
Public IframeInfos As Collection

Public IsCurrentTargetScheduledNavigation As Boolean
Public IsCurrentTargetStartedLoading As Boolean
Public IsCurrentTargetClearedScheduledNavigation As Boolean
Public IsCurrentTargetNavigated As Boolean

Public IsIsolatedFrameScheduledNavigation As Boolean
Public IsIsolatedFrameStartedLoading As Boolean
Public IsIsolatedFrameClearedScheduledNavigation As Boolean
Public IsIsolatedFrameNavigated As Boolean

Private Sub Class_Initialize()
    Set DialogInfoDic = CreateObject("Scripting.Dictionary")
    InitializeDialogInfo
    Set PageInfos = New Collection
    Set IframeInfos = New Collection
End Sub
'各種イベント発生時に情報を取得する
Public Sub GetInfo(EventInfo As String, chrome As ChromeDriver)
    Dim eventName As String
    eventName = chrome.Json.GetValueFrom(EventInfo, Array("method"))
    
    If InStr(eventName, "frame") > 0 Then
        Dim id As String
        If InStr(eventName, "Navigated") = 0 Then
            id = chrome.Json.GetValueFrom(EventInfo, Array("params", "frameId"))
        Else
            id = chrome.Json.GetValueFrom(EventInfo, Array("params", "frame", "id"))
        End If
        Dim isCurrentTarget As Boolean
        isCurrentTarget = (id = chrome.TargetId)
        Dim isIsolatedFrameId
        isIsolatedFrameId = (id = chrome.isolatedFrameId_)
    End If
    
    Select Case eventName
        Case "Page.javascriptDialogOpening"
            SetDialogInfo EventInfo, chrome.Json
            Debug.Print "★★DialogOpening"
        Case "Page.javascriptDialogClosed"
            InitializeDialogInfo
            Debug.Print "★★DialogClosed"
        Case "Target.targetCreated"
            AddTargetInfo EventInfo, chrome.Json
        Case "Target.targetInfoChanged"
            ChangeTargetInfo EventInfo, chrome.Json
        Case "Target.targetDestroyed"
            RemoveTargetInfo EventInfo, chrome.Json
            
        Case "Page.frameScheduledNavigation" '非推奨となっているがこれが返ってくる。
            If InStr(EventInfo, "anchorClick") > 0 Then
                If isCurrentTarget Then IsCurrentTargetScheduledNavigation = True
                If isIsolatedFrameId Then
                    IsIsolatedFrameScheduledNavigation = True
                    Debug.Print "★★IsolatedFrameScheduledNavigation"
                End If
            End If
        Case "Page.frameStartedLoading"
            If isCurrentTarget Then IsCurrentTargetStartedLoading = True
            If isIsolatedFrameId Then
                IsIsolatedFrameStartedLoading = True
                chrome.SleepByWindowsAPIInner 1000
            End If
'        Case "Page.frameClearedScheduledNavigation" '非推奨となっているがこれが返ってくる。
'            If isCurrentTarget Then IsCurrentTargetClearedScheduledNavigation = True
'            If isIsolatedFrameId Then IsIsolatedFrameClearedScheduledNavigation = True
        Case "Page.frameNavigated"
            If isCurrentTarget Then IsCurrentTargetNavigated = True
            If isIsolatedFrameId Then
                IsIsolatedFrameNavigated = True
                Debug.Print "★★IsolatedFrameNavigated"
            End If
    End Select
End Sub
'アラートダイアログ出現時に情報を取得
Private Sub SetDialogInfo(ByVal EventInfo As String, Json As a_0_x4_JSONHandler)
    DialogInfoDic("IsExistDialog") = True
    DialogInfoDic("DialogMessage") = Json.GetValueFrom(EventInfo, Array("params", "message"))
    DialogInfoDic("DialogType") = Json.GetValueFrom(EventInfo, Array("params", "type"))
    If Json.GetValueFrom(EventInfo, Array("params", "hasBrowserHandler")) = "True" Then
        DialogInfoDic("HasBrowserHandler") = True
    Else
        DialogInfoDic("HasBrowserHandler") = False
    End If
    DialogInfoDic("DefaultPrompt") = Json.GetValueFrom(EventInfo, Array("params", "defaultPrompt"))
End Sub
'アラートダイアログ出現時に情報を格納する辞書の初期化
Private Sub InitializeDialogInfo()
    DialogInfoDic.RemoveAll
    DialogInfoDic.add "IsExistDialog", False
    DialogInfoDic.add "DialogMessage", ""
    DialogInfoDic.add "DialogType", ""
    DialogInfoDic.add "HasBrowserHandler", False
    DialogInfoDic.add "DefaultPrompt", ""
End Sub

'TargetのTypeにより､それぞれコレクションに情報を取得する｡
Private Sub AddTargetInfo(ByVal EventInfo As String, Json As a_0_x4_JSONHandler)
    Dim tempdic As Object
    Set tempdic = CreateObject("Scripting.Dictionary")
    tempdic.add "targetId", Json.GetValueFrom(EventInfo, Array("params", "targetInfo", "targetId"))
    tempdic.add "title", Json.GetValueFrom(EventInfo, Array("params", "targetInfo", "title"))
    tempdic.add "url", Json.GetValueFrom(EventInfo, Array("params", "targetInfo", "url"))
    tempdic.add "browserContextId", Json.GetValueFrom(EventInfo, Array("params", "targetInfo", "browserContextId"))
    
    Dim targetType As String
    targetType = Json.GetValueFrom(EventInfo, Array("params", "targetInfo", "type"))
    Select Case targetType
      Case "page"
        AddInfoOf PageInfos, tempdic, EventInfo, Json
      Case "iframe"
        AddInfoOf IframeInfos, tempdic, EventInfo, Json
    End Select
End Sub
'Chrome上でTargetが追加されたら、各TargetTypeの情報を管理するCollectionにその情報をDictionaryとしてまとめて追加する。
Private Sub AddInfoOf(infos As Collection, tempdic As Object, _
                                    ByVal EventInfo As String, Json As a_0_x4_JSONHandler)
                                        
    '本来targetInfoChangedで来るべき内容がtargetCreatedで来るケースあり。
    'SetDiscoverTarget使用すると既に開かれているPageもCreateされる模様
    Dim i As Long
    For i = 1 To infos.Count
        If infos(i)("targetId") = tempdic("targetId") Then
            infos(i)("title") = tempdic("targetId")
            infos(i)("url") = tempdic("url")
            infos(i)("browserContextId") = tempdic("browserContextId")
            Exit Sub
        End If
    Next
    infos.add tempdic
End Sub

'Chrome上でPage(タブ)の情報が変更されたら、Page情報を管理するCollection内の該当Pageの情報を変更する。
Private Sub ChangeTargetInfo(ByVal EventInfo As String, Json As a_0_x4_JSONHandler)
    Dim targetType As String
    targetType = Json.GetValueFrom(EventInfo, Array("params", "targetInfo", "type"))
    
    Select Case targetType
      Case "page"
        ChangeInfoOf PageInfos, EventInfo, Json
      Case "iframe"
        ChangeInfoOf IframeInfos, EventInfo, Json
    End Select
End Sub
Private Sub ChangeInfoOf(infos As Collection, _
                                        ByVal EventInfo As String, Json As a_0_x4_JSONHandler)
    Dim TargetId As String
    TargetId = Json.GetValueFrom(EventInfo, Array("params", "targetInfo", "targetId"))
    Dim info As Variant
    For Each info In infos
        If info("targetId") = TargetId Then
            info("title") = Json.GetValueFrom(EventInfo, Array("params", "targetInfo", "title"))
            info("url") = Json.GetValueFrom(EventInfo, Array("params", "targetInfo", "url"))
            info("browserContextId") = Json.GetValueFrom(EventInfo, Array("params", "targetInfo", "browserContextId"))
        End If
    Next
End Sub

'Chrome上でPage(タブ)が削除されたら、Page情報を管理するCollectionからその情報を削除する。
Private Sub RemoveTargetInfo(ByVal EventInfo As String, Json As a_0_x4_JSONHandler)
    Dim targetType As String
    targetType = Json.GetValueFrom(EventInfo, Array("params", "targetInfo", "type"))
    
    Select Case targetType
      Case "page"
        RemoveInfoOf PageInfos, EventInfo, Json
      Case "iframe"
        RemoveInfoOf IframeInfos, EventInfo, Json
    End Select
End Sub
Private Sub RemoveInfoOf(infos As Collection, _
                                        ByVal EventInfo As String, Json As a_0_x4_JSONHandler)
    Dim TargetId As String
    TargetId = Json.GetValueFrom(EventInfo, Array("params", "targetInfo", "targetId"))
    
    Dim i As Long
    For i = 1 To infos.Count
        If infos(i)("targetId") = TargetId Then
            infos.Remove i
        End If
    Next
End Sub

Public Sub ForDebug_ShowtargetInfo()
    Dim i  As Long
    Debug.Print "PageInfos↓==============="
    For i = 1 To PageInfos.Count
        Debug.Print i & "：" & PageInfos(i)("targetId")
    Next
    Debug.Print "IframeInfos↓==============="
    For i = 1 To IframeInfos.Count
        Debug.Print i & "：" & IframeInfos(i)("targetId")
    Next
    Debug.Print "========================"
End Sub

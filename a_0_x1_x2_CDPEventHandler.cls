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
Public pageInfos As Collection
Public IsIsolatedFrameScheduledNavigation As Boolean
Public IsIsolatedFrameScheduledNavigationEnd As Boolean

Private Sub Class_Initialize()
    Set DialogInfoDic = CreateObject("Scripting.Dictionary")
    InitializeDialogInfo
    Set pageInfos = New Collection
End Sub

Public Sub GetInfo(EventInfo As String, chrome As ChromeDriver)
    Dim eventName As String
    eventName = chrome.jsonhandler.GetValueFromJson(EventInfo, Array("method"))
    Select Case eventName
        Case "Page.javascriptDialogOpening"
            SetDialogInfo EventInfo, chrome.jsonhandler
        Case "Page.javascriptDialogClosed"
            InitializeDialogInfo
            
        Case "Target.targetCreated"
            AddTargetInfoIfPage EventInfo, chrome.jsonhandler
        Case "Target.targetInfoChanged"
            ChangeTargetInfoIfPage EventInfo, chrome.jsonhandler
        Case "Target.targetDestroyed"
            RemoveTargetInfoIfPage EventInfo, chrome.jsonhandler
            
        Case "Page.frameScheduledNavigation" '非推奨となっているがこれが返ってくる。
            If chrome.isolatedFrameId_ = "" Then Exit Sub
            SetFlag_IfIsolatedFrameScheduledNavigation EventInfo, chrome
        Case "Page.frameClearedScheduledNavigation" '非推奨となっているがこれが返ってくる。
            If chrome.isolatedFrameId_ = "" Then Exit Sub
            SetFlag_IfIsolatedFrameScheduledNavigationEnd EventInfo, chrome
        Case "Page.frameNavigated"
            If chrome.isolatedFrameId_ = "" Then Exit Sub
            SetFlag_IfIsolatedFrameScheduledNavigationEnd EventInfo, chrome
    End Select
End Sub

Private Sub SetDialogInfo(ByVal EventInfo As String, jsonhandler As a_0_x4_JSONHandler)
    DialogInfoDic("IsExistDialog") = True
    DialogInfoDic("DialogMessage") = jsonhandler.GetValueFromJson(EventInfo, Array("params", "message"))
    DialogInfoDic("DialogType") = jsonhandler.GetValueFromJson(EventInfo, Array("params", "type"))
    If jsonhandler.GetValueFromJson(EventInfo, Array("params", "hasBrowserHandler")) = "True" Then
        DialogInfoDic("HasBrowserHandler") = True
    Else
        DialogInfoDic("HasBrowserHandler") = False
    End If
    DialogInfoDic("DefaultPrompt") = jsonhandler.GetValueFromJson(EventInfo, Array("params", "defaultPrompt"))
End Sub

Private Sub InitializeDialogInfo()
    DialogInfoDic.RemoveAll
    DialogInfoDic.Add "IsExistDialog", False
    DialogInfoDic.Add "DialogMessage", ""
    DialogInfoDic.Add "DialogType", ""
    DialogInfoDic.Add "HasBrowserHandler", False
    DialogInfoDic.Add "DefaultPrompt", ""
End Sub

'Chrome上でPage(タブ)が追加されたら、Page情報を管理するCollectionにその情報をDictionaryとしてまとめて追加する。
Private Sub AddTargetInfoIfPage(ByVal EventInfo As String, jsonhandler As a_0_x4_JSONHandler)
    Dim targetType As String
    targetType = jsonhandler.GetValueFromJson(EventInfo, Array("params", "targetInfo", "type"))
    If targetType <> "page" Then Exit Sub
    
    Dim tempDic As Object
    Set tempDic = CreateObject("Scripting.Dictionary")
    tempDic.Add "targetId", jsonhandler.GetValueFromJson(EventInfo, Array("params", "targetInfo", "targetId"))
    tempDic.Add "title", jsonhandler.GetValueFromJson(EventInfo, Array("params", "targetInfo", "title"))
    tempDic.Add "url", jsonhandler.GetValueFromJson(EventInfo, Array("params", "targetInfo", "url"))
    tempDic.Add "browserContextId", jsonhandler.GetValueFromJson(EventInfo, Array("params", "targetInfo", "browserContextId"))
    '本来targetInfoChangedで来るべき内容がtargetCreatedで来るケースあり。
    'SetDiscoverTarget使用すると既に開かれているPageもCreateされる模様
    Dim i As Long
    For i = 1 To pageInfos.Count
        If pageInfos(i)("targetId") = tempDic("targetId") Then
            pageInfos(i)("title") = tempDic("targetId")
            pageInfos(i)("url") = tempDic("url")
            pageInfos(i)("browserContextId") = tempDic("browserContextId")
            Exit Sub
        End If
    Next

    pageInfos.Add tempDic
End Sub

'Chrome上でPage(タブ)の情報が変更されたら、Page情報を管理するCollection内の該当Pageの情報を変更する。
Private Sub ChangeTargetInfoIfPage(ByVal EventInfo As String, jsonhandler As a_0_x4_JSONHandler)
    Dim targetType As String
    targetType = jsonhandler.GetValueFromJson(EventInfo, Array("params", "targetInfo", "type"))
    If targetType <> "page" Then Exit Sub
    
    Dim targetId As String
    targetId = jsonhandler.GetValueFromJson(EventInfo, Array("params", "targetInfo", "targetId"))
    Dim pageInfo As Variant
    For Each pageInfo In pageInfos
        If pageInfo("targetId") = targetId Then
            pageInfo("title") = jsonhandler.GetValueFromJson(EventInfo, Array("params", "targetInfo", "title"))
            pageInfo("url") = jsonhandler.GetValueFromJson(EventInfo, Array("params", "targetInfo", "url"))
            pageInfo("browserContextId") = jsonhandler.GetValueFromJson(EventInfo, Array("params", "targetInfo", "browserContextId"))
        End If
    Next
End Sub

'Chrome上でPage(タブ)が削除されたら、Page情報を管理するCollectionからその情報を削除する。
Private Sub RemoveTargetInfoIfPage(ByVal EventInfo As String, jsonhandler As a_0_x4_JSONHandler)
    Dim targetType As String
    targetType = jsonhandler.GetValueFromJson(EventInfo, Array("params", "targetInfo", "type"))
    If targetType <> "page" Then Exit Sub
    
    Dim targetId As String
    targetId = jsonhandler.GetValueFromJson(EventInfo, Array("params", "targetInfo", "targetId"))
    Dim i As Long
    For i = 1 To pageInfos.Count
        If pageInfos(i)("targetId") = targetId Then
            pageInfos.Remove i
        End If
    Next
End Sub

Private Sub SetFlag_IfIsolatedFrameScheduledNavigation( _
                    ByVal EventInfo As String, chrome As ChromeDriver)
    Dim frameId As String
    frameId = chrome.jsonhandler.GetValueFromJson(EventInfo, Array("params", "frameId"))
    If frameId <> chrome.isolatedFrameId_ Then Exit Sub

    IsIsolatedFrameScheduledNavigation = True
End Sub

Private Sub SetFlag_IfIsolatedFrameScheduledNavigationEnd( _
                    ByVal EventInfo As String, chrome As ChromeDriver)
                    
    Dim methodName As String
    methodName = chrome.jsonhandler.GetValueFromJson(EventInfo, Array("method"))

    Dim frameId As String
    If methodName = "Page.frameNavigated" Then
        frameId = chrome.jsonhandler.GetValueFromJson(EventInfo, Array("params", "frame", "id"))
    ElseIf methodName = "Page.frameClearedScheduledNavigation" Then
        frameId = chrome.jsonhandler.GetValueFromJson(EventInfo, Array("params", "frameId"))
    End If

    If frameId <> chrome.isolatedFrameId_ Then Exit Sub

    IsIsolatedFrameScheduledNavigationEnd = True
End Sub


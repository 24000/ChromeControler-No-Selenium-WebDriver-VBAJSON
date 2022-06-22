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
Public pageInfoDic As Object
Public IsIsolatedFrameScheduledNavigation As Boolean
Public IsIsolatedFrameScheduledNavigationEnd As Boolean
Public CurrentPageNo As Long

Private Sub Class_Initialize()
    Set DialogInfoDic = CreateObject("Scripting.Dictionary")
    InitializeDialogInfo
    Set pageInfoDic = CreateObject("Scripting.Dictionary")
    CurrentPageNo = 0
End Sub

Public Sub GetInfo(EventInfoJson As String, chrome As ChromeControler)
    Dim eventName As String
    eventName = chrome.jsonhandler.GetValueFromJson(EventInfoJson, Array("method"))
    Select Case eventName
        Case "Page.javascriptDialogOpening"
            SetDialogInfo EventInfoJson, chrome.jsonhandler
        Case "Page.javascriptDialogClosed"
            InitializeDialogInfo
        Case "Target.targetCreated"
            AddTargetInfoIfPage EventInfoJson, chrome.jsonhandler
        Case "Target.targetInfoChanged"
            ChangeTargetInfoIfPage EventInfoJson, chrome.jsonhandler
        Case "Target.targetDestroyed"
            RemoveTargetInfoIfPage EventInfoJson, chrome.jsonhandler
        Case "Page.frameScheduledNavigation" '非推奨となっているがこれが返ってくる。
            If chrome.isolatedFrameId_ = 0 Then Exit Sub
            SetFlag_IfIsolatedFrameScheduledNavigation EventInfoJson, chrome
        Case "Page.frameClearedScheduledNavigation" '非推奨となっているがこれが返ってくる。
            If chrome.isolatedFrameId_ = 0 Then Exit Sub
            SetFlag_IfIsolatedFrameScheduledNavigationEnd EventInfoJson, chrome
        Case "Page.frameNavigated"
            If chrome.isolatedFrameId_ = 0 Then Exit Sub
            SetFlag_IfIsolatedFrameScheduledNavigationEnd EventInfoJson, chrome
    End Select
End Sub

Private Sub SetDialogInfo(ByVal EventInfoJson As String, jsonhandler As a_0_x4_JSONHandler)
    DialogInfoDic("IsExistDialog") = True
    DialogInfoDic("DialogMessage") = jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "message"))
    DialogInfoDic("DialogType") = jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "type"))
    If jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "hasBrowserHandler")) = "True" Then
        DialogInfoDic("HasBrowserHandler") = True
    Else
        DialogInfoDic("HasBrowserHandler") = False
    End If
    DialogInfoDic("DefaultPrompt") = jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "defaultPrompt"))
End Sub

Private Sub InitializeDialogInfo()
    DialogInfoDic.RemoveAll
    DialogInfoDic.Add "IsExistDialog", False
    DialogInfoDic.Add "DialogMessage", ""
    DialogInfoDic.Add "DialogType", ""
    DialogInfoDic.Add "HasBrowserHandler", False
    DialogInfoDic.Add "DefaultPrompt", ""
End Sub

Private Sub AddTargetInfoIfPage(ByVal EventInfoJson As String, jsonhandler As a_0_x4_JSONHandler)
    Dim targetType As String
    targetType = jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "targetInfo", "type"))
    If targetType <> "page" Then Exit Sub
    
    Dim tempDic As Object
    Set tempDic = CreateObject("Scripting.Dictionary")
    tempDic.Add "targetId", jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "targetInfo", "targetId"))
    tempDic.Add "title", jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "targetInfo", "title"))
    tempDic.Add "url", jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "targetInfo", "url"))
    tempDic.Add "browserContextId", jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "targetInfo", "browserContextId"))
    
    pageInfoDic.Add "page" & pageInfoDic.Count, tempDic
End Sub

Private Sub ChangeTargetInfoIfPage(ByVal EventInfoJson As String, jsonhandler As a_0_x4_JSONHandler)
    Dim targetType As String
    targetType = jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "targetInfo", "type"))
    If targetType <> "page" Then Exit Sub
    
    Dim targetId As String
    targetId = jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "targetInfo", "targetId"))
    Dim key As Variant
    For Each key In pageInfoDic
        If pageInfoDic(key)("targetId") = targetId Then
            pageInfoDic(key)("title") = jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "targetInfo", "title"))
            pageInfoDic(key)("url") = jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "targetInfo", "url"))
            pageInfoDic(key)("browserContextId") = jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "targetInfo", "browserContextId"))
        End If
    Next
End Sub

Private Sub RemoveTargetInfoIfPage(ByVal EventInfoJson As String, jsonhandler As a_0_x4_JSONHandler)
    Dim targetType As String
    targetType = jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "targetInfo", "type"))
    If targetType <> "page" Then Exit Sub
    
    Dim targetId As String
    targetId = jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "targetInfo", "targetId"))
    Dim key As Variant
    For Each key In pageInfoDic
        If pageInfoDic(key)("targetId") = targetId Then
            pageInfoDic.Remove key
        End If
    Next
End Sub

Private Sub SetFlag_IfIsolatedFrameScheduledNavigation( _
                    ByVal EventInfoJson As String, chrome As ChromeControler)
    Dim frameId As String
    frameId = chrome.jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "frameId"))
    If frameId <> chrome.isolatedFrameId_ Then Exit Sub

    IsIsolatedFrameScheduledNavigation = True
End Sub

Private Sub SetFlag_IfIsolatedFrameScheduledNavigationEnd( _
                    ByVal EventInfoJson As String, chrome As ChromeControler)
                    
    Dim methodName As String
    methodName = chrome.jsonhandler.GetValueFromJson(EventInfoJson, Array("method"))

    Dim frameId As String
    If methodName = "Page.frameNavigated" Then
        frameId = chrome.jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "frame", "id"))
    ElseIf methodName = "Page.frameClearedScheduledNavigation" Then
        frameId = chrome.jsonhandler.GetValueFromJson(EventInfoJson, Array("params", "frameId"))
    End If

    If frameId <> chrome.isolatedFrameId_ Then Exit Sub

    IsIsolatedFrameScheduledNavigationEnd = True
End Sub


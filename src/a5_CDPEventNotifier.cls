VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a5_CDPEventNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Event JavaScriptDialogOpening(json As String)
Public Event JavaScriptDialogClosed(json As String)

Public Event TargetCreated(json As String)
Public Event TargetInfoChange(json As String)
Public Event TargetDestroyed(json As String)

Public Event FrameScheduledNavigation(json As String)
Public Event frameClearedScheduledNavigation(json As String)
Public Event FrameStartedLoading(json As String)
Public Event FrameNavigated(json As String)

Public Event ExecutionContextCreated(json As String)
Public Event ExecutionContextDestroyed(json As String)

Public Event windowOpen()


Public Sub Notify(json As String)
    Dim eventName As String
    eventName = xJson.GetValue(json, "method")
    
'    If InStr(eventName, "frame") > 0 Then
'
'        Dim Id As String
'        If InStr(eventName, "Navigated") = 0 Then
'            Id = xJson.GetValue(json, "params", "frameId")
'        Else
'            Id = xJson.GetValue(json, "params", "frame", "id")
'        End If
'        Dim isCurrentTarget As Boolean
'        isCurrentTarget = (Id = basicInfo_.targetId)
'        Dim isIsolatedFrameId
'        isIsolatedFrameId = (Id = basicInfo_.IsolatedFrameId)
'    End If
    
    Select Case eventName
        Case "Page.javascriptDialogOpening"
            RaiseEvent JavaScriptDialogOpening(xJson.GetObject(json, "params", "targetInfo"))
        Case "Page.javascriptDialogClosed"
            RaiseEvent JavaScriptDialogClosed(xJson.GetObject(json, "params", "targetInfo"))

            
        Case "Target.targetCreated"
            RaiseEvent TargetCreated(xJson.GetObject(json, "params", "targetInfo"))
        Case "Target.targetInfoChanged"
            RaiseEvent TargetInfoChange(xJson.GetObject(json, "params", "targetInfo"))
        Case "Target.targetDestroyed"
            RaiseEvent TargetDestroyed(xJson.GetObject(json, "params"))

            
        Case "Page.frameScheduledNavigation" '”ñ„§‚Æ‚È‚Á‚Ä‚¢‚é‚ª‚±‚ê‚ª•Ô‚Á‚Ä‚­‚éB
            RaiseEvent FrameScheduledNavigation(xJson.GetObject(json, "params", "targetInfo"))
        Case "Page.frameStartedLoading"
            RaiseEvent FrameStartedLoading(xJson.GetObject(json, "params", "targetInfo"))
        Case "Page.frameClearedScheduledNavigation" '”ñ„§‚Æ‚È‚Á‚Ä‚¢‚é‚ª‚±‚ê‚ª•Ô‚Á‚Ä‚­‚éB
            RaiseEvent frameClearedScheduledNavigation(xJson.GetObject(json, "params", "targetInfo"))
        Case "Page.frameNavigated"
            RaiseEvent FrameNavigated(xJson.GetObject(json, "params", "targetInfo"))

        Case "Runtime.executionContextCreated"
            RaiseEvent ExecutionContextCreated(xJson.GetObject(json, "params", "targetInfo"))
        Case "Runtime.executionContextDestroyed"
            RaiseEvent ExecutionContextDestroyed(xJson.GetObject(json, "params", "targetInfo"))

    End Select
End Sub

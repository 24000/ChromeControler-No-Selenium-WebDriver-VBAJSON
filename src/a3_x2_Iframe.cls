VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a3_x2_Iframe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IsScheduledNavigation As Boolean
Public IsStartedLoading As Boolean
Public IsClearedScheduledNavigation As Boolean
Public IsTargetNavigated As Boolean
Public List As Object

Public IsolatedId As String



Private Sub Class_Initialize()
    Set List = CreateObject("Scripting.Dictionary")
End Sub

Public Function Exists(targetId As String) As Boolean
    If List.Exists(targetId) Then
        Exists = True
    Else
        Exists = False
    End If
End Function

Public Sub AddTarget(id As String, t As String, json As String)
    Dim target As Object: Set target = CreateObject("Scripting.Dictionary")
    target.add "id", id
    target.add "title", xJson.GetValue(json, "title")
    target.add "type", t
    target.add "url", xJson.GetValue(json, "url")
    target.add "browserContextId", xJson.GetValue(json, "browserContextId")
    
    List.add id, target
End Sub

Public Sub ChangeTarget(id As String, json As String)
    Dim target As Object: Set target = List(id)
    target("title") = xJson.GetValue(json, "title")
    target("url") = xJson.GetValue(json, "url")
    target("browserContextId") = xJson.GetValue(json, "browserContextId")
End Sub

Public Sub RemoveTarget(id As String)
    List.Remove id
End Sub

Public Function IsIsolatedFrameID(frameId As String) As Boolean
    Dim i As Long
    For i = 1 To List.Count
        If List(i)("id") = frameId Then
            IsIsolatedFrameID = True
            Exit Function
        End If
    Next
    
    IsIsolatedFrameID = False
End Function


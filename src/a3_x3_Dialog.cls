VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a3_x3_Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dialogInfo_ As Object

Private Sub Class_Initialize()
    Set dialogInfo_ = CreateObject("Scripting.Dictionary")
    InitializeDialogInfo
End Sub

'アラートダイアログ出現時に情報を格納する辞書の初期化
Private Sub InitializeDialogInfo()
    dialogInfo_.RemoveAll
    dialogInfo_.add "IsExistDialog", False
    dialogInfo_.add "DialogMessage", ""
    dialogInfo_.add "DialogType", ""
    dialogInfo_.add "HasBrowserHandler", False
    dialogInfo_.add "DefaultPrompt", ""
End Sub

'アラートダイアログ出現時に情報を取得
Private Sub SetDialogInfo(ByVal EventInfo As String)
    dialogInfo_("IsExistDialog") = True
    dialogInfo_("DialogMessage") = xJson.GetValue(EventInfo, "params", "message")
    dialogInfo_("DialogType") = xJson.GetValue(EventInfo, "params", "type")
    If xJson.GetValue(EventInfo, "params", "hasBrowserHandler") = "True" Then
        dialogInfo_("HasBrowserHandler") = True
    Else
        dialogInfo_("HasBrowserHandler") = False
    End If
    dialogInfo_("DefaultPrompt") = xJson.GetValue(EventInfo, "params", "defaultPrompt")
End Sub



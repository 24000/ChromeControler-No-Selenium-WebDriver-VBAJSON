VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a4x2_CDPMessageHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private webSocket_ As a1_WebSocketCommunicator
Private Json_ As a2_JSONHandler
Private events_ As a5_CDPEventHandler

Public Sub Init(WebSocket As a1_WebSocketCommunicator, _
                        Json As a2_JSONHandler, _
                        events As a5_CDPEventHandler)
    Set webSocket_ = WebSocket
    Set Json_ = Json
    Set events_ = events
End Sub

'message�������M����M
'��M���ɂ́A���M�ɑ΂����M�����ł͂Ȃ��A�N���[������̃C�x���g�ʒm���������Ă���B
'�C�x���g�ʒm�̏ꍇ�Arespons��id���܂܂�Ȃ����߁Aid�̗L���ő��M�ɑ΂����M�Ȃ̂��C�x���g�ʒm�Ȃ̂��𔻒f
'����:���M�ɐ������Ă��A�������M����Json�ɕs��������ꍇ�Arespons�ɂ�id���܂܂�Ȃ��B
Public Function SendAndReceive(sendMsg As String) As String
    
    Dim sendId As Long: sendId = Json_.GetValue(sendMsg, "id")
    
    Debug.Print vbCrLf & ">>>>>>" & "���M���e�F" & sendMsg
    If webSocket_.SendMessage(sendMsg) = False Then GoTo quit

    Dim res As String, resId As Variant
    Do
        res = webSocket_.GetMessage
        If res = "��M�G���[����" Then GoTo quit
        
        resId = Json_.GetValue(res, "id")
        If resId = "" Then
            If InStr(res, "method") > 0 Then
                events_.GetInfo res
                '�����D���_�C�A���O�����݂��A����DialogOpen�̃C�x���g�ʒm������Ƃ��̌�A���[�v��M�łƂ܂�B
                '��F�N���b�N���M�˃N���b�N�ɑ΂����M�܂ł̊Ԃɂ��̃C�x���g�ʒm������ƁA��M���ł����Ɏ~�܂�B
                '���̂��߁ADialogOpen�̒ʒm�̂ݓ���Ή��B��������ɂ͂����_�C�A���O��������K�v����B
                If events_.DialogInfoDic("IsExistDialog") Then Exit Function
            ElseIf InStr(res, "error") > 0 Then
                GoTo quit
            End If
        End If
    Debug.Print "��M�F" & res
    Loop While (sendId <> resId)
    
    SendAndReceive = res
Exit Function

quit:
Dim errMsg As String
If res = "" Then
    errMsg = "���M�G���[�����F�C�~�f�B�G�C�g�Q��"
ElseIf res = "��M�G���[����" Then
    errMsg = "��M�G���[����:�C�~�f�B�G�C�g�Q��"
Else
    errMsg = "�N���[������̃G���[�ʒm����:�C�~�f�B�G�C�g�Q��"
End If
MsgBox errMsg, vbCritical, "�G���["
webSocket_.Http.CloseHInternetHandles
End
End Function


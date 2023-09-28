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
Private event_ As a5_CDPEventNotifier

Public Sub Init(WebSocket As a1_WebSocketCommunicator, _
                      notifier As a5_CDPEventNotifier)
    Set webSocket_ = WebSocket
    Set event_ = notifier
End Sub

'message�������M����M
'��M���ɂ́A���M�ɑ΂����M�����ł͂Ȃ��A�N���[������̃C�x���g�ʒm���������Ă���B
'�C�x���g�ʒm�̏ꍇ�Arespons��id���܂܂�Ȃ����߁Aid�̗L���ő��M�ɑ΂����M�Ȃ̂��C�x���g�ʒm�Ȃ̂��𔻒f
'����:���M�ɐ������Ă��A�������M����Json�ɕs��������ꍇ�Arespons�ɂ�id���܂܂�Ȃ��B
Public Function SendAndReceive(sendMsg As String) As String
    
    Dim sendId As Long: sendId = xJson.GetValue(sendMsg, "id")
    
    Debug.Print vbCrLf & ">>>>>>" & "���M���e�F" & sendMsg
    If webSocket_.SendMessage(sendMsg) = False Then GoTo quit

    Dim res As String, resId As Variant
    Do
        res = webSocket_.GetMessage
        If res = "��M�G���[����" Then GoTo quit
        
        resId = xJson.GetValue(res, "id")
        If resId = "" Then
            If InStr(res, "method") > 0 Then
                event_.Notify res
                '�����D���_�C�A���O�����݂��A����DialogOpen�̃C�x���g�ʒm������Ƃ��̌�A���[�v��M�łƂ܂�B
                '��F�N���b�N���M�˃N���b�N�ɑ΂����M�܂ł̊Ԃɂ��̃C�x���g�ʒm������ƁA��M���ł����Ɏ~�܂�B
                '���̂��߁ADialogOpen�̒ʒm�̂ݓ���Ή��B��������ɂ͂����_�C�A���O��������K�v����B
'                If event_.DialogInfoDic("IsExistDialog") Then Exit Function
            ElseIf InStr(res, "error") > 0 Then
                GoTo quit
            End If
        End If
    Debug.Print "��M�F" & res
    Loop While (sendId <> resId)
    
    SendAndReceive = res
    
    Dim resobj As Object
    Set resobj = parseJSON(res)
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
webSocket_.http_.CloseHInternetHandles
End
End Function




Function parseJSON(jsonString As Variant) As Object
    Dim result As Object
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim value As Variant
    Dim key As String
    
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Pattern = "(?:\{[^\[\]\{\}\:\,]*\})|(?:\[[^\[\]\{\}\:\,]*\])|(?:[^\[\]\{\}\:\,]+)"
    Set matches = regex.Execute(jsonString)
    Set result = CreateObject("Scripting.Dictionary")
    
    For Each match In matches
        value = match
        If Left(value, 1) = """" And Right(value, 1) = """" Then
            value = Mid(value, 2, Len(value) - 2)
        ElseIf IsNumeric(value) Then
            value = CDbl(value)
        ElseIf value = "true" Then
            value = True
        ElseIf value = "false" Then
            value = False
        ElseIf value = "null" Then
            value = Null
        End If
        
        If IsObject(value) Then
            Set value = parseJSON(value)
        End If
        
        If IsArray(value) Then
            Dim i As Long
            Dim tempDict As Object
            Set tempDict = CreateObject("Scripting.Dictionary")
            For i = 0 To UBound(value)
                tempDict.add i, value(i)
            Next i
            Set value = tempDict
        End If
        
        If result.Count Mod 2 = 0 Then
            If IsObject(value) Then
                Set result(value("key")) = value("value")
            Else
                key = value
            End If
        Else
            result(key) = value
        End If
    Next match
    
    Set parseJSON = result
End Function
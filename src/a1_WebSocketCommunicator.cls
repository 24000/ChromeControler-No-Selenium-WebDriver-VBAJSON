VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a1_WebSocketCommunicator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const NO_ERROR As Long = 0
Private Declare PtrSafe Function WinHttpWebSocketSend Lib "WinHttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal eBufferType As Long, _
   ByVal pvBuffer As LongPtr, _
   ByVal dwBufferLength As Long _
   ) As Long
Private Const WINHTTP_WEB_SOCKET_BINARY_MESSAGE_BUFFER_TYPE = 0 '�o�C�i�����b�Z�[�W�S�̂��Ō�̕����B
Private Const WINHTTP_WEB_SOCKET_BINARY_FRAGMENT_BUFFER_TYPE = 1 '�o�C�i�����b�Z�[�W�̈ꕔ���̂ݡ
Private Const WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE = 2 'UTF-8���b�Z�[�W�S�̂��Ō�̕����B
Private Const WINHTTP_WEB_SOCKET_UTF8_FRAGMENT_BUFFER_TYPE = 3 'UTF-8 ���b�Z�[�W�̈ꕔ���̂݁B
Private Const WINHTTP_WEB_SOCKET_CLOSE_BUFFER_TYPE = 4 'CloseFrame�B

Private Declare PtrSafe Function WinHttpWebSocketReceive Lib "WinHttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByRef pvBuffer As Any, _
   ByVal dwBufferLength As Long, _
   ByRef pdwBytesRead As Long, _
   ByRef peBufferType As Long _
   ) As Long

Private Declare PtrSafe Function WinHttpWebSocketClose Lib "WinHttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByVal usStatus As Integer, _
   ByVal pvReason As LongPtr, _
   ByVal dwReasonLength As Long _
   ) As Long
Private Const WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS = 1000
Private Const WINHTTP_WEB_SOCKET_ENDPOINT_TERMINATED_CLOSE_STATUS = 1001
Private Const WINHTTP_WEB_SOCKET_PROTOCOL_ERROR_CLOSE_STATUS = 1002
Private Const WINHTTP_WEB_SOCKET_INVALID_DATA_TYPE_CLOSE_STATUS = 1003
Private Const WINHTTP_WEB_SOCKET_EMPTY_CLOSE_STATUS = 1005
Private Const WINHTTP_WEB_SOCKET_ABORTED_CLOSE_STATUS = 1006
Private Const WINHTTP_WEB_SOCKET_INVALID_PAYLOAD_CLOSE_STATUS = 1007
Private Const WINHTTP_WEB_SOCKET_POLICY_VIOLATION_CLOSE_STATUS = 1008
Private Const WINHTTP_WEB_SOCKET_MESSAGE_TOO_BIG_CLOSE_STATUS = 1009
Private Const WINHTTP_WEB_SOCKET_UNSUPPORTED_EXTENSIONS_CLOSE_STATUS = 1010
Private Const WINHTTP_WEB_SOCKET_SERVER_ERROR_CLOSE_STATUS = 1011
Private Const WINHTTP_WEB_SOCKET_SECURE_HANDSHAKE_ERROR_CLOSE_STATUS = 1015

Private Declare PtrSafe Function WinHttpWebSocketQueryCloseStatus Lib "WinHttp" ( _
   ByVal hWebSocket As LongPtr, _
   ByRef usStatus As Integer, _
   ByRef pvReason As Any, _
   ByVal dwReasonLength As Long, _
   ByRef pdwReasonLengthConsumed As Long _
   ) As Long
   
'Websocket��M�Ɏg�p(����\����)
Private Type response
    Buffer(4095) As Byte   '��1����
    BufferLength As Long    '��2����
    ReceiveBytes As Long    '��3����
    Status As Long              '��4����
    CurrentPointer As Long  '��5����
    result As Long               '�߂�l
    collect As Collection
End Type

Public http_ As a1x1_HTTPCommunicator


Public Function Init(http As a1x1_HTTPCommunicator, path As String) As Boolean
    Set http_ = http
    If http_.Connect(path) = False Then GoTo quit
    If http_.UpgradeToWebSocket() = False Then GoTo quit
    
    Init = True
Exit Function
quit:
    http_.CloseHInternetHandles
    Init = False
End Function

'�N���[���փ��b�Z�[�W���M
Public Function SendMessage(message As String) As Boolean

    message = StrConv(message, vbFromUnicode)
    Dim ary() As Byte: ary = message
    Dim messageByteSize As Long: messageByteSize = LenB(message)
    Dim result As Long
    result = WinHttpWebSocketSend(http_.websockethandle, _
                        WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE, _
                        StrPtr(message), messageByteSize)
    If result = NO_ERROR Then
        SendMessage = True
    Else
        Debug.Print "���M�G���[�����F(�߂�l)" & result & ":(LastDllError)" & Err.LastDllError
        SendMessage = False
    End If
    
End Function

'�N���[������̃��X�|���X(JSON)��M
Public Function GetMessage() As String
    Dim res As response
    InitializeBuffer res.Buffer, res.CurrentPointer, res.BufferLength
    Set res.collect = New Collection

    Do '���X�|���X��M���[�v
        res.result = WinHttpWebSocketReceive( _
                        http_.websockethandle, res.Buffer(res.CurrentPointer), res.BufferLength, _
                        res.ReceiveBytes, res.Status)
        If res.result <> NO_ERROR Then GoTo quit

        If res.Status = WINHTTP_WEB_SOCKET_UTF8_MESSAGE_BUFFER_TYPE Then
            res.collect.add res.Buffer
            GetMessage = MakeResponseMessage(res.collect)
            Exit Function
        End If
        
        UpdateBufferInfo res.CurrentPointer, res.BufferLength, res.ReceiveBytes
        If (res.BufferLength <= 0) Then
            res.collect.add res.Buffer
            InitializeBuffer res.Buffer, res.CurrentPointer, res.BufferLength
        End If
    Loop
    
quit:
Debug.Print "��M�G���[�����F(�߂�l)" & res.result
GetMessage = "��M�G���[����"
End Function
'�o�b�t�@�Ǘ����X�V
Private Sub UpdateBufferInfo(CurrentPointer As Long, BufferLength As Long, ReceiveBytes)
    CurrentPointer = CurrentPointer + ReceiveBytes
    BufferLength = BufferLength - ReceiveBytes
End Sub
'�o�b�t�@������
Private Sub InitializeBuffer(responseBuffer As Variant, CurrentPointer As Long, BufferLength As Long)
    Erase responseBuffer
    CurrentPointer = 0
    BufferLength = UBound(responseBuffer) + 1
End Sub

'�t���O�����g�g�ݗ��āF�����R�[�h��ANSI��UTF16�֕ϊ��FNull����
Private Function MakeResponseMessage(collect As Collection) As String
    Dim ary As Variant
    Dim message As String
    For Each ary In collect
        message = message & ary
    Next
    message = StrConv(message, vbUnicode)
    MakeResponseMessage = Left(message, InStr(message, vbNullChar) - 1)
End Function



Public Function CloseWebSocket() As Boolean
    Dim result As Long
    result = WinHttpWebSocketClose(http_.websockethandle, _
                        WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS, 0, 0)
    If result <> NO_ERROR Then GoTo quit

    ' �T�[�o�[����Ԃ��ꂽ�N���[�Y�X�e�[�^�X���m�F����B
    Dim CloseReasonBuffer(123) As Byte
    Dim BufferSize As Long: BufferSize = UBound(CloseReasonBuffer) + 1
    Dim Status As Integer
    Dim CloseReasonLength As Long
    
    result = WinHttpWebSocketQueryCloseStatus(http_.websockethandle, Status, _
                        CloseReasonBuffer(0), BufferSize, CloseReasonLength)
    If result <> NO_ERROR Then GoTo quit
    If Status <> WINHTTP_WEB_SOCKET_SUCCESS_CLOSE_STATUS Then GoTo quit

    CloseWebSocket = True
Exit Function

quit:
If result <> NO_ERROR Then Debug.Print result
Debug.Print "The server closed the connection with status code: " & Status
Dim strCloseReason As String
Dim i As Long
For i = LBound(CloseReasonBuffer) To UBound(CloseReasonBuffer)
    If CloseReasonBuffer(i) <> 0 Then
      strCloseReason = strCloseReason & Chr(CloseReasonBuffer(i))
    End If
Next
If Len(strCloseReason) > 0 Then
    Debug.Print " and reason: " & strCloseReason
End If
CloseWebSocket = False
End Function



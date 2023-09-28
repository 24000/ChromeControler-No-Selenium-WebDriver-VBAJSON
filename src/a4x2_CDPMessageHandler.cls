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

'message処理送信＆受信
'受信時には、送信に対する受信だけではなく、クロームからのイベント通知も混ざってくる。
'イベント通知の場合、responsにidが含まれないため、idの有無で送信に対する受信なのかイベント通知なのかを判断
'注意:送信に成功しても、もし送信したJsonに不備がある場合、responsにはidが含まれない。
Public Function SendAndReceive(sendMsg As String) As String
    
    Dim sendId As Long: sendId = xJson.GetValue(sendMsg, "id")
    
    Debug.Print vbCrLf & ">>>>>>" & "送信内容：" & sendMsg
    If webSocket_.SendMessage(sendMsg) = False Then GoTo quit

    Dim res As String, resId As Variant
    Do
        res = webSocket_.GetMessage
        If res = "受信エラー発生" Then GoTo quit
        
        resId = xJson.GetValue(res, "id")
        If resId = "" Then
            If InStr(res, "method") > 0 Then
                event_.Notify res
                '制御を奪うダイアログが存在し、そのDialogOpenのイベント通知がくるとその後、ループ受信でとまる。
                '例：クリック送信⇒クリックに対する受信までの間にそのイベント通知が来ると、受信ができずに止まる。
                'そのため、DialogOpenの通知のみ特殊対応。抜けた後にはすぐダイアログ処理する必要あり。
'                If event_.DialogInfoDic("IsExistDialog") Then Exit Function
            ElseIf InStr(res, "error") > 0 Then
                GoTo quit
            End If
        End If
    Debug.Print "受信：" & res
    Loop While (sendId <> resId)
    
    SendAndReceive = res
    
    Dim resobj As Object
    Set resobj = parseJSON(res)
Exit Function

quit:
Dim errMsg As String
If res = "" Then
    errMsg = "送信エラー発生：イミディエイト参照"
ElseIf res = "受信エラー発生" Then
    errMsg = "受信エラー発生:イミディエイト参照"
Else
    errMsg = "クロームからのエラー通知あり:イミディエイト参照"
End If
MsgBox errMsg, vbCritical, "エラー"
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

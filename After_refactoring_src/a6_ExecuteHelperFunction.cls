VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a6_ExecuteHelperFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Public Sub SleepByWinAPI(ms As Long)
    Sleep ms
End Sub

'FindByXPathのJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetXpathCodes(xPath As String) As Variant
    Dim codes(1) As String
    codes(0) = "document.evaluate('" & xPath & "', document, null, " & _
                        "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;"
    codes(1) = "function(xpath){return  document.evaluate(xpath, document, null, " & _
                            "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;}"
    GetXpathCodes = codes
End Function

'FindByIdのJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetIdCodes(Id As String) As Variant
    Dim codes(1) As String
    codes(0) = "document.getElementById('" & Id & "');"
    codes(1) = "function(id){return  document.getElementById(id);}"
    GetIdCodes = codes
End Function

'FindByNameのJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetNameCodes(name As String) As Variant
    Dim codes(1) As String
    codes(0) = "document.getElementsByName('" & name & "')[0];"
    codes(1) = "function(name){return  document.getElementsByName(name)[0];}"
    GetNameCodes = codes
End Function

'FindByTagのJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetTagCodes(tagName As String, Index As Long) As Variant
    Dim codes(1) As String
    codes(0) = "document.getElementsByTagName('" & tagName & "')[" & Index & "];"
    codes(1) = "function(tagName){return  document.getElementsByTagName(tagName)[" & Index & "];}"
    GetTagCodes = codes
End Function

'FindByClassNameのJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetClassNameCodes(className As String, Index As Long) As Variant
    Dim codes(1) As String
    codes(0) = "document.getElementsByClassName('" & className & "')[" & Index & "];"
    codes(1) = "function(className){return  document.getElementsByClassName(tagName)[" & Index & "];}"
    GetClassNameCodes = codes
End Function

'FindByTagのJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetTagCountCodes(tagName As String) As Variant
    Dim countCodes(1) As String
    countCodes(0) = "document.getElementsByTagName('" & tagName & "').length;"
    countCodes(1) = "function(tagName){return  document.getElementsByTagName('" & tagName & "').length;}"
    GetTagCountCodes = countCodes
End Function

'FindByClassのJavaScriptCode(通常のコードとFunction版のコード)を返す
Public Function GetClassNameCountCodes(className As String) As Variant
    Dim countCodes(1) As String
    countCodes(0) = "document.getElementsByClassName('" & className & "').length;"
    countCodes(1) = "function(className){return  document.getElementsByClassName('" & className & "').length;}"
    GetClassNameCountCodes = countCodes
End Function


'共通エラー対応。メッセージ表示、かつインターネットハンドルを解放して強制終了とする｡
'※エラー発生時にデバッグモードに入り、ユーザーに任せると解放されない可能性が高いので強制終了とする。
Public Sub DisplayErrorAndEnd(displayText As String)
    displayText = displayText & vbCrLf & "処理を終了します。"
    MsgBox displayText, vbCritical, "エラー"
    End
End Sub

Public Function GetSelectItemCode() As String
    GetSelectItemCode = _
    "function(value){value = decodeURI(value);" & _
        "for(let i = 0; i < this.options.length; i++){" & _
            "let val1 = this.options[i].value.replace(/\\r?\\n/g,'');" & _
            "let val2 = this.options[i].textContent.replace(/\\r?\\n/g,'');" & _
            "console.log('(you)' + value + ':(value)' + val1 + ':(textContent)'+ val2);" & _
            "if(val1 === value || val2 === value){" & _
                "this.options[i].selected = true;" & _
                "return true;" & _
            "}" & _
        "}" & _
    "return false;}"
End Function

Public Function GetSelectedTextContentCode() As String
    GetSelectedTextContentCode = _
    "function(){" & _
        "for(let i = 0; i < this.options.length; i++){" & _
            "if(this.options[i].selected){" & _
                "return this.options[i].textContent.replace(/\\s/g, ' ');" & _
            "}" & _
        "}" & _
    "}"
End Function

Public Function GetDispatchChangeEventCode() As String
    GetDispatchChangeEventCode = _
    "function(value){" & _
        "const event = new CustomEvent('change', {bubbles: false, cancelable: false});" & _
        "this.dispatchEvent(event);" & _
    "}"
End Function

'elementをFindした方法についての情報を配列にして返す(作成するElementに持たせるため）
Public Function MakeFindInfo(way_string As String, way As String, Optional Index As Long = -1) As Collection
    Dim findInfo As Collection: Set findInfo = New Collection
    findInfo.add way_string
    findInfo.add way
    findInfo.add Index
    Set MakeFindInfo = findInfo
End Function

'Find系メソッドで要素が見つからなかった場合のエラーメッセージを作成する
Public Function MakeFindErrorMessage(wayString As String, way As String, Optional wait As String = "") As String
    Dim errMsg As String
    errMsg = _
    "＜FindElementBy" & wayString & "＞" & vbCrLf & _
    "指定された以下" & wayString & "の要素は存在しませんでした。" & vbCrLf & _
    way & vbCrLf & "(イミディエイトウィンドウで確認可能)" & vbCrLf & vbCrLf
    
    If wayString = "XPath" Then
        errMsg = errMsg & _
        "※引数のXPathにはChromeの検証ツールで「完全なXPath」が" & vbCrLf & _
        "　 コピペできるので、それを指定してください。" & vbCrLf & vbCrLf
    End If
    
    errMsg = errMsg & _
        "※もし取得したい要素がiframe内の要素である場合、SwitchFrameを" & vbCrLf & _
        "　 使用して、フレームを切り替えてからでないと取得できません。" & vbCrLf & vbCrLf
    
    errMsg = errMsg & _
    "※上記に沿っていても、要素が画面に出現する前だと失敗します。" & vbCrLf
    
    If wait = "" Then
        errMsg = errMsg & _
        "　 必要であればWaitForElementToAppearBy" & wayString & "で" & vbCrLf
    Else
        errMsg = errMsg & _
        "　 必要であればWaitForElementToAppearBy" & wait & "で" & vbCrLf
    End If
    
    errMsg = errMsg & _
    "　 要素の出現を確認する処理を入れましょう。" & vbCrLf
    
    MakeFindErrorMessage = errMsg
End Function

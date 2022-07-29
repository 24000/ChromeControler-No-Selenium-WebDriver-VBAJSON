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

'FindByXPath��JavaScriptCode(�ʏ�̃R�[�h��Function�ł̃R�[�h)��Ԃ�
Public Function GetXpathCodes(xPath As String) As Variant
    Dim codes(1) As String
    codes(0) = "document.evaluate('" & xPath & "', document, null, " & _
                        "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;"
    codes(1) = "function(xpath){return  document.evaluate(xpath, document, null, " & _
                            "XPathResult.FIRST_ORDERED_NODE_TYPE, null ).singleNodeValue;}"
    GetXpathCodes = codes
End Function

'FindById��JavaScriptCode(�ʏ�̃R�[�h��Function�ł̃R�[�h)��Ԃ�
Public Function GetIdCodes(Id As String) As Variant
    Dim codes(1) As String
    codes(0) = "document.getElementById('" & Id & "');"
    codes(1) = "function(id){return  document.getElementById(id);}"
    GetIdCodes = codes
End Function

'FindByName��JavaScriptCode(�ʏ�̃R�[�h��Function�ł̃R�[�h)��Ԃ�
Public Function GetNameCodes(name As String) As Variant
    Dim codes(1) As String
    codes(0) = "document.getElementsByName('" & name & "')[0];"
    codes(1) = "function(name){return  document.getElementsByName(name)[0];}"
    GetNameCodes = codes
End Function

'FindByTag��JavaScriptCode(�ʏ�̃R�[�h��Function�ł̃R�[�h)��Ԃ�
Public Function GetTagCodes(tagName As String, Index As Long) As Variant
    Dim codes(1) As String
    codes(0) = "document.getElementsByTagName('" & tagName & "')[" & Index & "];"
    codes(1) = "function(tagName){return  document.getElementsByTagName(tagName)[" & Index & "];}"
    GetTagCodes = codes
End Function

'FindByTag��JavaScriptCode(�ʏ�̃R�[�h��Function�ł̃R�[�h)��Ԃ�
Public Function GetTagCountCodes(tagName As String) As Variant
    Dim countCodes(1) As String
    countCodes(0) = "document.getElementsByTagName('" & tagName & "').length;"
    countCodes(1) = "function(tagName){return  document.getElementsByTagName('" & tagName & "').length;}"
    GetTagCountCodes = countCodes
End Function


'���ʃG���[�Ή��B���b�Z�[�W�\���A���C���^�[�l�b�g�n���h����������ċ����I���Ƃ���
'���G���[�������Ƀf�o�b�O���[�h�ɓ���A���[�U�[�ɔC����Ɖ������Ȃ��\���������̂ŋ����I���Ƃ���B
Public Sub DisplayErrorAndEnd(displayText As String)
    displayText = displayText & vbCrLf & "�������I�����܂��B"
    MsgBox displayText, vbCritical, "�G���["
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

'element��Find�������@�ɂ��Ă̏���z��ɂ��ĕԂ�(�쐬����Element�Ɏ������邽�߁j
Public Function MakeFindInfo(way_string As String, way As String, Optional Index As Long = -1) As Collection
    Dim findInfo As Collection: Set findInfo = New Collection
    findInfo.add way_string
    findInfo.add way
    findInfo.add Index
    Set MakeFindInfo = findInfo
End Function

'Find�n���\�b�h�ŗv�f��������Ȃ������ꍇ�̃G���[���b�Z�[�W���쐬����
Public Function MakeFindErrorMessage(wayString As String, way As String, Optional wait As String = "") As String
    Dim errMsg As String
    errMsg = _
    "��FindElementBy" & wayString & "��" & vbCrLf & _
    "�w�肳�ꂽ�ȉ�" & wayString & "�̗v�f�͑��݂��܂���ł����B" & vbCrLf & _
    way & vbCrLf & "(�C�~�f�B�G�C�g�E�B���h�E�Ŋm�F�\)" & vbCrLf & vbCrLf
    
    If wayString = "XPath" Then
        errMsg = errMsg & _
        "��������XPath�ɂ�Chrome�̌��؃c�[���Łu���S��XPath�v��" & vbCrLf & _
        "�@ �R�s�y�ł���̂ŁA������w�肵�Ă��������B" & vbCrLf & vbCrLf
    End If
    
    errMsg = errMsg & _
        "�������擾�������v�f��iframe���̗v�f�ł���ꍇ�ASwitchFrame��" & vbCrLf & _
        "�@ �g�p���āA�t���[����؂�ւ��Ă���łȂ��Ǝ擾�ł��܂���B" & vbCrLf & vbCrLf
    
    errMsg = errMsg & _
    "����L�ɉ����Ă��Ă��A�v�f����ʂɏo������O���Ǝ��s���܂��B" & vbCrLf
    
    If wait = "" Then
        errMsg = errMsg & _
        "�@ �K�v�ł����WaitForElementToAppearBy" & wayString & "��" & vbCrLf
    Else
        errMsg = errMsg & _
        "�@ �K�v�ł����WaitForElementToAppearBy" & wait & "��" & vbCrLf
    End If
    
    errMsg = errMsg & _
    "�@ �v�f�̏o�����m�F���鏈�������܂��傤�B" & vbCrLf
    
    MakeFindErrorMessage = errMsg
End Function
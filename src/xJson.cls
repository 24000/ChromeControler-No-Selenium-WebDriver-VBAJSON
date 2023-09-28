VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private html As Object
Private Sub Class_Initialize()
    If Not Me Is xJson Then
        MsgBox "���̃N���X��Singleton�Ƃ��Đ݌v����Ă��܂��B" & vbCrLf & _
                    "�C���X�^���X�����Ȃ��ł��������B�������I�����܂��B"
        End
    End If
    
    Set html = CreateObject("htmlfile")
    html.write "<meta http-equiv='X-UA-Compatible' content='IE=edge' />"

    'URL�G���R�[�f�B���O
    html.write _
        "<script>" & vbCrLf & _
            "document.EncodeURL = " & vbCrLf & _
            "function(s){" & vbCrLf & _
            "return encodeURI(s);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    
    html.write _
        "<script>" & vbCrLf & _
            "document.DecodeURL = " & vbCrLf & _
            "function(s){" & vbCrLf & _
            "return decodeURI(s);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
'JSON����
'============================
    '�z�񂩂�I�u�W�F�N�g���o��
    html.write _
        "<script>" & vbCrLf & _
            "document.getObjectFromArrayBySetIndex = " & vbCrLf & _
            "function(targetJson,index){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return JSON.stringify(vals[index]);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    '�I�u�W�F�N�g����I�u�W�F�N�g�����o��
    html.write _
        "<script>" & vbCrLf & _
            "document.getObjectFromObjectBySetKey = " & vbCrLf & _
            "function(targetJson,key){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return JSON.stringify(vals[key]);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    '�z�񂩂�l�����o��
    html.write _
        "<script>" & vbCrLf & _
            "document.getValueFromArrayBySetIndex = " & vbCrLf & _
            "function(targetJson,index){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return vals[index];" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    '�I�u�W�F�N�g����l�����o��
    html.write _
        "<script>" & vbCrLf & _
            "document.getValueFromObjectBySetKey = " & vbCrLf & _
            "function(targetJson,key){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return vals[key];" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
        
    html.write _
        "<script>" & vbCrLf & _
            "document.getJsonLength = " & vbCrLf & _
            "function(targetJson){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return vals.length;" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
End Sub

Public Function EncodeURL(json As String) As String
    EncodeURL = html.EncodeURL(json)
End Function

Public Function DecodeURL(json As String) As String
    DecodeURL = html.DecodeURL(json)
End Function

Public Function GetValue(ByVal json As String, ParamArray argsArray() As Variant) As String
    Dim i As Long
    Dim result As String
    For i = LBound(argsArray) To UBound(argsArray)
        If json = "" Then
            GetValue = result
            Exit Function
        End If
        If i <> UBound(argsArray) Then
            If IsNumeric(argsArray(i)) Then
                json = html.getObjectFromArrayBySetIndex(json, argsArray(i))
            Else
                json = html.getObjectFromObjectBySetKey(json, argsArray(i))
            End If
        Else
            If IsNumeric(argsArray(i)) Then
                result = html.getValueFromArrayBySetIndex(json, argsArray(i))
            Else
                result = html.getValueFromObjectBySetKey(json, argsArray(i))
            End If
        End If
    Next
    GetValue = result
End Function

Public Function GetObject(ByVal json As String, ParamArray argsArray() As Variant) As String
    Dim i As Long
    Dim result As String
    For i = LBound(argsArray) To UBound(argsArray)
        If json = "" Then
            GetObject = json
            Exit Function
        End If
     
        If IsNumeric(argsArray(i)) Then
            json = html.getObjectFromArrayBySetIndex(json, argsArray(i))
        Else
            json = html.getObjectFromObjectBySetKey(json, argsArray(i))
        End If
    Next
    GetObject = json
End Function


Public Function GetJsonLength(json As String) As Long
    GetJsonLength = html.GetJsonLength(json)
End Function
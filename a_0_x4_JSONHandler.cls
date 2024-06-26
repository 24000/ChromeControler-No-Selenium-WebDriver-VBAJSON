VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "a_0_x4_JSONHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private html As Object
Private Sub Class_Initialize()
    Set html = CreateObject("htmlfile")
    html.write "<meta http-equiv='X-UA-Compatible' content='IE=edge' />"

    'URLエンコーディング
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
'JSON操作
'============================
    '配列からオブジェクト取り出す
    html.write _
        "<script>" & vbCrLf & _
            "document.getObjectFromArrayBySetIndex = " & vbCrLf & _
            "function(targetJson,index){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return JSON.stringify(vals[index]);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    'オブジェクトからオブジェクトを取り出す
    html.write _
        "<script>" & vbCrLf & _
            "document.getObjectFromObjectBySetKey = " & vbCrLf & _
            "function(targetJson,key){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return JSON.stringify(vals[key]);" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    '配列から値を取り出す
    html.write _
        "<script>" & vbCrLf & _
            "document.getValueFromArrayBySetIndex = " & vbCrLf & _
            "function(targetJson,index){" & vbCrLf & _
            "var vals = JSON.parse(targetJson);" & vbCrLf & _
            "return vals[index];" & vbCrLf & _
            "}" & vbCrLf & _
        "</script>"
    'オブジェクトから値を取り出す
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

Public Function EncodeURL(Json As String) As String
    EncodeURL = html.EncodeURL(Json)
End Function

Public Function DecodeURL(Json As String) As String
    DecodeURL = html.DecodeURL(Json)
End Function

Public Function GetValueFrom(ByVal Json As String, argsArray As Variant) As String
    Dim i As Long
    Dim result As String
    For i = LBound(argsArray) To UBound(argsArray)
        If Json = "" Then
            GetValueFrom = result
            Exit Function
        End If
        If i <> UBound(argsArray) Then
            If IsNumeric(argsArray(i)) Then
                Json = html.getObjectFromArrayBySetIndex(Json, argsArray(i))
            Else
                Json = html.getObjectFromObjectBySetKey(Json, argsArray(i))
            End If
        Else
            If IsNumeric(argsArray(i)) Then
                result = html.getValueFromArrayBySetIndex(Json, argsArray(i))
            Else
                result = html.getValueFromObjectBySetKey(Json, argsArray(i))
            End If
        End If
    Next
    GetValueFrom = result
End Function

Public Function GetObjectFrom(ByVal Json As String, argsArray As Variant) As String
    Dim i As Long
    Dim result As String
    For i = LBound(argsArray) To UBound(argsArray)
        If Json = "" Then
            GetObjectFrom = Json
            Exit Function
        End If
     
        If IsNumeric(argsArray(i)) Then
            Json = html.getObjectFromArrayBySetIndex(Json, argsArray(i))
        Else
            Json = html.getObjectFromObjectBySetKey(Json, argsArray(i))
        End If
    Next
    GetObjectFrom = Json
End Function


Public Function GetJsonLength(Json As String) As Long
    GetJsonLength = html.GetJsonLength(Json)
End Function

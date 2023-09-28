VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Global定数クラス（Singleton)


Private Sub Class_Initialize()
    If Not Me Is xG Then
        MsgBox "このクラスはSingletonとして設計されています。" & vbCrLf & _
                    "インスタンス化しないでください。処理を終了します。"
        End
    End If
End Sub

Public Property Get CHROME_EXE_REG_KEY() As String
    CHROME_EXE_REG_KEY = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\"
End Property
Public Property Get DEBUG_MODE_OPTION() As String
    DEBUG_MODE_OPTION = " --remote-debugging-Port=9222"
End Property
Public Property Get USER_DATA_DIR_OPTION() As String
    USER_DATA_DIR_OPTION = " --user-data-dir="
End Property
Public Property Get ENABLE_AUTOMATION_OPTION() As String
    ENABLE_AUTOMATION_OPTION = " --enable-automation"
End Property
Public Property Get HOME_PAGE() As String
    HOME_PAGE = " --homepage=https://google.com"
End Property

Public Property Get LoopBackAddress() As String
    LoopBackAddress = "http://127.0.0.1"
End Property
Public Property Get LoopBackAddressOnly() As String
    LoopBackAddressOnly = "127.0.0.1"
End Property
Public Property Get PORT() As String
    PORT = "9222"
End Property

Public Property Get AGENT() As String
    AGENT = "Client"
End Property
Public Property Get GET_METHOD() As String
    GET_METHOD = "GET"
End Property
Public Property Get HTTP_VERSION() As String
    HTTP_VERSION = "HTTP/1.1"
End Property
Public Property Get TARGET_LIST_PATH() As String
    TARGET_LIST_PATH = "/json"
End Property


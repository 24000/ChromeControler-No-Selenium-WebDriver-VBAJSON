VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebDriver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub URLOpen(url As String, Optional maxWaitTime_ms As Long = 30000)

End Sub

Public Function ExecuteJavaScript(Script As String) As String

End Function

Public Function GetElementByXpath(xpath As String) As IWebElement

End Function

Public Sub SwitchFrameByXpath(xpath As String)

End Sub

Public Sub SwitchFrameToDefault()

End Sub

Public Function IsExistsElementByXpath(xpath As String) As Boolean

End Function

Public Sub WaitForElementToAppearByXpath(xpath As String, maxWaitTime_ms As Long)

End Sub

Public Sub PushOkButtonOnDialog()

End Sub

Public Sub SleepByWindowsAPI(milliSecond As Long)

End Sub

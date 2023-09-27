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
'

Public Property Get URL() As String

End Property

Public Sub OpenURL(targetURL As String)

End Sub

Public Function ExecuteScript(script As String) As String

End Function

Public Function FindElementById(Id As String) As IWebElement

End Function

Public Function FindElementByXPath(xPath As String) As IWebElement

End Function

Public Function FindElementByName(name As String) As IWebElement

End Function

Public Function FindElementsByTag(tagName As String) As IWebElements

End Function

Public Function FindElementsByClassName(className As String) As IWebElements

End Function

Public Function IsExistsElementByXpath(xPath As String) As Boolean

End Function

Public Function IsExistsElementById(Id As String) As Boolean

End Function

Public Function IsExistsElementByName(name As String) As Boolean

End Function

Public Sub WaitForElementToAppearByXpath(xPath As String, maxWaitTime_sec As Long)

End Sub

Public Sub WaitForElementToAppearById(Id As String, maxWaitTime_sec As Long)

End Sub

Public Sub WaitForElementToAppearByName(name As String, maxWaitTime_sec As Long)

End Sub

Public Sub HandleAlertDialog()

End Sub

Public Sub SleepByWinAPI(milliSecond As Long)

End Sub

Public Sub SwitchTabByIndex(Index As Long)
    
End Sub

Public Sub SwitchFrameByNameOrURLOrIndex(value As Variant)

End Sub

Public Sub SwitchFrameByIframeElement(IframeElement As IWebElement)

End Sub

Public Sub SwitchFrameToDefault()

End Sub

Public Sub DebugPrintAccessibleFrameInfo()

End Sub

Public Sub CloseWindow()
    
End Sub

Public Sub ScrollWindow(yPixel As Long, Optional xPixel As Long = 0)

End Sub

Public Sub SetForeGroundWindow()

End Sub

Public Sub AllowPopUpWindowForURL(targetURL As String)

End Sub

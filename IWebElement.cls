VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWebElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get Checked() As Boolean

End Property
Public Property Let Checked(bool As Boolean)

End Property


Public Sub SetText(text As String)

End Sub

Public Sub SetTextContent(text As String)

End Sub

Public Sub Click()

End Sub

Public Sub SimulateClick()

End Sub

Public Sub SelectItemByText(text As String)

End Sub

Public Function GetElementByName(Name As String) As IWebElement

End Function

Public Function GetElementsByTagName(tagName As String) As IWebElements

End Function

Public Function GetTextContent() As String

End Function

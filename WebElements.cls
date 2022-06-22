VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IWebElements

Private elements_ As Collection

Public Sub Init(elements As Collection)
    Set elements_ = elements
End Sub

Private Function IWebElements_Count() As Long
    IWebElements_Count = elements_.Count
End Function

Private Function IWebElements_Item(index As Variant) As IWebElement
    Set IWebElements_Item = elements_.Item(index)
End Function

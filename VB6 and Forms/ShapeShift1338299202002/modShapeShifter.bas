Attribute VB_Name = "Module1"
Public Declare Function CreatePolygonRgn Lib "gdi32" _
    (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode _
    As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" _
 (ByVal hWnd As Long, ByVal hRgn As Long, _
 ByVal bRedraw As Boolean) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type


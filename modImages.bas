Attribute VB_Name = "modImages"
'image processing subs
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Const STRETCH_ANDSCANS = 1
Public Const STRETCH_DELETESCANS = 3
Public Const STRETCH_HALFTONE = 4
Public Const STRETCH_ORSCANS = 2
Public Const StretchBltC = 2048

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public REF_BT As Byte

Public Function BreakColor(ByVal Color As Long, ByRef Red As Byte, ByRef Green As Byte, ByRef Blue As Byte) As Long

    Dim byt_Colors(1 To 3) As Byte
    

    If Color > -1 Then
        CopyMemory ByVal VarPtr(byt_Colors(1)), Color, 3
        Red = byt_Colors(1)
        Green = byt_Colors(2)
        Blue = byt_Colors(3)
        BreakColor = 1
    End If

    
End Function

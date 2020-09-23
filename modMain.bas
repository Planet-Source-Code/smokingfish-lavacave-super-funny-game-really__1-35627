Attribute VB_Name = "modMain"
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As RasterOpConstants) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function SetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const MERGEPAINT = &HBB0226
Public Const SRCCOPY = &HCC0020
Public Game As New clsMain

Function TwipsToPixelsX(Twips As Variant) As Variant
   TwipsToPixelsX = Twips / Screen.TwipsPerPixelX
End Function

Function PixelsToTwipsX(Pixels As Variant) As Variant
   PixelsToTwipsX = Pixels * Screen.TwipsPerPixelX
End Function

Function TwipsToPixelsY(Twips As Variant) As Variant
   TwipsToPixelsY = Twips / Screen.TwipsPerPixelY
End Function

Function PixelsToTwipsY(Pixels As Variant) As Variant
   PixelsToTwipsY = Pixels * Screen.TwipsPerPixelY
End Function

Public Function FileExists(FullFileName As String) As Boolean
    On Error GoTo MakeF
        Open FullFileName For Input As #1
        Close #1
        FileExists = True
    Exit Function
MakeF:
        FileExists = False
    Exit Function
End Function

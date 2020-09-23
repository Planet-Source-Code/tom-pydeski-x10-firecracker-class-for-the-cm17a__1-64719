Attribute VB_Name = "modRotatePic"
'
' Created by E.Spencer (elliot@spnc.demon.co.uk) - This code is public domain.
'
Option Explicit
Global Const Pi = 3.14159265359
Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Const DI_NORMAL As Long = 3
Public Const SRCCOPY  As Long = &HCC0020
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'
Public StartFlag As Boolean
Public ImageHandle As Long
Public BMPHandle As Long
Public BMPhDc As Long
Public ObjHandle As Long
Public L1 As Long
Public MyBMP As BITMAP
Public PicRect As RECT
Public Theta As Single

Sub InitPic(picSource As PictureBox)
'usage
'InitPic picSource
'DoEvents
'n = 360 - ((2 / 12) * 360) 'position for 2 O'clock
'Theta = (6.28 / 360) * n   ' Convert degrees to theta value
'RotatePic Theta, picSource, picDest
'
' Now I want a copy of the bitmap privately saved so I don't have to do this work again.
L1 = GetWindowRect(picSource.hwnd, PicRect)
' Get src Rectangle dimensions
BMPHandle = CreateCompatibleBitmap(picSource.hdc, (PicRect.Right - PicRect.Left + 1), (PicRect.Bottom - PicRect.Top + 1))
' Get handle to bitmap
L1 = GetObject(BMPHandle, Len(MyBMP), MyBMP)
' Need a new device context compatible with a picture control
BMPhDc = CreateCompatibleDC(picSource.hdc)
' Assign the bitmap to the device context
ObjHandle = SelectObject(BMPhDc, BMPHandle)
' Copy the bitmap into our new bmp memory save area
L1 = BitBlt(BMPhDc, 0, 0, picSource.ScaleWidth, picSource.ScaleHeight, picSource.hdc, 0, 0, SRCCOPY)
' Set flag to show we've done this operation
StartFlag = True
End Sub

Sub RotatePic(Theta As Single, picSource As PictureBox, picDest As PictureBox)
'usage
'InitPic picSource
'n=degrees to rotate
'Convert degrees to theta value
'Theta = (6.28 / 360) * n
'RotatePic Theta, picSource, picDest
'ClosePic
'
Dim c1x As Integer
Dim c1y As Integer
Dim c2x As Integer
Dim c2y As Integer
Dim p1x As Integer
Dim p1y As Integer
Dim p2x As Integer
Dim p2y As Integer
Dim a As Single
Dim n As Integer
Dim r As Integer
Dim P1Hdc As Long
Dim P2Hdc As Long
picDest.Cls
' Get the device context - saves time in loop
P1Hdc = picSource.hdc
P2Hdc = picDest.hdc
c1x = picSource.ScaleWidth \ 2
c1y = picSource.ScaleHeight \ 2
c2x = picDest.ScaleWidth \ 2
c2y = picDest.ScaleHeight \ 2
Dim c0 As Long
Dim c1 As Long
Dim c2 As Long
Dim c3 As Long
Dim xret As Long
If c2x < c2y Then n = c2y Else n = c2x
n = n - 1
For p2x = 0 To n
    For p2y = 0 To n
        If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x)
        r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
        p1x = r * Cos(a + Theta)
        p1y = r * Sin(a + Theta)
        c0 = GetPixel(P1Hdc, c1x + p1x, c1y + p1y)
        c1 = GetPixel(P1Hdc, c1x - p1x, c1y - p1y)
        c2 = GetPixel(P1Hdc, c1x + p1y, c1y - p1x)
        c3 = GetPixel(P1Hdc, c1x - p1y, c1y + p1x)
        If c0 <> -1 Then xret = SetPixel(P2Hdc, c2x + p2x, c2y + p2y, c0)
        If c1 <> -1 Then xret = SetPixel(P2Hdc, c2x - p2x, c2y - p2y, c1)
        If c2 <> -1 Then xret = SetPixel(P2Hdc, c2x + p2y, c2y - p2x, c2)
        If c3 <> -1 Then xret = SetPixel(P2Hdc, c2x - p2y, c2y + p2x, c3)
    Next
Next
End Sub

Sub ClosePic()
' Get rid of our saved BMP device context
If StartFlag = True Then
    L1 = DeleteObject(ObjHandle)
    L1 = DeleteDC(BMPhDc)
End If
End Sub

Public Function ArcTangent(p_dblVal As Double) As Double
' Comments :
' Parameters: p_dblVal -
' Returns: Double -
' Modified :
'
' -------------------------
'Radian Input Degree Output
On Error GoTo PROC_ERR
Dim dblPi As Double
Dim dblDegree As Double
' xx Calculate the value of Pi.
dblPi = 4 * Atn(1)
' xx To convert radians to degrees,
' multiply radians by 180/pi.
dblDegree = 180 / dblPi
p_dblVal = Val(p_dblVal)
ArcTangent = Atn(p_dblVal) * dblDegree
PROC_EXIT:
Exit Function
PROC_ERR:
ArcTangent = 0
MsgBox err.Description, vbExclamation
Resume PROC_EXIT
End Function

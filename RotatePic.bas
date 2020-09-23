Attribute VB_Name = "RotatePic"

'**************************************
' Name: A Very Fast Rotate (Any Degree)
' Description:SORRY!! (FIXED) Rotates An
'     Image Any Angle With GetBitmapBits And S
'     etBitmapBits Very Fast (PLEASE IF YOU TE
'     ST THIS AND IT DOESNT WORK EMAIL ME)
' By: Show
'
' Assumes:Put 2 PictureBoxs
'Picture1 AutoSize = True
'ScaleMode = Pixel
'Picture2 Same Thing
'CommandButton
'Put This Code In It:
'RotateAnyAngle Picture1, Picture2, 45, True 'Clockwise
'RotateAnyAngle Picture1, Picture2, 45, False 'CounterClockwise
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=51162&lngWId=1'for details.
'**************************************
Declare Sub GetBitmapBits Lib "GDI32" (ByVal hBitmap As Long, ByVal nwCount As Long, lpBits As Any)
Declare Sub SetBitmapBits Lib "GDI32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any)

Sub RotateAnyAngle(Pic As PictureBox, Pic1 As PictureBox, A As Double, ClockWise As Long)
Dim X&, Y&, SA!, CA!, nX&, nY&
Dim sW&, sH&, dH&, dW&, dW2&, dH2&, cH&, cW&
Const Pi = 0.017453292519943
If A < 0 Then A = 0
If A > 360 Then A = 360
If ClockWise = False Then A = 360 - A
CA = Cos(A * Pi * -1)
SA = Sin(A * Pi * -1)
sH = Pic.ScaleHeight
sW = Pic.ScaleWidth
cW = sW / 2
cH = sH / 2
dH = Pic1.ScaleHeight
dW = Pic1.ScaleWidth
dW2 = dW / 2
dH2 = dH / 2
ReDim SrcBits(1 To sW, 1 To sH) As Integer
ReDim DesBits(1 To dW, 1 To dH) As Integer
'VB 32 Users If This Doesnt Work Try
'ReDim SrcBits(1 To Pic.ScaleWidth, 1 To
'     Pic.ScaleHeight) As Byte
'ReDim DesBits(1 To Pic1.ScaleWidth, 1 T
'     o Pic1.ScaleHeight) As Byte
Call GetBitmapBits(Pic.Image, (2 * sW) * sH, SrcBits(1, 1))
For Y = 1 To Pic1.ScaleHeight
    For X = 1 To Pic1.ScaleWidth
        nX = CA * (X - dW2) - SA * (Y - dH2) + cW
        nY = SA * (X - dW2) + CA * (Y - dH2) + cH
        If X < dW And Y < dH And nY > 0 And nX > 0 And nY < sH And nX < sW Then
            DesBits(X, Y) = SrcBits(nX, nY)
        End If
    Next X
Next Y
Call SetBitmapBits(Pic1.Image, (2 * dW) * dH, DesBits(1, 1))
Pic1.Refresh
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsX10CM17A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Tom Pydeski's X10 Firecracker palmpad
'This code was originally in source code for a keware ocx.
'I have modified it to utilize it in a class module instead.

' this code is public domain and distribution is unlimited
' Keware acknowledges the work of the following authors:
' Bruce Lawson for figuring out the CM17A protocol
' Bill Jones for some sample VB code
Option Explicit
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
'
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Dim UseComport As Integer
Dim gComOpen As Boolean
Private Comm1 As MSCommLib.MSComm
Private mHouseCode As Integer
Private mDeviceCode As Integer
Private UnitTable(16)
Private HouseTable(16)
'these are the events that are generated in our class file and
'are accessed by the calling module
Public Event Initialized()
'
Private eTitle$
Private eMess$
Private mError As Long

Private Sub Class_Initialize()
Set Comm1 = New MSCommLib.MSComm
'below are the possible events
' 0 =All Units Off
' 1 =All Lights On
' 2 =On
' 3 =Off
' 4 =Dim
' 5 =Bright
' 6 =All Lights Off
' 7 =Extended
' 8 =Hail Request
' 9 =Hail Ack
' 10 =Pre-set dim1
' 11 =Pre-set dim2
' 12 =Extended Data
' 13 =Status On
' 14 =Status Off
' 15 =Status Request
' 16 =-1
'
X10Command$(0) = "All Units Off"
X10Command$(1) = "All Lights On"
X10Command$(2) = "On"
X10Command$(3) = "Off"
X10Command$(4) = "Dim"
X10Command$(5) = "Bright"
X10Command$(6) = "All Lights Off"
X10Command$(7) = "Extended"
X10Command$(8) = "Hail Request"
X10Command$(9) = "Hail Ack"
X10Command$(10) = "Pre-set Dim1"
X10Command$(11) = "Pre-set Dim2"
X10Command$(12) = "Extended Data"
X10Command$(13) = "Status On"
X10Command$(14) = "Status Off"
X10Command$(15) = "Status Request"
X10Command$(16) = "-1"
'  Look at FireCracker datasheet for reference
UnitTable(0) = &H0
UnitTable(1) = &H10
UnitTable(2) = &H8
UnitTable(3) = &H18
UnitTable(4) = &H40
UnitTable(5) = &H50
UnitTable(6) = &H48
UnitTable(7) = &H58
'this code will represent what we need to transmit for a given housecode
HouseTable(0) = &H60
HouseTable(1) = &H70
HouseTable(2) = &H40
HouseTable(3) = &H50
HouseTable(4) = &H80
HouseTable(5) = &H90
HouseTable(6) = &HA0
HouseTable(7) = &HB0
HouseTable(8) = &HE0
HouseTable(9) = &HF0
HouseTable(10) = &HC0
HouseTable(11) = &HD0
HouseTable(12) = &H0
HouseTable(13) = &H10
HouseTable(14) = &H20
HouseTable(15) = &H30
End Sub

Private Sub Class_Terminate()
'MsgBox "terminate", vbOKOnly
If Comm1.PortOpen = True Then
    ResetCom
End If
End Sub

Public Property Get ComPort() As Integer
ComPort = UseComport
End Property

Public Property Let ComPort(ByVal newport As Integer)
UseComport = newport
End Property

Public Property Get HouseCode() As Integer
HouseCode = mHouseCode
End Property

Public Property Let HouseCode(ByVal NewHouse As Integer)
mHouseCode = NewHouse
End Property

Public Property Get DeviceCode() As Integer
DeviceCode = mDeviceCode
End Property

Public Property Let DeviceCode(ByVal NewDevice As Integer)
mDeviceCode = NewDevice
End Property

Sub WaitMicroSecs(ms As Long)
Dim s As LARGE_INTEGER
Dim e As LARGE_INTEGER
Dim total
Dim err As Long
total = ms / 0.8
err = QueryPerformanceCounter(s)
Do
    err = QueryPerformanceCounter(e)
Loop While e.LowPart < s.LowPart + total
'Debug.Print Str(s.highpart) + " " + Str(s.lowpart)
'Debug.Print Str(e.highpart) + " " + Str(e.lowpart)
End Sub

Public Sub ResetCom()
' simply closes the com port
Comm1.PortOpen = False
End Sub

Public Function Init() As Integer
On Error GoTo Oops
X10Init = 0
Comm1.CommPort = UseComport
' 4800 baud, no parity, 8 data, and 1 stop bit.
Comm1.Settings = "4800,N,8,1"
' Tell the control to read entire buffer when Input is used.
Comm1.InputLen = 0
' Open the port.
Comm1.PortOpen = True
Init = 0
gComOpen = True
RaiseEvent Initialized
X10Init = 1
GoTo Exit_Init
Oops:
gComOpen = False
Init = err.Number
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Init "
eMess$ = "Error # " & err.Number & " - " & err.Description & vbCrLf
eMess$ = eMess$ & "Occurred in Init"
eMess$ = eMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(eMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Init:
End Function

Public Sub Exec(HouseCodeStr As String, DeviceCodeStr As String, Command As Integer, Optional Brightness As Integer)
' 2=60 = A
' 2=70 = B
' 3=98 = dim
' 3=88 = bright
' 3=10=1 on
' 3=18=3 on
'--------------------------------------------------------------------------
'Each transmission consists of 40 bits. This is made up of the following:
' Header: 16 bits
' Data: 16 bits
' Footer: 8 bits
'The header and footer never change and are defined as follows:
' Header: 11010101 10101010 => &HD5 &HAA
' Footer: 10101101 => &HAD
'--------------------------------------------------------------------------
Dim u As Integer
Dim buf(5) As Byte
Dim i As Integer
If Not gComOpen Then Exit Sub
HouseCodeStr = UCase(HouseCodeStr)
HouseCode = Asc(HouseCodeStr) - Asc("A") '0-15
DeviceCode = Val(DeviceCodeStr)
'set the 2 byte header
buf(0) = &HD5 '1101 0101 = 213
buf(1) = &HAA '1010 1010 = 170
'set the footer
buf(4) = &HAD
' handle housecode
buf(2) = 0
'Select Case HouseCodeStr
'    Case "A"
'        buf(2) = 6
'    Case "B"
'        buf(2) = 7
'    Case "C"
'        buf(2) = 4
'    Case "D"
'        buf(2) = 5
'    Case "E"
'        buf(2) = 8
'    Case "F"
'        buf(2) = 9
'    Case "G"
'        buf(2) = 10
'    Case "H"
'        buf(2) = 11
'    Case "I"
'        buf(2) = 14
'    Case "J"
'        buf(2) = 15
'    Case "K"
'        buf(2) = 12
'    Case "L"
'        buf(2) = 13
'    Case "M"
'        buf(2) = 0
'    Case "N"
'        buf(2) = 1
'    Case "O"
'        buf(2) = 2
'    Case "P"
'        buf(2) = 3
'    Case Else
'        buf(2) = 1
'End Select
'move to upper nibble
'buf(2) = buf(2) * 16
'above changed
'lets try a new way
buf(2) = HouseTable(HouseCode)
' handle unit code
'If D16 is a "1" the code is a function (command).
'Otherwise D1-D8 is an 'address'.
'Once addressed, a module responds to any command code.
'It becomes 'unaddressed' by the first 'address' message after a command,
'or by 'All Units Off'.
'it seems the firecracker is different from below
'It actually appears to mirror the below table, which is for cm11
'Unit           D8  D4  D2  D1     Hex
'1               0   1   1   0      6h
'2               1   1   1   0      Eh
'3               0   0   1   0      2h
'4               1   0   1   0      Ah
'5               0   0   0   1      1h
'6               1   0   0   1      9h
'7               0   1   0   1      5h
'8               1   1   0   1      Dh
'9               0   1   1   1      7h
'10              1   1   1   1      Fh
'11              0   0   1   1      3h
'12              1   0   1   1      Bh
'13              0   0   0   0      0h
'14              1   0   0   0      8h
'15              0   1   0   0      4h
'16              1   1   0   0      Ch
u = Val(DeviceCodeStr)
buf(3) = 0
If (u - 1) And 2 Then
    buf(3) = buf(3) Or &H8
End If
If (u - 1) And 1 Then
    buf(3) = buf(3) Or &H10
End If
If (u - 1) And 4 Then
    buf(3) = buf(3) Or &H40
End If
If (u - 1) And &H8 Then
    buf(2) = buf(2) Or &H4
End If
If Command = C_OFF Then
    'off adds 20h to the 2nd data byte
    buf(3) = buf(3) Or &H20
End If
'below is from a web page on x10, but it makes no sense to me...
'0  0  0  x  x     House-Unit "On"
'0  0  1  x  x     House-Unit "Off"
'0  1  0  0  0     House "All Units Off"
'0  1  1  0  0     House "All Lights Off"
'0  1  x  1  0     House "All Lights On"
'0  1  x  0  1     House "Bright 005"
'0  1  x  1  1     House "Dim 005"
'
'all lights on = command &h90?
'all lights/units off = command &ha0?
'
Select Case Command
    'I'm guessing at the all on and all off, they seem to work
    Case ALL_LIGHTS_ON
        buf(3) = &H90
        SendToCM17 buf
    Case ALL_LIGHTS_OFF
        buf(3) = &H80
        SendToCM17 buf
    Case C_ON, C_OFF
        ' commands are ready, send it
        SendToCM17 buf
    Case C_DIM, C_BRIGHT
        'To brighten or dim a lamp, it is necessary to send consecutive signals (each
        'equivalent to a 5% dim or brighten) until the desired level is attained.
        '
        ' address device with an ON if device passed
        If DeviceCodeStr <> "" Then
            ' must have a device, send an ON
            'SendToCM17 buf
        End If
        If Brightness > 0 And LastDim(DeviceCode) = 0 Then
            ' must have a device, send an ON
            Debug.Print "Sending an on command..."
            SendToCM17 buf
            Sleep 500
        End If
        ' send the dim command
        If Command = C_DIM Then
            buf(3) = &H98
            Debug.Print "Dim ";
            'we must invert the brightness to be dimness
            If Brightness > 0 Then
                Brightness = 100 - Brightness
            End If
            Debug.Print "Dim ";
        Else
            buf(3) = &H88
            Debug.Print "Bright ";
        End If
        buf(2) = buf(2) And Not &H4
        ' convert % to number of dims
        ' looks like cm17 has 8 dim levels so we must increment by 12.5 %
        'note:changing the brightness changes the calling parameter
        If Brightness > 0 And Brightness <= 100 Then
            Brightness = (Brightness / 100) * 8
            If Brightness < 1 Then Brightness = 1
            'SendToCM17 buf
            Debug.Print "Level="; Brightness; ": "; ' 1; " ";
            For i = 1 To Brightness
                SendToCM17 buf
                Sleep 500
                Debug.Print i; " ";
            Next i
            Debug.Print
        ElseIf IsEmpty(Brightness) = True Or Brightness = 0 Then
            'just send one dim/bright
            SendToCM17 buf
            Debug.Print " once"
        End If
End Select
Debug.Print HouseCodeStr; DeviceCodeStr; " "; X10Command$(Command); " => ",
'don't show 0 or 1 or 4 as the headers and footers are fixed
Debug.Print Hex$(buf(2)); " "; Hex$(buf(3))
'example of device table
'A1 ON           0110 0000   0000 0000    60 00
'A1 OFF          0110 0000   0010 0000    60 20
'A2 ON           0110 0000   0001 0000    60 10
'A2 OFF          0110 0000   0011 0000    60 30
'A3 ON           0110 0000   0000 1000    60 08
'A3 OFF          0110 0000   0010 1000    60 28
'A4 ON           0110 0000   0001 1000    60 18
'A4 OFF          0110 0000   0011 1000    60 38
'A5 ON           0110 0000   0100 0000    60 40
'A5 OFF          0110 0000   0110 0000    60 60
'A6 ON           0110 0000   0101 0000    60 50
'A6 OFF          0110 0000   0111 0000    60 70
'A7 ON           0110 0000   0100 1000    60 48
'A7 OFF          0110 0000   0110 1000    60 68
'A8 ON           0110 0000   0101 1000    60 58
'A8 OFF          0110 0000   0111 1000    60 78
'A9 ON           0110 0100   0000 0000    64 00
'A9 OFF          0110 0100   0010 0000    64 20
'A10 ON          0110 0100   0001 0000    64 10
'A10 OFF         0110 0100   0011 0000    64 30
'A11 ON          0110 0100   0000 1000    64 08
'A11 OFF         0110 0100   0010 1000    64 28
'A12 ON          0110 0100   0001 1000    64 18
'A12 OFF         0110 0100   0011 1000    64 38
'A13 ON          0110 0100   0100 0000    64 40
'A13 OFF         0110 0100   0110 0000    64 60
'A14 ON          0110 0100   0101 0000    64 50
'A14 OFF         0110 0100   0111 0000    64 70
'A15 ON          0110 0100   0100 1000    64 48
'A15 OFF         0110 0100   0110 1000    64 68
'A16 ON          0110 0100   0101 1000    64 58
'A16 OFF         0110 0100   0111 1000    64 78
'A BRIGHT 005    0110 0000   1000 1000    64 88
'A DIM 005       0110 0000   1001 1000    64 98
'--------------------------------------------------------------------------
End Sub

Sub SendToCM17(buf() As Byte)
Dim b As Byte
Dim k As Integer
Dim i As Integer
Dim j As Integer
For k = 1 To 2
    For i = 0 To 4
        b = buf(i)
        For j = 7 To 0 Step -1
            If b And 2 ^ j Then
                LineHigh
            Else
                LineLow
            End If
        Next j
    Next i
Next k
' wait between blocks (how much? less than 6000 millisecs does not work)
WaitMicroSecs 600000
End Sub

Private Sub LineLow()
Comm1.RTSEnable = False
WaitMicroSecs 500
Comm1.RTSEnable = True
End Sub

Private Sub LineHigh()
Comm1.DTREnable = False
WaitMicroSecs 500
Comm1.DTREnable = True
End Sub


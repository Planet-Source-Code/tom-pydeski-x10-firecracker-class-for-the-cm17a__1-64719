VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCM17A 
   ClientHeight    =   6555
   ClientLeft      =   825
   ClientTop       =   1365
   ClientWidth     =   6510
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmCM17A.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6555
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbHouseCode 
      Height          =   360
      Left            =   4200
      TabIndex        =   44
      Text            =   "HC"
      Top             =   360
      Width           =   735
   End
   Begin VB.ComboBox cmbDeviceCode 
      Height          =   360
      ItemData        =   "frmCM17A.frx":074A
      Left            =   4200
      List            =   "frmCM17A.frx":074C
      TabIndex        =   43
      Text            =   "DC"
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox cmbX10Com 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCM17A.frx":074E
      Left            =   5160
      List            =   "frmCM17A.frx":0785
      Style           =   2  'Dropdown List
      TabIndex        =   42
      Top             =   1000
      Width           =   1335
   End
   Begin VB.VScrollBar scrDim 
      Height          =   495
      LargeChange     =   10
      Left            =   6120
      Max             =   0
      Min             =   100
      SmallChange     =   5
      TabIndex        =   32
      Top             =   480
      Value           =   5
      Width           =   255
   End
   Begin VB.CommandButton cmdAllOff 
      Appearance      =   0  'Flat
      Caption         =   "All Off"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4200
      TabIndex        =   14
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdAllOn 
      Appearance      =   0  'Flat
      Caption         =   "All On"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4200
      TabIndex        =   13
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Init 
      Caption         =   "Init"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox TxtCom 
      Height          =   285
      Left            =   5400
      TabIndex        =   9
      Text            =   "1"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtDim 
      Height          =   285
      Left            =   5520
      TabIndex        =   6
      Text            =   "5"
      Top             =   600
      Width           =   580
   End
   Begin VB.CommandButton ButBRIGHT 
      Appearance      =   0  'Flat
      Caption         =   "BRIGHT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5400
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton ButDIM 
      Appearance      =   0  'Flat
      Caption         =   "DIM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5400
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton ButOff 
      Appearance      =   0  'Flat
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4200
      TabIndex        =   3
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton ButOn 
      Appearance      =   0  'Flat
      Caption         =   "ON"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox picPalm 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6540
      Left            =   0
      Picture         =   "frmCM17A.frx":0845
      ScaleHeight     =   6480
      ScaleWidth      =   3990
      TabIndex        =   11
      Top             =   0
      Width           =   4050
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   580
         Index           =   8
         Left            =   1500
         TabIndex        =   41
         Top             =   4440
         Width           =   200
         _ExtentX        =   344
         _ExtentY        =   1032
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.PictureBox picPointer 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   0
         Picture         =   "frmCM17A.frx":54E87
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   7
         TabIndex        =   16
         Top             =   6000
         Width           =   105
      End
      Begin VB.PictureBox picHouseCode 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1110
         Left            =   390
         Picture         =   "frmCM17A.frx":55229
         ScaleHeight     =   74
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   74
         TabIndex        =   15
         ToolTipText     =   "Click to select the House Code"
         Top             =   4980
         Width           =   1110
      End
      Begin VB.TextBox txtDeviceName 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   750
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   31
         Text            =   "frmCM17A.frx":5932B
         ToolTipText     =   "Name of the device"
         Top             =   4440
         Width           =   700
      End
      Begin VB.TextBox txtDeviceName 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   750
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   30
         Text            =   "frmCM17A.frx":59331
         ToolTipText     =   "Name of the device"
         Top             =   3880
         Width           =   700
      End
      Begin VB.TextBox txtDeviceName 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   750
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   29
         Text            =   "frmCM17A.frx":59337
         ToolTipText     =   "Name of the device"
         Top             =   3300
         Width           =   700
      End
      Begin VB.TextBox txtDeviceName 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   750
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   28
         Text            =   "frmCM17A.frx":5933D
         ToolTipText     =   "Name of the device"
         Top             =   2760
         Width           =   700
      End
      Begin VB.TextBox txtDeviceName 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   750
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   27
         Text            =   "frmCM17A.frx":59343
         ToolTipText     =   "Name of the device"
         Top             =   2160
         Width           =   700
      End
      Begin VB.TextBox txtDeviceName 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   750
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   26
         Text            =   "frmCM17A.frx":59349
         ToolTipText     =   "Name of the device"
         Top             =   1600
         Width           =   700
      End
      Begin VB.TextBox txtDeviceName 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   750
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   25
         Text            =   "frmCM17A.frx":5934F
         ToolTipText     =   "Name of the device"
         Top             =   1000
         Width           =   700
      End
      Begin VB.TextBox txtDeviceName 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   750
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   24
         Text            =   "frmCM17A.frx":59355
         ToolTipText     =   "Name of the device"
         Top             =   450
         Width           =   700
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   580
         Index           =   1
         Left            =   1500
         TabIndex        =   34
         Top             =   400
         Width           =   200
         _ExtentX        =   344
         _ExtentY        =   1032
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   580
         Index           =   2
         Left            =   1500
         TabIndex        =   35
         Top             =   960
         Width           =   200
         _ExtentX        =   344
         _ExtentY        =   1032
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   580
         Index           =   3
         Left            =   1500
         TabIndex        =   36
         Top             =   1560
         Width           =   200
         _ExtentX        =   344
         _ExtentY        =   1032
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   580
         Index           =   4
         Left            =   1500
         TabIndex        =   37
         Top             =   2160
         Width           =   200
         _ExtentX        =   344
         _ExtentY        =   1032
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   580
         Index           =   5
         Left            =   1500
         TabIndex        =   38
         Top             =   2760
         Width           =   200
         _ExtentX        =   344
         _ExtentY        =   1032
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   580
         Index           =   6
         Left            =   1500
         TabIndex        =   39
         Top             =   3240
         Width           =   200
         _ExtentX        =   344
         _ExtentY        =   1032
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin MSComctlLib.ProgressBar pbDim 
         Height          =   580
         Index           =   7
         Left            =   1500
         TabIndex        =   40
         Top             =   3840
         Width           =   200
         _ExtentX        =   344
         _ExtentY        =   1032
         _Version        =   393216
         Appearance      =   1
         Orientation     =   1
      End
      Begin VB.Label lblDeviceNum 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "(Active)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   480
         TabIndex        =   33
         ToolTipText     =   "Device Code for the device"
         Top             =   180
         Width           =   915
      End
      Begin VB.Image imgLED 
         Height          =   240
         Index           =   8
         Left            =   3520
         Picture         =   "frmCM17A.frx":59369
         Top             =   4560
         Width           =   240
      End
      Begin VB.Image imgLED 
         Height          =   240
         Index           =   7
         Left            =   3520
         Picture         =   "frmCM17A.frx":596AB
         Top             =   3986
         Width           =   240
      End
      Begin VB.Image imgLED 
         Height          =   240
         Index           =   6
         Left            =   3520
         Picture         =   "frmCM17A.frx":599ED
         Top             =   3415
         Width           =   240
      End
      Begin VB.Image imgLED 
         Height          =   240
         Index           =   5
         Left            =   3520
         Picture         =   "frmCM17A.frx":59D2F
         Top             =   2844
         Width           =   240
      End
      Begin VB.Image imgLED 
         Height          =   240
         Index           =   4
         Left            =   3520
         Picture         =   "frmCM17A.frx":5A071
         Top             =   2273
         Width           =   240
      End
      Begin VB.Image imgLED 
         Height          =   240
         Index           =   3
         Left            =   3520
         Picture         =   "frmCM17A.frx":5A3B3
         Top             =   1702
         Width           =   240
      End
      Begin VB.Image imgLED 
         Height          =   240
         Index           =   2
         Left            =   3520
         Picture         =   "frmCM17A.frx":5A6F5
         Top             =   1131
         Width           =   240
      End
      Begin VB.Image imgLED 
         Height          =   240
         Index           =   1
         Left            =   3520
         Picture         =   "frmCM17A.frx":5AA37
         Top             =   560
         Width           =   240
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         FillColor       =   &H0000EEEE&
         Height          =   255
         Left            =   2520
         Shape           =   2  'Oval
         Top             =   5640
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image imgAllOff 
         Height          =   375
         Left            =   2910
         Picture         =   "frmCM17A.frx":5AD79
         ToolTipText     =   "Click to decrease brightness of device"
         Top             =   5280
         Width           =   600
      End
      Begin VB.Image imgAllOn 
         Height          =   375
         Left            =   2080
         Picture         =   "frmCM17A.frx":5B973
         ToolTipText     =   "Click to increase brightness of device"
         Top             =   5280
         Width           =   600
      End
      Begin VB.Image CloseButton 
         Height          =   315
         Left            =   3670
         Picture         =   "frmCM17A.frx":5C56D
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   315
      End
      Begin VB.Label lblDeviceNum 
         BackColor       =   &H80000009&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   465
         TabIndex        =   23
         ToolTipText     =   "Device Code for the device"
         Top             =   4440
         Width           =   315
      End
      Begin VB.Label lblDeviceNum 
         BackColor       =   &H80000009&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   465
         TabIndex        =   22
         ToolTipText     =   "Device Code for the device"
         Top             =   3885
         Width           =   315
      End
      Begin VB.Label lblDeviceNum 
         BackColor       =   &H80000009&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   465
         TabIndex        =   21
         ToolTipText     =   "Device Code for the device"
         Top             =   3300
         Width           =   315
      End
      Begin VB.Label lblDeviceNum 
         BackColor       =   &H80000009&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   465
         TabIndex        =   20
         ToolTipText     =   "Device Code for the device"
         Top             =   2760
         Width           =   315
      End
      Begin VB.Label lblDeviceNum 
         BackColor       =   &H80000009&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   465
         TabIndex        =   19
         ToolTipText     =   "Device Code for the device"
         Top             =   2160
         Width           =   315
      End
      Begin VB.Label lblDeviceNum 
         BackColor       =   &H80000009&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   465
         TabIndex        =   18
         ToolTipText     =   "Device Code for the device"
         Top             =   1605
         Width           =   315
      End
      Begin VB.Label lblDeviceNum 
         BackColor       =   &H80000009&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   450
         TabIndex        =   17
         ToolTipText     =   "Device Code for the device"
         Top             =   1005
         Width           =   315
      End
      Begin VB.Label lblDeviceNum 
         BackColor       =   &H80000009&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   465
         TabIndex        =   12
         ToolTipText     =   "Device Code for the device"
         Top             =   450
         Width           =   315
      End
      Begin VB.Image imgSel 
         Height          =   270
         Left            =   2500
         Picture         =   "frmCM17A.frx":5CAEF
         ToolTipText     =   "Click to select devices 1-8 or 9-16"
         Top             =   5640
         Width           =   660
      End
      Begin VB.Image imgDimUp 
         Height          =   480
         Left            =   2040
         Picture         =   "frmCM17A.frx":5D479
         ToolTipText     =   "Click to increase brightness of device"
         Top             =   4850
         Width           =   660
      End
      Begin VB.Image imgDimDown 
         Height          =   480
         Left            =   2860
         Picture         =   "frmCM17A.frx":5E53B
         ToolTipText     =   "Click to decrease brightness of device"
         Top             =   4845
         Width           =   660
      End
      Begin VB.Image imgLEDOn 
         Height          =   240
         Left            =   2715
         Picture         =   "frmCM17A.frx":5F5FD
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgOff 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   8
         Left            =   2880
         Picture         =   "frmCM17A.frx":5F93F
         ToolTipText     =   "Click to turn Device Off"
         Top             =   4500
         Width           =   600
      End
      Begin VB.Image imgOff 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   7
         Left            =   2880
         Picture         =   "frmCM17A.frx":60539
         ToolTipText     =   "Click to turn Device Off"
         Top             =   3930
         Width           =   600
      End
      Begin VB.Image imgOff 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   6
         Left            =   2880
         Picture         =   "frmCM17A.frx":61133
         ToolTipText     =   "Click to turn Device Off"
         Top             =   3360
         Width           =   600
      End
      Begin VB.Image imgOff 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   5
         Left            =   2880
         Picture         =   "frmCM17A.frx":61D2D
         ToolTipText     =   "Click to turn Device Off"
         Top             =   2790
         Width           =   600
      End
      Begin VB.Image imgOff 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   4
         Left            =   2880
         Picture         =   "frmCM17A.frx":62927
         ToolTipText     =   "Click to turn Device Off"
         Top             =   2220
         Width           =   600
      End
      Begin VB.Image imgOff 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   3
         Left            =   2880
         Picture         =   "frmCM17A.frx":63521
         ToolTipText     =   "Click to turn Device Off"
         Top             =   1650
         Width           =   600
      End
      Begin VB.Image imgOff 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   2
         Left            =   2880
         Picture         =   "frmCM17A.frx":6411B
         ToolTipText     =   "Click to turn Device Off"
         Top             =   1080
         Width           =   600
      End
      Begin VB.Image imgOff 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   1
         Left            =   2880
         Picture         =   "frmCM17A.frx":64D15
         ToolTipText     =   "Click to turn Device Off"
         Top             =   500
         Width           =   600
      End
      Begin VB.Image imgOn 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   7
         Left            =   2000
         Picture         =   "frmCM17A.frx":6590F
         ToolTipText     =   "Click to turn Device On"
         Top             =   3930
         Width           =   600
      End
      Begin VB.Image imgOn 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   5
         Left            =   1950
         Picture         =   "frmCM17A.frx":66509
         ToolTipText     =   "Click to turn Device On"
         Top             =   2790
         Width           =   600
      End
      Begin VB.Image imgOn 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   4
         Left            =   1970
         Picture         =   "frmCM17A.frx":67103
         ToolTipText     =   "Click to turn Device On"
         Top             =   2220
         Width           =   600
      End
      Begin VB.Image imgOn 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   3
         Left            =   2000
         Picture         =   "frmCM17A.frx":67CFD
         ToolTipText     =   "Click to turn Device On"
         Top             =   1650
         Width           =   600
      End
      Begin VB.Image imgOn 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   2
         Left            =   2070
         Picture         =   "frmCM17A.frx":688F7
         ToolTipText     =   "Click to turn Device On"
         Top             =   1080
         Width           =   600
      End
      Begin VB.Image imgOn 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   1
         Left            =   2160
         Picture         =   "frmCM17A.frx":694F1
         ToolTipText     =   "Click to turn Device On"
         Top             =   495
         Width           =   600
      End
      Begin VB.Image imgOn 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   6
         Left            =   1970
         Picture         =   "frmCM17A.frx":6A0EB
         ToolTipText     =   "Click to turn Device On"
         Top             =   3360
         Width           =   600
      End
      Begin VB.Image imgOn 
         Appearance      =   0  'Flat
         Height          =   375
         Index           =   8
         Left            =   2040
         Picture         =   "frmCM17A.frx":6ACE5
         ToolTipText     =   "Click to turn Device On"
         Top             =   4500
         Width           =   600
      End
   End
   Begin VB.Image imgAllOnHover 
      Height          =   375
      Left            =   5640
      Picture         =   "frmCM17A.frx":6B8DF
      Top             =   5760
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgAllOffHover 
      Height          =   375
      Left            =   5640
      Picture         =   "frmCM17A.frx":6C4D9
      Top             =   6120
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgAllOffDown 
      Height          =   375
      Left            =   5040
      Picture         =   "frmCM17A.frx":6D0D3
      Top             =   6120
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgAllOffUp 
      Height          =   375
      Left            =   4440
      Picture         =   "frmCM17A.frx":6DCCD
      Top             =   6120
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgAllOnDown 
      Height          =   375
      Left            =   5040
      Picture         =   "frmCM17A.frx":6E8C7
      Top             =   5760
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgAllOnUp 
      Height          =   375
      Left            =   4440
      Picture         =   "frmCM17A.frx":6F4C1
      Top             =   5760
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgSel9 
      Height          =   270
      Left            =   5040
      Picture         =   "frmCM17A.frx":700BB
      Top             =   5400
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgSel1 
      Height          =   270
      Left            =   4320
      Picture         =   "frmCM17A.frx":70A45
      Top             =   5400
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgDimDownHover 
      Height          =   480
      Left            =   5040
      Picture         =   "frmCM17A.frx":713CF
      Top             =   4920
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgDimUpHover 
      Height          =   480
      Left            =   4440
      Picture         =   "frmCM17A.frx":72491
      Top             =   4920
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgDimUpOff 
      Height          =   480
      Left            =   4440
      Picture         =   "frmCM17A.frx":73553
      Top             =   3960
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgDimDownOff 
      Height          =   480
      Left            =   5040
      Picture         =   "frmCM17A.frx":74615
      Top             =   3960
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgDimDownOn 
      Height          =   465
      Left            =   5040
      Picture         =   "frmCM17A.frx":756D7
      Top             =   4440
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgDimUpOn 
      Height          =   465
      Left            =   4440
      Picture         =   "frmCM17A.frx":76715
      Top             =   4440
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgHover 
      Height          =   375
      Left            =   5640
      Picture         =   "frmCM17A.frx":77753
      Top             =   3480
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgLEDOff 
      Height          =   240
      Left            =   4080
      Picture         =   "frmCM17A.frx":7834D
      Top             =   3480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgDown 
      Height          =   375
      Left            =   5040
      Picture         =   "frmCM17A.frx":7868F
      Top             =   3480
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image imgUp 
      Height          =   375
      Left            =   4440
      Picture         =   "frmCM17A.frx":79289
      Top             =   3480
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Com Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Dim/ Bright %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "House Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Device Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mSetHouse 
         Caption         =   "Set &House Code"
         Begin VB.Menu mHouse 
            Caption         =   "A"
            Index           =   0
         End
      End
      Begin VB.Menu mBar 
         Caption         =   "-"
      End
      Begin VB.Menu mShow 
         Caption         =   "Show Extended Controls"
      End
      Begin VB.Menu mMin 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "&TrayMenu"
      Begin VB.Menu mnuTrayRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmCM17A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'X10 Firecracker (CM17A) PalmPad Controller
'submitted by Tom Pydeski
'X10 put out some nifty home control products that communicate
'over the power lines and respond to commands to turn on; turn off; dim; etc.
'The CM17A is an rf controlled device that receives rf commands from a remote or
'the "firecracker" rf pc interface and relays them to the device modules.
'The HouseCode can be any 1 of 16 (A through P) and each house code can handle
'any of 16 devices (1 to 16).
'Keware (http://www.homeseer.com/downloads/index.htm)
'had put out an open source usercontrol for the firecracker, but I changed it to make
'it a class module, in order to eliminate packaging an ocx with the project.
'I took X10's firecracker interface picture of the palmpad and added my own graphic
'buttons to simulate the button presses and hovering.  I also added the ability to
'label the buttons with the device name.  Additionally I found some neat code to
'rotate an image and implemented that in allowing the selection of a housecode via
'the palmpad's rotary switch.  I also added LED's for each device to indicate
'their on/off status and used a vertical progress bar to set the dim level.
'I implemented a device status for each device.  Finally, I was able to figure out
'the all on and all off commands (which the firecracker does not support) and
'implemented them as well.
'Of course all of this is useless if you don't have the x10 hardware that it
'interfaces with and controls (http://www.x10.com/automation/ck18a_s_ps32.html)
'(They were practically giving away the firecracker starter kit a few years back.)
'
Option Explicit
'for dim up and dim down routines
Dim Released As Byte
'for moving form
Dim StartX As Long
Dim StartY As Long
Dim IgnoreHouse As Byte
Dim h As Integer
Dim i As Integer
Dim f As Integer
Dim n As Integer
'for rotating the house selection
Dim picHouseCenterX As Single
Dim picHouseCenterY As Single
Dim picX As Single
Dim picY As Single
Dim NewHouse As Integer
Dim NewVal As Integer
'our X10 class
Dim WithEvents X10CM17A As clsX10CM17A
Attribute X10CM17A.VB_VarHelpID = -1
'hand cursor
Dim lHandle As Long
Const HandCursor = 32649&
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Private Sub CloseButton_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc("Q") Or KeyAscii = Asc("q") Then
    Unload Me
    End
End If
End Sub

Private Sub Form_Load()
'get the stored locations from the ini file
With Me
    .Left = GetSettingIni(App.EXEName, "Settings", "Left", 0)
    .Top = GetSettingIni(App.EXEName, "Settings", "Top", 0)
    .Height = 6630
    .Width = 4125
End With
'display the form differently if the controlbox is turned on
If Me.ControlBox = False Then
    CloseButton.Visible = True
    mFile.Visible = False
    mnuTray.Visible = False
Else
    CloseButton.Visible = False
    mFile.Visible = True
    mnuTray.Visible = True
    Me.Caption = "Tom Py's X10 PalmPad"
    Me.Height = 7365
End If
'Populate the X10 Command Combo Box
cmbX10Com.Clear
cmbX10Com.List(0) = "All Units Off"
cmbX10Com.List(1) = "All Lights On"
cmbX10Com.List(2) = "On"
cmbX10Com.List(3) = "Off"
cmbX10Com.List(4) = "Dim"
cmbX10Com.List(5) = "Bright"
cmbX10Com.List(6) = "All Lights Off"
cmbX10Com.List(7) = "Extended"
cmbX10Com.List(8) = "Hail Request"
cmbX10Com.List(9) = "Hail Ack"
cmbX10Com.List(10) = "Pre-set dim1"
cmbX10Com.List(11) = "Pre-set dim2"
cmbX10Com.List(12) = "Extended Data"
cmbX10Com.List(13) = "Status On"
cmbX10Com.List(14) = "Status Off"
cmbX10Com.List(15) = "Status Request"
cmbX10Com.ListIndex = 0
'Initialize our X10 class
Set X10CM17A = New clsX10CM17A
'houscode 0-15 and devicecode 1-16
'get our last house and device codes from the ini file
X10CM17A.HouseCode = Val(GetSettingIni(App.EXEName, "Settings", "HouseCode"))
X10CM17A.DeviceCode = Val(GetSettingIni(App.EXEName, "Settings", "DeviceCode"))
'selector switch setting (1-8 or 9-16) is also stored
SelSwitch = Val(GetSettingIni(App.EXEName, "Settings", "SelSwitch"))
'retrieve the names of our devices from the ini file
For h = 0 To 15
    X10(h).Configured = False
    For i = 1 To 16
        X10(h).DeviceName(i) = GetSettingIni(App.EXEName, "HouseCode" & h, "DeviceName" & i)
        If Len(Trim$(X10(h).DeviceName(i))) > 0 Then
            'if we have a name for any device, then set the flag that tells
            'us this house has a configuration
            X10(h).Configured = True
        Else
            'If there is no name saved for a device, then
            'give it a default name with housecode and devicecode...i.e. A1, A2, etc.
            X10(h).DeviceName(i) = Chr$(Asc("A") + h) & i
        End If
    Next i
Next h
DoEvents
'populate the possible device codes into the combo box
cmbDeviceCode.Clear
For i = 0 To 16
    cmbDeviceCode.AddItem i
Next i
cmbDeviceCode.ListIndex = X10CM17A.DeviceCode
'populate the possible house codes into the combo box and menu
cmbHouseCode.AddItem mHouse(0).Caption
For i = 1 To 15
    Load mHouse(i)
    mHouse(i).Caption = Chr$(Asc("A") + i)
    cmbHouseCode.AddItem mHouse(i).Caption
Next i
'check the menu item for the selected house
mHouse(X10CM17A.HouseCode).Checked = True
'below allows us to not process the events that fire when we change a house code
IgnoreHouse = 1
cmbHouseCode.ListIndex = X10CM17A.HouseCode
'select the proper picture for the selector switch
If SelSwitch = 0 Then
    imgSel.Picture = imgSel1.Picture
    imgSel.ToolTipText = "Click to select devices 9-16"
Else
    imgSel.Picture = imgSel9.Picture
    imgSel.ToolTipText = "Click to select devices 1-8"
End If
'load the devicenames and numbers
For i = 1 To 8
    'if the selector switch is 1-8, then display these numbers
    'if we are 9-16, then the buttons are labeled with those numbers instead
    lblDeviceNum(i).Caption = (8 * SelSwitch) + i
    txtDeviceName(i).Text = X10(X10CM17A.HouseCode).DeviceName(i)
Next i
'below is the standard way to put an icon in the tray.
With Nid
    .cbSize = Len(Nid)
    .hwnd = Me.hwnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = "Tom Py's X10 PalmPad" & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, Nid
'we can call the program with a switch indicating a device to toggle at load
If Command$ = "" Then Me.Show
DoEvents
'now we setup everything for our housecode
SetupHouse X10CM17A.HouseCode
'initialize the communications to the firecracker cm17a
Init_Click
DoEvents
Sleep (500)
DoEvents
'get the last status of each device from the data file
f = FreeFile
Open "X10Status.dat" For Random As #f Len = Len(X10Out(0))
For h = 0 To 15
    Get #f, h + 1, X10Out(h)
Next h
Close #f
'display the status of all devices for the palmpad leds
UpdateDevice
'
If Command$ <> "" Then
    'if we pass a device number, then just send the command and exit
    picPalm_KeyPress Asc(Command$)
    DoEvents
    mExit_Click
Else
    Me.picPalm.SetFocus
End If
'load our cute little hand cursor
'this was from the sample by LaVolpe (thanks!)
'at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=63065&lngWId=1
lHandle = LoadCursor(0, HandCursor)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    'the right click displays the menu
    PopupMenu mFile
End If
End Sub

Private Sub Form_Resize()
If WindowState = vbMinimized Then
    Me.Hide
    Me.Refresh
Else
    'Shell_NotifyIcon NIM_DELETE, Nid
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
'this is for when we are in the tray
i = X / Screen.TwipsPerPixelX
Select Case i
    Case WM_RBUTTONDOWN:
        Me.PopupMenu mnuTray
End Select
If i = WM_LBUTTONDOWN Then 'WM_LBUTTONDBLCLK
    Me.WindowState = vbNormal
    Me.Show
    Me.ZOrder 0
    Me.Refresh
    'Shell_NotifyIcon NIM_DELETE, Nid
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.MousePointer = 11
'save all of our settings to the ini file
SaveSettingIni App.EXEName, "Settings", "Left", Me.Left
SaveSettingIni App.EXEName, "Settings", "Top", Me.Top
SaveSettingIni App.EXEName, "Settings", "HouseCode", X10CM17A.HouseCode
SaveSettingIni App.EXEName, "Settings", "DeviceCode", X10CM17A.DeviceCode
SaveSettingIni App.EXEName, "Settings", "SelSwitch", SelSwitch
'save the devicenames for all of our housecodes
ChDir App.Path
f = FreeFile
'store our labels into a file
Open "X10Labels.ini" For Random As #f Len = Len(X10(0))
For h = 0 To 15
    Put #f, h + 1, X10(h)
Next h
Close #f
'
'store the status of each device
f = FreeFile
Open "X10Status.dat" For Random As #f Len = Len(X10Out(0))
For h = 0 To 15
    Put #f, h + 1, X10Out(h)
Next h
Close #f
'
'save the device names for our selected housecode
'I had originally done all 16 house codes, but it takes too long
'so now we will only save the houses we have a configuration for
h = X10CM17A.HouseCode
For h = 0 To 15
    If X10(h).Configured = True Then
        For i = 1 To 16
            SaveSettingIni App.EXEName, "HouseCode" & h, "DeviceName" & i, X10(h).DeviceName(i)
        Next i
    End If
Next h
'reset the communications to the cm17A
X10CM17A.ResetCom
Me.MousePointer = 0
'take the icon out of the system tray
Shell_NotifyIcon NIM_DELETE, Nid
End Sub

'----------------------------------------------------------------------------------
'below are the buttons for utilizing the extended settings and not the gui
Private Sub ButBRIGHT_Click()
X10Out(X10CM17A.HouseCode).Device(X10CM17A.DeviceCode) = 1
X10CM17A.Exec cmbHouseCode.Text, cmbDeviceCode.Text, C_BRIGHT, Val(txtDim)
pbDim(X10CM17A.DeviceCode).Value = Val(txtDim)
End Sub

Private Sub ButDIM_Click()
X10Out(X10CM17A.HouseCode).Device(X10CM17A.DeviceCode) = 1
X10CM17A.Exec cmbHouseCode.Text, cmbDeviceCode.Text, C_DIM, Val(txtDim)
pbDim(X10CM17A.DeviceCode).Value = Val(txtDim)
End Sub

Private Sub ButOFF_Click()
X10Out(X10CM17A.HouseCode).Device(X10CM17A.DeviceCode) = 0
X10CM17A.Exec cmbHouseCode.Text, cmbDeviceCode.Text, C_OFF
End Sub

Private Sub ButON_Click()
X10Out(X10CM17A.HouseCode).Device(X10CM17A.DeviceCode) = 1
X10CM17A.Exec cmbHouseCode.Text, cmbDeviceCode.Text, C_ON
End Sub

Private Sub cmdAllOn_Click()
X10CM17A.Exec cmbHouseCode.Text, 1, ALL_LIGHTS_ON
For i = 1 To 16
    X10Out(X10CM17A.HouseCode).Device(i) = 1
Next i
UpdateDevice
End Sub

Private Sub cmdAllOff_Click()
X10CM17A.Exec cmbHouseCode.Text, 1, ALL_LIGHTS_OFF
For i = 1 To 16
    X10Out(X10CM17A.HouseCode).Device(i) = 0
Next i
UpdateDevice
End Sub
'----------------------------------------------------------------------------------

Private Sub imgDimUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'start the loop for increasing the brightness until we get a mouseup event
DimUp
End Sub

Private Sub imgDimUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'that LaVolpe....he gave us this
SetCursor lHandle
'
If Button = 0 Then
    'change all the buttons to their normal up pictures
    HoverOff
    'set the selected button to show its hover picture
    imgDimUp.Picture = imgDimUpHover.Picture
End If
End Sub

Private Sub imgDimUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'the mouseup event occurred, so lets turn off the led pic and set the released flag
'so we can exit our dim loop
imgDimUp.Picture = imgDimUpOff.Picture
imgLEDOn.Visible = False
Released = 1
End Sub

Private Sub imgDimDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'the dimdown routine is handled the same way as the dimup described above
DimDown
End Sub

Private Sub imgDimDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor lHandle
'
If Button = 0 Then
    HoverOff
    imgDimDown.Picture = imgDimDownHover.Picture
End If
End Sub

Private Sub imgDimDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDimDown.Picture = imgDimDownOff.Picture
imgLEDOn.Visible = False
Released = 1
End Sub

Sub DimUp()
'set the button to show us the "on" picture
imgDimUp.Picture = imgDimUpOn.Picture
'light up our fake LED
imgLEDOn.Visible = True
DoEvents
Debug.Print "DimUp Pressed"
'start the loop for increasing the brightness until we get a mouseup event
Do
    imgLEDOn.Visible = True
    DoEvents
    'wait for the mouse up or keyup event
    If Released = 1 Then Exit Do
    'send the bright command to the cm17a
    X10CM17A.Exec cmbHouseCode.Text, Str$(X10CM17A.DeviceCode), C_BRIGHT
    'wait for 100 msec.
    Sleep 100 'X10CM17A.WaitMicroSecs (2500)
    'turn our fake LED off
    imgLEDOn.Visible = False
    DoEvents
    Sleep 100
Loop
'clear the released flag
Released = 0
Debug.Print "DimUp Released"
End Sub

Sub DimDown()
'same as above, only this dims the light down instead of brightening it
imgDimDown.Picture = imgDimDownOn.Picture
imgLEDOn.Visible = True
DoEvents
Debug.Print "DimDown Pressed"
Do
    imgLEDOn.Visible = True
    DoEvents
    If Released = 1 Then Exit Do
    X10CM17A.Exec cmbHouseCode.Text, Str$(X10CM17A.DeviceCode), C_DIM
    Sleep 100 'X10CM17A.WaitMicroSecs (2500)
    imgLEDOn.Visible = False
    DoEvents
    Sleep 100
Loop
Released = 0
Debug.Print "DimDown Released"
End Sub

Private Sub imgOn_Click(Index As Integer)
'turn our device on
X10CM17A.DeviceCode = (8 * SelSwitch) + Index
X10CM17A.Exec cmbHouseCode, CStr(X10CM17A.DeviceCode), C_ON
X10Out(X10CM17A.HouseCode).Device(X10CM17A.DeviceCode) = 1
UpdateDevice
LastDim(Index) = 100
End Sub

Private Sub imgOn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'show the down picture and light our led
imgOn(Index).Picture = imgDown.Picture
imgLEDOn.Visible = True
End Sub

Private Sub imgOn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor lHandle
'
If Button = 0 Then
    HoverOff
    'show our hover picture
    imgOn(Index).Picture = imgHover.Picture
End If
End Sub

Private Sub imgOn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOn(Index).Picture = imgUp.Picture
imgLEDOn.Visible = False
End Sub

Private Sub imgOff_Click(Index As Integer)
X10CM17A.DeviceCode = (8 * SelSwitch) + Index
X10CM17A.Exec cmbHouseCode, CStr(X10CM17A.DeviceCode), C_OFF
X10Out(X10CM17A.HouseCode).Device(X10CM17A.DeviceCode) = 0
UpdateDevice
LastDim(Index) = 0
End Sub

Private Sub imgOff_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOff(Index).Picture = imgDown.Picture
imgLEDOn.Visible = True
End Sub

Private Sub imgOff_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor lHandle
'
If Button = 0 Then
    HoverOff
    imgOff(Index).Picture = imgHover.Picture
End If
End Sub

Private Sub imgOff_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOff(Index).Picture = imgUp.Picture
imgLEDOn.Visible = False
End Sub

Private Sub imgAllOn_Click()
'turn on all of the lights
cmdAllOn_Click
End Sub

Private Sub imgAllOn_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgAllOn.Picture = imgAllOnDown.Picture
imgLEDOn.Visible = True
End Sub

Private Sub imgAllOn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor lHandle
'
If Button = 0 Then
    HoverOff
    imgAllOn.Picture = imgAllOnHover.Picture
End If
End Sub

Private Sub imgAllOn_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgAllOn.Picture = imgAllOnUp.Picture
imgLEDOn.Visible = False
End Sub

Private Sub imgAllOff_Click()
'cmdAllOff_Click
'X10CM17A.Exec cmbHouseCode, CStr(X10CM17A.DeviceCode), ALL_LIGHTS_OFF
Dim HC As Integer
Dim OldHC As Integer
'store our current housecode because this routine will change it
OldHC = X10CM17A.HouseCode
'send all off to all housecodes that are configured
For HC = 0 To 15
    If X10(HC).Configured = True Then
        Debug.Print "Sending All Off to HouseCode "; Chr$(Asc("A") + HC)
        imgLEDOn.Visible = True
        DoEvents
        X10CM17A.Exec Chr$(Asc("A") + HC), 1, ALL_LIGHTS_OFF
        For i = 1 To 16
            X10Out(HC).Device(i) = 0
        Next i
        Sleep 10
        imgLEDOn.Visible = False
        DoEvents
        Sleep 10
    End If
Next HC
'retrieve our stored house code
X10CM17A.HouseCode = OldHC
'update all the devices
UpdateDevice
End Sub

Private Sub imgAllOff_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgAllOff.Picture = imgAllOffDown.Picture
imgLEDOn.Visible = True
End Sub

Private Sub imgAllOff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor lHandle
'
If Button = 0 Then
    HoverOff
    imgAllOff.Picture = imgAllOffHover.Picture
End If
End Sub

Private Sub imgAllOff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgAllOff.Picture = imgAllOffUp.Picture
imgLEDOn.Visible = False
End Sub

Private Sub imgSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor lHandle
'
If Button = 0 Then
    'imgSel.Picture = imgSelHover.Picture
    Shape1.Visible = True
End If
End Sub

Private Sub imgSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this selects devices 1-8 or 9-16
SelSwitch = 1 - SelSwitch
If SelSwitch = 0 Then
    imgSel.Picture = imgSel1.Picture
    imgSel.ToolTipText = "Click to select devices 9-16"
Else
    imgSel.Picture = imgSel9.Picture
    imgSel.ToolTipText = "Click to select devices 1-8"
End If
For i = 1 To 8
    lblDeviceNum(i).Caption = (8 * SelSwitch) + i
    txtDeviceName(i).Text = X10(X10CM17A.HouseCode).DeviceName((8 * SelSwitch) + i)
Next i
UpdateDevice
Refresh
DoEvents
End Sub

Sub HoverOff()
'this displays all of the up pictures in our buttons
For i = 1 To 8
    'this is something neat i tried so that the control we are over is not
    'updated with the up picture, but rather only the other buttons
    If imgOn(i) <> Me.ActiveControl Then imgOn(i).Picture = imgUp.Picture
    If imgOff(i) <> Me.ActiveControl Then imgOff(i).Picture = imgUp.Picture
Next i
If imgDimUp <> Me.ActiveControl Then imgDimUp.Picture = imgDimUpOff.Picture
If imgDimDown <> Me.ActiveControl Then imgDimDown.Picture = imgDimDownOff.Picture
If SelSwitch = 0 Then
    imgSel.Picture = imgSel1.Picture
Else
    imgSel.Picture = imgSel9.Picture
End If
'we have a little elipse around the selector to show hovering
Shape1.Visible = False
If imgAllOn <> Me.ActiveControl Then imgAllOn.Picture = imgAllOnUp.Picture
If imgAllOff <> Me.ActiveControl Then imgAllOff.Picture = imgAllOffUp.Picture
End Sub

Private Sub Init_Click()
Dim init_error
'initialize the communications to the CM17A
X10CM17A.ComPort = Val(TxtCom)
init_error = X10CM17A.Init
If init_error <> 0 Then
    MsgBox "Error initializing CM17A", vbExclamation + vbOKOnly
End If
Init.Caption = "Initialized"
Init.Enabled = False
End Sub

Private Sub cmbHouseCode_Change()
If IgnoreHouse = 1 Then Exit Sub
'setup the program for the new house
SetupHouse cmbHouseCode.ListIndex
End Sub

Private Sub cmbHouseCode_Click()
If IgnoreHouse = 1 Then Exit Sub
SetupHouse cmbHouseCode.ListIndex
End Sub

Private Sub cmbHouseCode_Scroll()
If IgnoreHouse = 1 Then Exit Sub
SetupHouse cmbHouseCode.ListIndex
End Sub

Private Sub mExit_Click()
Unload Me
End Sub

Private Sub mHouse_Click(Index As Integer)
'setup the program for the new house when the menu is changed
SetupHouse Index
End Sub

Sub SetMenu()
End Sub

Private Sub mMin_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub mnuTrayExit_Click()
mExit_Click
End Sub

Private Sub mShow_Click()
'hide or display the extended controls
If mShow.Checked = False Then
     mShow.Checked = True
     Do
        Me.Width = Me.Width + 1
        DoEvents
    Loop Until Me.Width >= 6630
Else
     mShow.Checked = False
     Do
        Me.Width = Me.Width - 1
        DoEvents
    Loop Until Me.Width <= 4150
End If
End Sub

Private Sub mnuTrayRestore_Click()
'restore the program from the tray
Me.WindowState = vbNormal
Me.Show
Me.Refresh
'i think i want to keep the icon in the tray...
'Shell_NotifyIcon NIM_DELETE, Nid
End Sub

Private Sub pbDim_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'pbdim is a vertical progress bar that i am using to set the dim 0-100%
If Button = 1 Then
    With pbDim(Index)
            IgnScrl = 1
            .Value = NewVal
            ChangeDim (Index)
    End With
End If
End Sub

Private Sub pbDim_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
With pbDim(Index)
    NewVal = ((.Height - Y) / .Height) * 100
    'the top and bottom edges set the level to be 0 or 100
    If NewVal > 90 Then NewVal = 100
    If NewVal < 10 Then NewVal = 10
    If NewVal > 0 And NewVal < 100 Then
        If Button = 1 Then
            IgnScrl = 1
            .Value = NewVal
            Refresh
        End If
        .ToolTipText = "Level = " & .Value & " => Set Level to " & NewVal '.Value
    End If
End With
End Sub

Sub ChangeDim(Index As Integer)
'this routine will set the brightness of the selected device
Dim DesiredLevel As Integer
Debug.Print pbDim(Index).Value, LastDim(Index)
'if it is our first dim, start from full volume
If LastDim(Index) = 0 Then LastDim(Index) = 100
If pbDim(Index).Value > LastDim(Index) Then
    cmbX10Com.ListIndex = 5 'bright
Else
    cmbX10Com.ListIndex = 4 'dim
End If
cmbDeviceCode.ListIndex = Index
DesiredLevel = pbDim(Index).Value
txtDim.Text = DesiredLevel
LastDim(Index) = DesiredLevel
If X10Init = 0 Then Exit Sub
'
'toggle our status bit for the selected device
X10Out(X10CM17A.HouseCode).Device(Index) = 1
imgLED(Index).Picture = imgLEDOn.Picture
'i think that the brightness parameter changes the calling parameter's value, so this
'is why I pass another variable (DesiredLevel) insted of the pbdim value
X10CM17A.Exec cmbHouseCode.Text, cmbDeviceCode.Text, cmbX10Com.ListIndex, DesiredLevel
End Sub

Private Sub picHouseCode_DblClick()
'this nifty little rotation thingy was modified from a submission by E.Spencer
'
'get the center point of our house code picture
picHouseCenterX = (picHouseCode.Width / Screen.TwipsPerPixelX) / 2
picHouseCenterY = (picHouseCode.Height / Screen.TwipsPerPixelY) / 2
'now display the code with the picture
InitPic picPointer
Theta = (6.28 / 360) * (360 - GetHouseFromPoints)
picPointer.Visible = True
DoEvents
'
'q = 90
RotatePic Theta, picPointer, picHouseCode
'
ClosePic
SetupHouse NewHouse
picPointer.Visible = False
End Sub

Private Sub picHouseCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'get the center point of our house code picture
picHouseCenterX = (picHouseCode.Width / Screen.TwipsPerPixelX) / 2
picHouseCenterY = (picHouseCode.Height / Screen.TwipsPerPixelY) / 2
'now display the code with the picture
InitPic picPointer
End Sub

Private Sub picHouseCode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetCursor lHandle
'
If Button <> 1 Then Exit Sub
'get the offset of the mouse position from center
picX = X - picHouseCenterX
picY = picHouseCenterY - Y
Theta = (6.28 / 360) * (360 - GetHouseFromPoints)
picPointer.Visible = True
DoEvents
'
RotatePic Theta, picPointer, picHouseCode
End Sub

Private Sub picHouseCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClosePic
SetupHouse NewHouse
picPointer.Visible = False
End Sub

Function GetHouseFromPoints() As Integer
'This routine was my way of dealing with rotating by using the mouse
Dim q As Integer
Dim TanX As Double
Dim QuadrantOffset As Integer
'lets label the quadrants 1-4 starting with +x +y as 1 and working clockwise.
'We have to do this because the tangent is not represented 0-360, but rather -90 to +90
'we can change the angle to abs and then add an offset to put it in the proper quadrant.
'quadrant 1 is 0-90
'quadrant 2 is 90-180
'quadrant 3 is 180-270
'quadrant 4 is 270-360 (0)
'
'get the offset of the mouse position from center
If picX > 0 And picY > 0 Then
    QuadrantOffset = 0
    TanX = picX / picY
ElseIf picX > 0 And picY < 0 Then
    QuadrantOffset = 90
    TanX = Abs(picY / picX)
ElseIf picX < 0 And picY < 0 Then
    QuadrantOffset = 180
    TanX = (picX / picY)
ElseIf picX < 0 And picY > 0 Then
    QuadrantOffset = 270
    TanX = (picY / picX)
End If
picX = Abs(picX)
picY = Abs(picY)
If picX = 0 Then
    'we are on the y axis
    'it is either 90 deg or 270 deg.
    If picY > 0 Then
        q = 0
    Else
        q = 180
    End If
ElseIf picY = 0 Then
    'we are on the x axis
    'it is either 90 deg or 270 deg.
    If picX > 0 Then
        q = 270
    Else
        q = 90
    End If
Else
    q = Abs(ArcTangent(TanX))
End If
'
'Debug.Print X; " "; Y; " "; picX; " "; picY; " "; q, QuadrantOffset, q + QuadrantOffset
q = q + QuadrantOffset
'now lets make it in increments of 1/16 to snap to the letter
NewHouse = q \ 22.5
If NewHouse = 16 Then NewHouse = 0
q = NewHouse * 22.5
GetHouseFromPoints = q
End Function

Private Sub picPalm_KeyDown(KeyCode As Integer, Shift As Integer)
'allow the arrow up and down keys to dim or brighten
If KeyCode = vbKeyDown Then
    DimDown
End If
If KeyCode = vbKeyUp Then
    DimUp
End If
End Sub

Sub picPalm_KeyPress(KeyAscii As Integer)
'q will quit
If KeyAscii = Asc("Q") Or KeyAscii = Asc("q") Then
    Unload Me
    End
End If
'the numbers
If KeyAscii = 48 Then Exit Sub
If KeyAscii >= 49 And KeyAscii <= 57 Then '0-9
    X10CM17A.DeviceCode = KeyAscii - 48
End If
Select Case UCase(Chr$(KeyAscii))
    Case "A"
        X10CM17A.DeviceCode = 10
    Case "B"
        X10CM17A.DeviceCode = 11
    Case "C"
        X10CM17A.DeviceCode = 12
    Case "D"
        X10CM17A.DeviceCode = 13
    Case "E"
        X10CM17A.DeviceCode = 14
    Case "F"
        X10CM17A.DeviceCode = 15
End Select
'toggle the selected device
X10Out(X10CM17A.HouseCode).Device(X10CM17A.DeviceCode) = 1 - X10Out(X10CM17A.HouseCode).Device(X10CM17A.DeviceCode)
If X10Out(X10CM17A.HouseCode).Device(X10CM17A.DeviceCode) = 1 Then
    'turn the device on
    XCommand = C_ON
Else
    'turn the device off
    XCommand = C_OFF
End If
UpdateDevice
X10CM17A.Exec cmbHouseCode.Text, Str$(X10CM17A.DeviceCode), XCommand
'delay so we don't fire again
Sleep (1000)
End Sub

Private Sub picPalm_KeyUp(KeyCode As Integer, Shift As Integer)
Released = 1
End Sub

Private Sub picPalm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mFile
ElseIf Button = 1 Then
    StartX = X
    StartY = Y
End If
End Sub

Private Sub picPalm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Me.Left = Me.Left + (X - StartX)
    Me.Top = Me.Top + (Y - StartY)
    SaveSetting App.EXEName, "Settings", "Top", Me.Top
    SaveSetting App.EXEName, "Settings", "Left", Me.Left
End If
HoverOff
End Sub

Private Sub scrDim_Change()
If IgnScrl = 1 Then Exit Sub
txtDim.Text = scrDim.Value
End Sub

Private Sub txtDeviceName_Change(Index As Integer)
If IgnoreHouse = 1 Then Exit Sub
SaveSettingIni App.EXEName, "HouseCode" & X10CM17A.HouseCode, "DeviceName" & Index, X10(X10CM17A.HouseCode).DeviceName(Index)
End Sub

Private Sub txtDeviceName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    'if we hit the enter key, send a backspace and set focus to the next text box
    KeyCode = 0
    'take out the carriage return from the text box
    SendKeys "{BACKSPACE}"
    DoEvents
    If Index < 8 Then
        txtDeviceName(Index + 1).SetFocus
    End If
End If
End Sub

Private Sub txtDeviceName_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'set the new text to the device name
X10(X10CM17A.HouseCode).DeviceName((8 * SelSwitch) + Index) = txtDeviceName(Index)
End Sub

Private Sub txtDim_Change()
If IgnScrl = 1 Then Exit Sub
scrDim.Value = Val(txtDim)
End Sub

Sub UpdateDevice()
Dim DevLabNum As Integer
For i = 1 To 8
    'set the backgrounds to the default white
    txtDeviceName(i).BackColor = &H80000005
    lblDeviceNum(i).BackColor = &H80000005
    'turn off the LED for the device
    imgLED(i).Picture = imgLEDOff.Picture
    If X10Out(X10CM17A.HouseCode).Device((8 * SelSwitch) + i) = 1 Then
        'if the device is on, display an on LED
        imgLED(i).Picture = imgLEDOn.Picture
        'also set the progress bar to max
        pbDim(i).Value = 100
        Refresh
    Else
        pbDim(i).Value = 0
    End If
Next i
'we only have 8 text boxes, so if we are devices 9-16, subtract 8
If X10CM17A.DeviceCode > 8 Then
    DevLabNum = X10CM17A.DeviceCode - 8
Else
    DevLabNum = X10CM17A.DeviceCode
End If
txtDeviceName(DevLabNum).BackColor = vbYellow
lblDeviceNum(DevLabNum).BackColor = vbYellow
cmbDeviceCode.ListIndex = X10CM17A.DeviceCode
Refresh
End Sub

Sub SetupHouse(NewHouseIn As Integer)
IgnoreHouse = 1
'set up our x10 class for the new housecode
X10CM17A.HouseCode = NewHouseIn
If X10CM17A.HouseCode < 0 Then
    X10CM17A.HouseCode = 0
End If
'reset the menu and check the appropriate menu item
For i = 0 To mHouse.UBound
    mHouse(i).Checked = False
Next i
mHouse(X10CM17A.HouseCode).Checked = True
'setup combo box
cmbHouseCode.ListIndex = X10CM17A.HouseCode
cmbHouseCode.Text = mHouse(X10CM17A.HouseCode).Caption
ShowHouse X10CM17A.HouseCode
'display new labels from the stored data for the selected house
For i = 1 To 8
    lblDeviceNum(i).Caption = (8 * SelSwitch) + i
    txtDeviceName(i).Text = Trim$(X10(X10CM17A.HouseCode).DeviceName((8 * SelSwitch) + i))
    txtDeviceName(i).ToolTipText = cmbHouseCode.Text & lblDeviceNum(i).Caption & ":" & txtDeviceName(i).Text
    imgOn(i).ToolTipText = cmbHouseCode.Text & lblDeviceNum(i).Caption & ":" & txtDeviceName(i).Text & " On"
    imgOff(i).ToolTipText = cmbHouseCode.Text & lblDeviceNum(i).Caption & ":" & txtDeviceName(i).Text & " Off"
Next i
IgnoreHouse = 0
End Sub

Sub ShowHouse(NewHouseIn As Integer)
'display the pointer to the selected house
picPointer.Visible = True
DoEvents
'now display the code with the picture
InitPic picPointer
n = 360 - ((NewHouseIn / 16) * 360)
'Debug.Print "HouseCode = "; NewHouseIn
'Convert degrees to theta value
Theta = (6.28 / 360) * n
RotatePic Theta, picPointer, picHouseCode
ClosePic
picPointer.Visible = False
End Sub


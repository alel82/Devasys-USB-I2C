VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7l.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmI2CCheck 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "I2C Check"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame15 
      BackColor       =   &H8000000E&
      Caption         =   "BKSV"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   4200
      TabIndex        =   86
      Top             =   5640
      Width           =   1335
      Begin VB.TextBox txtBKSV 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtBKSVslave 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   840
         MaxLength       =   2
         TabIndex        =   88
         Text            =   "74"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdReadBKSV 
         Caption         =   "Read"
         Height          =   375
         Left            =   120
         TabIndex        =   87
         ToolTipText     =   "Write To Set"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Slave"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   360
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   8160
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2400
      Left            =   5640
      TabIndex        =   64
      Top             =   5640
      Width           =   2775
      Begin VB.CheckBox chkDataEnable 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   79
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chkDataEnable 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   78
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chkDataEnable 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   77
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   76
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   75
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   74
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   73
         Top             =   480
         Width           =   375
      End
      Begin VB.CheckBox chkDataEnable 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   2400
         TabIndex        =   72
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox txtSubAdd2 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   71
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdRead1 
         Caption         =   "Read"
         Height          =   375
         Left            =   120
         TabIndex        =   68
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtSubAdd1 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   720
         MaxLength       =   2
         TabIndex        =   67
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtSlaveAdd 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         MaxLength       =   2
         TabIndex        =   66
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmdWrite1 
         Caption         =   "Write"
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblStatus2 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   82
         ToolTipText     =   "Function Status"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Sub 2"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   81
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   80
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblsub 
         Alignment       =   2  'Center
         Caption         =   "Sub 1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   70
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Slave"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write"
      Height          =   375
      Left            =   6240
      TabIndex        =   63
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset USB"
      Height          =   375
      Left            =   8520
      TabIndex        =   61
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H8000000E&
      Caption         =   "Special"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   120
      TabIndex        =   56
      Top             =   7080
      Width           =   2415
      Begin VB.CommandButton cmdWrite00 
         Caption         =   "Write 00"
         Height          =   375
         Left            =   120
         TabIndex        =   62
         ToolTipText     =   "Write To Set"
         Top             =   840
         Width           =   975
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "WP"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1080
         Left            =   1200
         TabIndex        =   58
         Top             =   130
         Width           =   1215
         Begin VB.CommandButton cmdWP1 
            Caption         =   "70-88-00"
            Height          =   350
            Left            =   120
            TabIndex        =   60
            ToolTipText     =   "Write To Set"
            Top             =   250
            Width           =   975
         End
         Begin VB.CommandButton cmdWP2 
            Caption         =   "70-8B-00"
            Height          =   350
            Left            =   120
            TabIndex        =   59
            ToolTipText     =   "Write To Set"
            Top             =   620
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdWriteFF 
         Caption         =   "Write FF"
         Height          =   375
         Left            =   120
         TabIndex        =   57
         ToolTipText     =   "Write To Set"
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H80000009&
      Caption         =   "PA0"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   6000
      TabIndex        =   53
      Top             =   6120
      Visible         =   0   'False
      Width           =   975
      Begin VB.OptionButton optPA0 
         Caption         =   "B"
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "SRQ High"
         Top             =   840
         Width           =   735
      End
      Begin VB.OptionButton optPA0 
         Caption         =   "A"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "SRQ High"
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bi-Directional Command"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   120
      TabIndex        =   42
      Top             =   5640
      Width           =   3975
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtBiDirectional 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   7
         Left            =   3480
         MaxLength       =   2
         TabIndex        =   50
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtBiDirectional 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   6
         Left            =   3000
         MaxLength       =   2
         TabIndex        =   49
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtBiDirectional 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   5
         Left            =   2520
         MaxLength       =   2
         TabIndex        =   48
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtBiDirectional 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   4
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   47
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtBiDirectional 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   3
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   46
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtBiDirectional 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   2
         Left            =   1080
         MaxLength       =   2
         TabIndex        =   45
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtBiDirectional 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   1
         Left            =   600
         MaxLength       =   2
         TabIndex        =   44
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtBiDirectional 
         Alignment       =   2  'Center
         Height          =   375
         Index           =   0
         Left            =   120
         MaxLength       =   2
         TabIndex        =   43
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblSendStatus 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   52
         ToolTipText     =   "Function Status"
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000E&
      Caption         =   "CheckSum"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   120
      TabIndex        =   40
      Top             =   4920
      Width           =   1575
      Begin VB.Label lblChkSum 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   240
         TabIndex        =   41
         ToolTipText     =   "Check Sum"
         Top             =   315
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delays"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   120
      TabIndex        =   33
      Top             =   3840
      Width           =   1575
      Begin VB.TextBox txtWriteDelay 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   600
         MaxLength       =   6
         TabIndex        =   35
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtReadDelay 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   600
         MaxLength       =   6
         TabIndex        =   34
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "ms"
         Height          =   255
         Left            =   1250
         TabIndex        =   39
         Top             =   640
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "ms"
         Height          =   255
         Left            =   1250
         TabIndex        =   38
         Top             =   280
         Width           =   255
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Write"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   640
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Read"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   280
         Width           =   375
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   11280
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H8000000E&
      Caption         =   "Slave"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1035
      Left            =   8520
      TabIndex        =   28
      Top             =   2160
      Width           =   1335
      Begin VB.VScrollBar VScroll1 
         Height          =   480
         Left            =   120
         Max             =   255
         Min             =   160
         TabIndex        =   30
         Top             =   360
         Value           =   160
         Width           =   350
      End
      Begin VB.TextBox txtSlave 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   480
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "A0"
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000E&
      Height          =   4815
      Left            =   8520
      TabIndex        =   24
      Top             =   3600
      Width           =   2295
      Begin VB.Frame Frame16 
         BackColor       =   &H8000000E&
         Caption         =   "Save KM > Japan"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1455
         Left            =   120
         TabIndex        =   91
         Top             =   2280
         Width           =   2055
         Begin VB.CommandButton cmdBLKEEP 
            Caption         =   "BLKEEP"
            Height          =   495
            Left            =   1080
            TabIndex        =   94
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdKMJapanDAT 
            Caption         =   "DAT"
            Height          =   495
            Left            =   120
            TabIndex        =   93
            ToolTipText     =   "xx;xx;xx;xx;xx..."
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton cmdKMJapanTX 
            Caption         =   "TX1/TX2"
            Height          =   495
            Left            =   120
            TabIndex        =   92
            ToolTipText     =   "000   FF FF FF FF FF FF..."
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H8000000E&
         Caption         =   "Load Japan > KM"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1455
         Left            =   120
         TabIndex        =   83
         Top             =   840
         Width           =   2055
         Begin VB.CommandButton cmdJapanTX2 
            Caption         =   "TX1/TX2"
            Height          =   495
            Left            =   120
            TabIndex        =   85
            ToolTipText     =   "000   FF FF FF FF FF FF..."
            Top             =   840
            Width           =   855
         End
         Begin VB.CommandButton cmdJapanDAT 
            Caption         =   "DAT"
            Height          =   495
            Left            =   120
            TabIndex        =   84
            ToolTipText     =   "xx;xx;xx;xx;xx..."
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdCompare 
         Caption         =   "Compare"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "Compare Function"
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdSaveFile 
         Height          =   495
         Left            =   720
         Picture         =   "frmI2CCheck.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Save File"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdLoadFile 
         Height          =   495
         Left            =   120
         Picture         =   "frmI2CCheck.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Load File"
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Go To Main"
         Top             =   4320
         Width           =   1095
      End
   End
   Begin VB.PictureBox shpBarMask 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1800
      ScaleHeight     =   195
      ScaleWidth      =   5580
      TabIndex        =   20
      Top             =   5400
      Width           =   5640
      Begin VB.Shape shpBar 
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   5655
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000E&
      Caption         =   "Function"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1575
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1575
      Begin VB.ComboBox cmbReadSize 
         Height          =   360
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "Number of Read Bytes"
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton cmdWrite 
         Caption         =   "Write"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Write To Set"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton cmdRead 
         Caption         =   "Read"
         Height          =   375
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Read From Set"
         Top             =   240
         Width           =   680
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Function Status"
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Memory Size"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2175
      Left            =   8520
      TabIndex        =   5
      Top             =   0
      Width           =   1335
      Begin VB.OptionButton optMemorySize 
         BackColor       =   &H8000000E&
         Caption         =   "128 K"
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   855
      End
      Begin VB.OptionButton optMemorySize 
         BackColor       =   &H8000000E&
         Caption         =   "64 K"
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton optMemorySize 
         BackColor       =   &H8000000E&
         Caption         =   "32 K"
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton optMemorySize 
         BackColor       =   &H8000000E&
         Caption         =   "16 K"
         Height          =   240
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optMemorySize 
         BackColor       =   &H8000000E&
         Caption         =   "8 K"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton optMemorySize 
         BackColor       =   &H8000000E&
         Caption         =   "4 K"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optMemorySize 
         BackColor       =   &H8000000E&
         Caption         =   "2 K"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5295
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   6650
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   4905
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Width           =   6405
         _cx             =   11298
         _cy             =   8643
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   14737632
         ForeColor       =   -2147483640
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   14737632
         BackColorAlternate=   14737632
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3360
         Top             =   2280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "I2C Box"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2295
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000E&
         Caption         =   "I2C Select"
         Height          =   1095
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
         Begin VB.OptionButton optBus 
            BackColor       =   &H8000000E&
            Caption         =   "I2C 1"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   "I2C Bus 1"
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton optBus 
            BackColor       =   &H8000000E&
            Caption         =   "I2C 0"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "I2C Bus 0"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblBusSelect 
            Alignment       =   2  'Center
            BackColor       =   &H80000012&
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "I2C Select Status"
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.OptionButton optSRQ 
         Caption         =   "SRQ High"
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "SRQ High"
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optSRQ 
         Caption         =   "SRQ Low"
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "SRQ Low"
         Top             =   650
         Width           =   1335
      End
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Left            =   7560
      Top             =   5350
      Width           =   855
   End
   Begin VB.Label lblTact 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   7560
      TabIndex        =   23
      ToolTipText     =   "Tact Time"
      Top             =   5350
      Width           =   855
   End
End
Attribute VB_Name = "frmI2CCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SlvNo As Integer

Private Sub InitGrdData()
Dim i As Integer

grdData.Clear
grdData.Cols = 17
grdData.Rows = 1

grdData.ColWidth(0) = 550
grdData.ColAlignment(0) = flexAlignCenterCenter

For i = 1 To 16
    grdData.ColWidth(i) = 350
    grdData.ColAlignment(i) = flexAlignCenterCenter
Next i
For i = 0 To 15
    grdData.TextMatrix(0, i + 1) = "0" & Hex(i)
Next i

End Sub

Private Sub chkDataEnable_Click(Index As Integer)

If chkDataEnable(Index).Value = 1 Then
    txtData(Index).Enabled = True
    txtData(Index).BackColor = vbWhite
Else
    txtData(Index).Enabled = False
    txtData(Index).BackColor = &H8000000A
End If

End Sub

Private Sub cmdBLKEEP_Click()
Dim tmpstr As String

If grdData.Rows < 2 Then Exit Sub

CommonDialog1.FileName = ""
CommonDialog1.Filter = "DAT Files (*.dat)|*.dat"
CommonDialog1.DialogTitle = "Save KM to Japan Format (BLKEEP)"
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowSave

If CommonDialog1.FileName = "" Then Exit Sub

Me.MousePointer = vbHourglass
DoEvents

Open CommonDialog1.FileName For Output As #1

'Print #1, Space(14) & "EEPROM DAMP DATA   SIZE=" & SlvNo & "KB"
'Print #1, Space(7) & "00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F"

For j = 1 To grdData.Rows - 1
    tmpstr = ""
    For i = 1 To 16
        tmpstr = tmpstr & grdData.TextMatrix(j, i)
    Next i
    tmpstr = tmpstr & vbCr
    Print #1, tmpstr
Next j

Close #1

Me.MousePointer = vbNormal

End Sub

Private Sub cmdCompare_Click()
If grdData.Rows < 3 Then
    MsgBox "Table is empty!", vbExclamation, "Compare"
    Exit Sub
End If

k = 0
For i = 1 To grdData.Rows - 1
    For j = 1 To 16
        ReDim Preserve CompareTable1(k)
        CompareTable1(k) = Trim(grdData.TextMatrix(i, j))
        k = k + 1
    Next j
Next i

frmCompareWith.Show vbModal

End Sub

Private Sub cmdConvertJapanKM_Click()
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdJapanTX2_Click()
Dim textline As String
Dim tmpstr As String
Dim mSlave As Integer
Dim mSub As Integer
Dim MSB As String
Dim LSB As String
Dim pointSize
Dim mCount As Integer
Dim mSize As String
Dim mSize_INT As Integer
Dim mFind As Integer

CommonDialog2.FileName = ""
CommonDialog2.Filter = "TX1 Files (*.tx1)|*.tx1|TX2 Files (*.tx2)|*.tx2"
CommonDialog2.DialogTitle = "Open Japan Format"
CommonDialog2.InitDir = App.Path
CommonDialog2.ShowOpen

Debug.Print CommonDialog2.FileName
If CommonDialog2.FileName = "" Then Exit Sub

Me.MousePointer = vbHourglass

InitGrdData

Open CommonDialog2.FileName For Input As #1

mSlave = 0
mSub = 0
mCount = 0
mSize = ""

Do While Not EOF(1) ' Loop until end of file.
    Line Input #1, textline ' Read line into variable.
    
    '''Check File Size
    If mCount = 0 Then
        mCount = 1
        pointSize = InStr(1, textline, "=")
        If pointSize = 0 Then
            MsgBox "Cannot Find Data Size!", vbExclamation, "TX1/TX2 Format"
            Close #1: Me.MousePointer = vbNormal: Exit Sub
        End If
        mSize = Mid(textline, pointSize + 1)
        mSize_INT = Val(mSize)
        mFind = 0
        For i = 0 To 4
            If mSize_INT = 8 * 2 ^ i Then mFind = 1: Exit For
        Next i
        If mFind = 0 Then
            MsgBox "Invalid Data Size Format!", vbExclamation, "TX1/TX2 Format"
            Close #1: Me.MousePointer = vbNormal: Exit Sub
        End If
    End If
    
    If Mid(textline, 1, 1) = " " Then
        GoTo mSKIP
    End If
    
    grdData.Rows = grdData.Rows + 1
    If mSize_INT > 32 Then
        tmpstr = Mid(textline, 8, 47)   '''extract 1 row of data (16 bytes)
        grdData.TextMatrix(grdData.Rows - 1, 0) = Mid(textline, 1, 4)
    ElseIf mSize_INT <= 32 Then
        tmpstr = Mid(textline, 7, 47)
        grdData.TextMatrix(grdData.Rows - 1, 0) = "0" & Mid(textline, 1, 3)
    End If
    'MSB = Mid(textline, 1, 2)
    'LSB = Mid(textline, 3, 1) & Hex(mSub)
    
    'grdData.TextMatrix(grdData.Rows - 1, 0) = "0" & Mid(textline, 1, 3)
    
    j = 1
    For i = 0 To 15
        grdData.TextMatrix(grdData.Rows - 1, i + 1) = Mid(tmpstr, j, 2)
        j = j + 3
    Next i
    
    mSub = mSub + 1
    If mSub = 16 Then
        mSlave = mSlave + 1
        mSub = 0
    End If

mSKIP:
Loop

Close #1

For i = 0 To 6
If optMemorySize(i).Caption = (grdData.Rows - 1) / 8 & " K" Then
    optMemorySize(i).Value = 1
End If
Next i

Me.MousePointer = vbNormal

End Sub

Private Sub cmdKMJapanTX_Click()
Dim tmpstr As String

If grdData.Rows < 2 Then Exit Sub

CommonDialog1.FileName = ""
CommonDialog1.Filter = "TX1 Files (*.tx1)|*.tx1|TX2 Files (*.tx2)|*.tx2"
CommonDialog1.DialogTitle = "Save KM to Japan Format"
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowSave

If CommonDialog1.FileName = "" Then Exit Sub

Me.MousePointer = vbHourglass
DoEvents

Open CommonDialog1.FileName For Output As #1

Print #1, Space(14) & "EEPROM DAMP DATA   SIZE=" & SlvNo & "KB"
Print #1, Space(7) & "00 01 02 03 04 05 06 07 08 09 0A 0B 0C 0D 0E 0F"

For j = 1 To grdData.Rows - 1
    tmpstr = grdData.TextMatrix(j, 0) & Space(3)
    For i = 1 To 16
        tmpstr = tmpstr & grdData.TextMatrix(j, i) & " "
    Next i
    Print #1, tmpstr
Next j

Close #1

Me.MousePointer = vbNormal

End Sub

Private Sub cmdLoadFile_Click()
Dim textline As String
Dim tmpstr As String
Dim ChkSum As Long

CommonDialog1.FileName = ""
CommonDialog1.Filter = "Data Files (*.dat,*.ver)|*.dat;*.ver|All Files (*.*)|*.*"
CommonDialog1.DialogTitle = "Load Data from File"
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then Exit Sub

lblChkSum.Caption = ""
Me.MousePointer = vbHourglass
DoEvents

InitGrdData

Open CommonDialog1.FileName For Input As #1

ChkSum = 0

Do While Not EOF(1) ' Loop until end of file.
    Line Input #1, textline ' Read line into variable.
    tmpstr = Mid(textline, 7)
    If Len(tmpstr) > 40 Then
        grdData.Rows = grdData.Rows + 1
        grdData.TextMatrix(grdData.Rows - 1, 0) = Mid(textline, 2, 4)
        j = 1
        For i = 1 To 48 Step 3
            grdData.TextMatrix(grdData.Rows - 1, j) = Mid(tmpstr, i, 2)
            ChkSum = ChkSum + Val("&H" & Mid(tmpstr, i, 2))
            j = j + 1
        Next i
    End If
Loop

Close #1

Me.MousePointer = vbNormal

For i = 0 To 6
If optMemorySize(i).Caption = (grdData.Rows - 1) / 8 & " K" Then
    optMemorySize(i).Value = 1
End If
Next i

lblChkSum.Caption = Hex(ChkSum)

End Sub

Private Sub cmdJapanDAT_Click()
Dim textline As String
Dim tmpstr() As String
Dim mSlave As Integer
Dim mSub As Integer
Dim MSB As String
Dim LSB As String

CommonDialog2.FileName = ""
CommonDialog2.Filter = "Dat Files (*.dat)|*.dat"
CommonDialog2.DialogTitle = "Open Japan Format"
CommonDialog2.InitDir = App.Path
CommonDialog2.ShowOpen

If CommonDialog2.FileName = "" Then Exit Sub

Me.MousePointer = vbHourglass

InitGrdData

Open CommonDialog2.FileName For Input As #1

mSlave = 0
mSub = 0

Do While Not EOF(1) ' Loop until end of file.
    Line Input #1, textline ' Read line into variable.
    grdData.Rows = grdData.Rows + 1
    tmpstr = Split(textline, ";")
    
    MSB = Hex(mSlave)
    If Len(MSB) = 1 Then MSB = "0" & MSB
    LSB = Hex(mSub) & "0"
    grdData.TextMatrix(grdData.Rows - 1, 0) = MSB & LSB
    
    If UBound(tmpstr) <> 16 Then
        MsgBox "Invalid Data", vbCritical, "Open Japan Format"
        Close #1
        Me.MousePointer = vbNormal
        Exit Sub
    End If
    
    For i = 0 To 15
        grdData.TextMatrix(grdData.Rows - 1, i + 1) = tmpstr(i)
    Next i
    
    mSub = mSub + 1
    If mSub = 16 Then
        mSlave = mSlave + 1
        mSub = 0
    End If
    
Loop

Close #1

For i = 0 To 6
If optMemorySize(i).Caption = (grdData.Rows - 1) / 8 & " K" Then
    optMemorySize(i).Value = 1
End If
Next i

Me.MousePointer = vbNormal

End Sub

Private Sub cmdRead_Click()

Dim I2cTrans As I2C_TRANS              ' Dimension an I2C_TRANS structure
Dim ChkSum As Long
Dim ReadData As String
Dim hAdd As String
Dim lAdd As Integer
Dim shpBarSTEP As Integer
Dim lapSTART As Long
Dim mDelay As Long

lblTact.Caption = ""
lblSTATUS.Caption = ""
lblChkSum.Caption = ""
DoEvents

InitGrdData

ChkSum = 0
shpBar.Width = 0
shpBar.FillColor = vbBlue
lapSTART = Timer
mDelay = txtReadDelay.Text

If cmbReadSize.Text = "16" Then

    shpBarSTEP = shpBarMask.Width / (SlvNo * 16)
    
    If SlvNo > 8 Then I2cTrans.byType = I2C_TRANS_16ADR Else I2cTrans.byType = I2C_TRANS_8ADR
    I2cTrans.wCount.hi = 0
    I2cTrans.wCount.lo = &H10
    
    grdData.Rows = grdData.Rows + 1
    
    For i = 0 To SlvNo - 1
    
    If SlvNo > 8 Then I2cTrans.byDevId = "&H" & txtSlave.Text Else I2cTrans.byDevId = "&HA" & Hex(i * 2)
    I2cTrans.wMemAddr.hi = "&H" & Hex(i)
    lAdd = 0
    
        For j = 0 To 15
        
            I2cTrans.wMemAddr.lo = "&H" & Hex(j) & "0"
                        
            mCount = 0
            Do While DAPI_ReadI2c(hDevInstance, I2cTrans) <> 16
                mCount = mCount + 1
                If mCount >= 10 Then
                    lblSTATUS.ForeColor = vbRed
                    lblSTATUS.Caption = "I2C Error!"
                    shpBar.FillColor = vbRed
                    Exit Sub
                End If
            Loop
            delay_ms mDelay
            
            For k = 0 To 15
                ChkSum = ChkSum + I2cTrans.Data(k)
                ReadData = Hex(I2cTrans.Data(k))
                If Len(ReadData) = 1 Then ReadData = "0" & ReadData
                grdData.TextMatrix(grdData.Rows - 1, k + 1) = ReadData
            Next k
            
            hAdd = Hex(i)
            If Len(hAdd) = 1 Then hAdd = "0" & hAdd
            grdData.TextMatrix(grdData.Rows - 1, 0) = hAdd & Hex(j) & "0"
            grdData.Rows = grdData.Rows + 1
            shpBar.Width = shpBar.Width + shpBarSTEP
            
        Next j
        
    Next i

ElseIf cmbReadSize.Text = "256" Then

    shpBarSTEP = shpBarMask.Width / (SlvNo)
    
    If SlvNo > 8 Then I2cTrans.byType = I2C_TRANS_16ADR Else I2cTrans.byType = I2C_TRANS_8ADR
    
    I2cTrans.wCount.hi = &H1
    I2cTrans.wCount.lo = &H0
    
    grdData.Rows = grdData.Rows + 1
    
    For i = 0 To SlvNo - 1
    
        If SlvNo > 8 Then I2cTrans.byDevId = "&H" & txtSlave.Text Else I2cTrans.byDevId = "&HA" & Hex(i * 2)
        I2cTrans.wMemAddr.hi = "&H" & Hex(i)
        I2cTrans.wMemAddr.lo = &H0
        lAdd = 0
    
        mCount = 0
        Do While DAPI_ReadI2c(hDevInstance, I2cTrans) <> 256
            mCount = mCount + 1
            If mCount >= 10 Then
                lblSTATUS.ForeColor = vbRed
                lblSTATUS.Caption = "I2C Error!"
                shpBar.FillColor = vbRed
                Exit Sub
            End If
        Loop
        delay_ms mDelay
        
        hAdd = Hex(i)
        If Len(hAdd) = 1 Then hAdd = "0" & hAdd
        
        j = 0
        l = 0
        For k = 0 To 255
            ChkSum = ChkSum + I2cTrans.Data(k)
            ReadData = Hex(I2cTrans.Data(k))
            If Len(ReadData) = 1 Then ReadData = "0" & ReadData
            grdData.TextMatrix(grdData.Rows - 1, j + 1) = ReadData
            If j = 0 Then grdData.TextMatrix(grdData.Rows - 1, 0) = hAdd & Hex(l) & "0"
            j = j + 1
            If j = 16 Then
                j = 0
                grdData.Rows = grdData.Rows + 1
                l = l + 1
                If l = 16 Then l = 0
            End If
        Next k
            
        shpBar.Width = shpBar.Width + shpBarSTEP
            
    Next i

End If

lblSTATUS.ForeColor = vbGreen
lblSTATUS.Caption = "Read OK!"
lblChkSum.Caption = Hex(ChkSum)
shpBar.FillColor = vbGreen
grdData.Rows = grdData.Rows - 1
lblTact.Caption = Format((Timer - lapSTART), "0.00") & " s"
'txtSlave.Text = "A0"

End Sub

Private Sub cmdRead1_Click()
Dim I2cTrans As I2C_TRANS              ' Dimension an I2C_TRANS structure
Dim ReadData As String
Dim byteSend As Integer
   
lblStatus2.Caption = ""
DoEvents

If txtSubAdd1.Text <> "" And txtSubAdd2.Text <> "" Then
    I2cTrans.byType = I2C_TRANS_16ADR
    I2cTrans.wMemAddr.hi = CByte("&H" & txtSubAdd1.Text)
    I2cTrans.wMemAddr.lo = CByte("&H" & txtSubAdd2.Text)
ElseIf txtSubAdd1.Text <> "" And txtSubAdd2.Text = "" Then
    I2cTrans.byType = I2C_TRANS_8ADR
    I2cTrans.wMemAddr.hi = 0
    I2cTrans.wMemAddr.lo = CByte("&H" & txtSubAdd1.Text)
ElseIf txtSubAdd1.Text = "" And txtSubAdd2.Text = "" Then
    I2cTrans.byType = I2C_TRANS_NOADR
    I2cTrans.wMemAddr.hi = 0
    I2cTrans.wMemAddr.lo = 0
Else
    MsgBox "Syntax error!"
    Exit Sub
End If

byteSend = 0

If chkDataEnable(0).Value = 1 Then byteSend = byteSend + 1 Else GoTo skip1
If chkDataEnable(1).Value = 1 Then byteSend = byteSend + 1 Else GoTo skip1
If chkDataEnable(2).Value = 1 Then byteSend = byteSend + 1 Else GoTo skip1
If chkDataEnable(3).Value = 1 Then byteSend = byteSend + 1 Else GoTo skip1
    
skip1:

If byteSend = 0 Then MsgBox "Syntax Error!": Exit Sub

I2cTrans.wCount.hi = 0
I2cTrans.wCount.lo = byteSend
    
I2cTrans.byDevId = CByte("&H" & txtSlaveAdd.Text)
                         
If DAPI_ReadI2c(hDevInstance, I2cTrans) <> byteSend Then
    lblStatus2.ForeColor = vbRed
    lblStatus2.Caption = "I2C Error!"
    Exit Sub
End If

For i = 0 To byteSend - 1
    txtData(i).Text = Hex(I2cTrans.Data(i))
Next i

lblStatus2.ForeColor = vbGreen
lblStatus2.Caption = "Read OK!"

End Sub

Private Sub cmdReadBKSV_Click()
Dim I2cTrans As I2C_TRANS              ' Dimension an I2C_TRANS structure
Dim tmpRead As String

txtBKSV.Text = ""
DoEvents

I2cTrans.byType = I2C_TRANS_8ADR
I2cTrans.wCount.hi = &H0
I2cTrans.wCount.lo = &H5

I2cTrans.byDevId = "&H" & txtBKSVslave.Text    '&H74  '
I2cTrans.wMemAddr.hi = &H0
I2cTrans.wMemAddr.lo = &H0
 
If DAPI_ReadI2c(hDevInstance, I2cTrans) <> 5 Then
    txtBKSV.Text = "I2C ERROR!"
    Exit Sub
End If

For i = 0 To 4
    tmpRead = Hex(I2cTrans.Data(i))
    If Len(tmpRead) = 1 Then tmpRead = "0" & tmpRead
    txtBKSV.Text = txtBKSV.Text & tmpRead
Next i

End Sub

Private Sub cmdReset_Click()
Dim t As Boolean
Dim u As Boolean
Dim mReq As Byte
Dim mInd As Long
Dim mVal As Long
Dim mLen As Long
Dim getData As Byte

mReq = &HE4    '// 0xE4 = VR_System
mInd = &H400      '// MSB = Reset, LSB = unused
mVal = &H0        '// MSB = unused, LSB = unused
mLen = 0         '// no data phase
    

t = DAPI_SetVendorRequest(hDevInstance, mReq, mVal, mInd, mLen, 0)

delay_ms 5000

OpenDevice

mLen = 1
t = DAPI_GetVendorRequest(hDevInstance, getData, mReq, mVal, mInd, mLen)
If getData = 16 Then
    MsgBox "Board was successfully reset!"
Else
    MsgBox "Board was not successfully reset!"
End If

End Sub

Private Sub cmdSaveFile_Click()

Dim tmpstr As String

If grdData.Rows < 2 Then Exit Sub

CommonDialog1.FileName = ""
CommonDialog1.Filter = "Data Files (*.dat)|*.dat"
CommonDialog1.DialogTitle = "Save Data to File"
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowSave

If CommonDialog1.FileName = "" Then Exit Sub


Me.MousePointer = vbHourglass
DoEvents

Open CommonDialog1.FileName For Output As #1

Print #1, "M" & UCase(Mid(CommonDialog1.FileTitle, 1, InStr(1, CommonDialog1.FileTitle, ".") - 1))

For j = 1 To grdData.Rows - 1
    tmpstr = ":" & grdData.TextMatrix(j, 0) & " "
    For i = 1 To 16
        tmpstr = tmpstr & grdData.TextMatrix(j, i) & " "
    Next i
    Print #1, tmpstr
Next j

Print #1, "@@"

Close #1

Me.MousePointer = vbNormal

End Sub

Private Sub cmdSend_Click()

Dim I2cTrans As I2C_TRANS                     ' Dimension an I2C_TRANS structure
Dim lWritten As Long
Dim biData() As Integer
Dim byteSend As Integer

lblSendStatus.Caption = ""

For i = 0 To 7
    If txtBiDirectional(i) <> "" Then
        ReDim Preserve biData(i)
        biData(i) = "&H" & txtBiDirectional(i).Text
    End If
Next i

byteSend = UBound(biData)

I2cTrans.byDevId = biData(0)
I2cTrans.byType = I2C_TRANS_NOADR
I2cTrans.wMemAddr.hi = 0
I2cTrans.wMemAddr.lo = 0
I2cTrans.wCount.hi = 0
I2cTrans.wCount.lo = byteSend

For j = 1 To UBound(biData)
    I2cTrans.Data(j - 1) = biData(j)
Next j

lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
If lWritten <> byteSend Then
    lblSendStatus.ForeColor = vbRed
    lblSendStatus.Caption = "Send NG!"
    Exit Sub
End If

lblSendStatus.ForeColor = vbGreen
lblSendStatus.Caption = "Send OK!"

''Status = tvif_iic_wrt(tvif_add, &HFE, &H80, 0,  0, 2)
'''Ack

'I2cTrans.byDevId = &HFE
'I2cTrans.byType = I2C_TRANS_8ADR
'I2cTrans.wMemAddr.hi = 0
'I2cTrans.wMemAddr.lo = &H80
'I2cTrans.wCount.hi = 0
'I2cTrans.wCount.lo = 2 '

'I2cTrans.Data(0) = &H0
'I2cTrans.Data(1) = &H0

'lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
'If lWritten <> 2 Then
'    lblSendStatus.ForeColor = vbRed
'    lblSendStatus.Caption = "ACK NG!"
'    Exit Sub
'End If

End Sub

Private Sub cmdWP1_Click()
Dim I2cTrans As I2C_TRANS                     ' Dimension an I2C_TRANS structure
Dim lWritten As Long
Dim byteSend As Integer

lblSendStatus.Caption = ""

byteSend = 2

I2cTrans.byDevId = &H70
I2cTrans.byType = I2C_TRANS_NOADR
I2cTrans.wMemAddr.hi = 0
I2cTrans.wMemAddr.lo = 0
I2cTrans.wCount.hi = 0
I2cTrans.wCount.lo = byteSend

I2cTrans.Data(0) = &H88
I2cTrans.Data(1) = &H0

lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
If lWritten <> byteSend Then
    lblSendStatus.ForeColor = vbRed
    lblSendStatus.Caption = "Send NG!"
    Exit Sub
End If

lblSendStatus.ForeColor = vbGreen
lblSendStatus.Caption = "Send OK!"

End Sub

Private Sub cmdWP2_Click()
Dim I2cTrans As I2C_TRANS                     ' Dimension an I2C_TRANS structure
Dim lWritten As Long
Dim byteSend As Integer

lblSendStatus.Caption = ""

byteSend = 2

I2cTrans.byDevId = &H70
I2cTrans.byType = I2C_TRANS_NOADR
I2cTrans.wMemAddr.hi = 0
I2cTrans.wMemAddr.lo = 0
I2cTrans.wCount.hi = 0
I2cTrans.wCount.lo = byteSend

I2cTrans.Data(0) = &H8B
I2cTrans.Data(1) = &H0

lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
If lWritten <> byteSend Then
    lblSendStatus.ForeColor = vbRed
    lblSendStatus.Caption = "Send NG!"
    Exit Sub
End If

lblSendStatus.ForeColor = vbGreen
lblSendStatus.Caption = "Send OK!"

End Sub

Private Sub cmdWrite_Click()
  
Dim I2cTrans As I2C_TRANS                     ' Dimension an I2C_TRANS structure
Dim lWritten As Long
Dim MyStep As Integer
Dim byteSend As Integer
Dim mCol, mRow As Integer
Dim Add As Integer
Dim mDelay As Long


shpBar.FillColor = vbBlue
Me.lblSTATUS.Caption = ""
lblChkSum.Caption = ""
DoEvents

If grdData.Rows <= 2 Then Exit Sub

'''max send byte:
'''2K   -   8 bytes
'''4K   -   16 bytes
'''16K  -   16 bytes
'''32K  -   32 bytes
'''64K  -   32 bytes
'''128K -   64 bytes

Select Case SlvNo

Case 1  '''2K
    MyStep = 1
    byteSend = 8
Case 2  '''4K
    MyStep = 1
    byteSend = 16
Case 4  '''8K
    MyStep = 1
    byteSend = 16
Case 8  '''16K
    MyStep = 1
    byteSend = 16
Case 16 '''32K
    MyStep = 2
    byteSend = 32
Case 32 '''64K
    MyStep = 2
    byteSend = 32
Case 64 '''128K
    MyStep = 2
    byteSend = 32
End Select

I2cTrans.wCount.hi = 0
I2cTrans.wCount.lo = byteSend

Add = 0
mRow = 1
mCol = 1
mDelay = txtWriteDelay.Text

For i = 0 To SlvNo - 1

    For j = 0 To 15 Step MyStep

        k = 0
        
Send8ByteLoop:
        
        For h = 0 To byteSend - 1
            I2cTrans.Data(h) = CByte(Val("&H" & grdData.TextMatrix(mRow, mCol)))
            mCol = mCol + 1
            If mCol = 17 Then mCol = 1: mRow = mRow + 1
        Next h
        
        If SlvNo <= 8 Then '''16K & below
                        
            I2cTrans.byDevId = "&HA" & Hex(i * 2)
            I2cTrans.byType = I2C_TRANS_8ADR
            I2cTrans.wMemAddr.hi = "&H" & Hex(i)
            If byteSend = 16 Or byteSend = 32 Then
                I2cTrans.wMemAddr.lo = ("&H" & Hex(j)) * &H10
            ElseIf byteSend = 8 Then
                If k = 0 Then I2cTrans.wMemAddr.lo = ("&H" & (Hex(j) & 0))
                If k = 1 Then I2cTrans.wMemAddr.lo = ("&H" & (Hex(j) & 8))
            End If
            
            lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
            If lWritten <> byteSend Then
                delay_ms 50
                lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
                If lWritten <> byteSend Then
                    GoTo Err
                End If
            End If
                        
        ElseIf SlvNo > 8 Then '''32K & above
        
            I2cTrans.byDevId = "&H" & txtSlave.Text
            I2cTrans.byType = I2C_TRANS_16ADR   '''2bytes
            I2cTrans.wMemAddr.hi = "&H" & Hex(i)
                        
            If byteSend = 16 Or byteSend = 32 Then
                I2cTrans.wMemAddr.lo = ("&H" & Hex(j)) * &H10
            ElseIf byteSend = 8 Then
                If k = 0 Then I2cTrans.wMemAddr.lo = ("&H" & (Hex(j) & 0))
                If k = 1 Then I2cTrans.wMemAddr.lo = ("&H" & (Hex(j) & 8))
            End If
            
            lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
            If lWritten <> byteSend Then
                delay_ms 50
                lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
                If lWritten <> byteSend Then
                    GoTo Err
                End If
            End If
            
        End If
        
        shpBar.Width = (Add / ((SlvNo) * 256)) * shpBarMask.Width
        
        Add = Add + byteSend
        
        'downloadloop = 10000
        'e = 0
        'For n = 0 To downloadloop
        '    e = e + 1
        'Next n
        delay_ms mDelay
        
        If byteSend = 8 And k = 0 Then
            k = 1
            GoTo Send8ByteLoop
        End If
        
    Next j
Next i

shpBar.FillColor = vbGreen
lblSTATUS.ForeColor = vbGreen
lblSTATUS.Caption = "Write OK!"

Exit Sub

Err:

lblSTATUS.ForeColor = vbRed
lblSTATUS.Caption = "I2C Error!"
shpBar.FillColor = vbRed

End Sub

Private Sub cmdWrite00_Click()
Dim I2cTrans As I2C_TRANS                     ' Dimension an I2C_TRANS structure
Dim lWritten As Long
Dim MyStep As Integer
Dim byteSend As Integer
Dim Add As Integer
Dim mDelay As Long

shpBar.FillColor = vbBlue
Me.lblSTATUS.Caption = ""
lblChkSum.Caption = ""
DoEvents

'''max send byte:
'''2K   -   8 bytes
'''4K   -   16 bytes
'''16K  -   16 bytes
'''32K  -   32 bytes
'''64K  -   32 bytes
'''128K -   64 bytes

Select Case SlvNo

Case 1  '''2K
    MyStep = 1
    byteSend = 8
Case 2  '''4K
    MyStep = 1
    byteSend = 16
Case 4  '''8K
    MyStep = 1
    byteSend = 16
Case 8  '''16K
    MyStep = 1
    byteSend = 16
Case 16 '''32K
    MyStep = 2
    byteSend = 32
Case 32 '''64K
    MyStep = 2
    byteSend = 32
Case 64 '''128K
    MyStep = 2
    byteSend = 32
End Select

I2cTrans.wCount.hi = 0
I2cTrans.wCount.lo = byteSend

Add = 0
mDelay = txtWriteDelay.Text

For i = 0 To SlvNo - 1

    For j = 0 To 15 Step MyStep

        k = 0
        
Send8ByteLoop:
        
        For h = 0 To byteSend - 1
            I2cTrans.Data(h) = &H0
        Next h
        
        If SlvNo <= 8 Then '''16K & below
                        
            I2cTrans.byDevId = "&HA" & Hex(i * 2)
            I2cTrans.byType = I2C_TRANS_8ADR
            I2cTrans.wMemAddr.hi = "&H" & Hex(i)
            If byteSend = 16 Then
                I2cTrans.wMemAddr.lo = ("&H" & Hex(j)) * &H10
            ElseIf byteSend = 8 Then
                If k = 0 Then I2cTrans.wMemAddr.lo = ("&H" & (Hex(j) & 0))
                If k = 1 Then I2cTrans.wMemAddr.lo = ("&H" & (Hex(j) & 8))
            End If
            
            lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
            If lWritten <> byteSend Then
                delay_ms 50
                lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
                If lWritten <> byteSend Then
                    GoTo Err
                End If
            End If
                        
        ElseIf SlvNo > 8 Then '''32K & above
        
            I2cTrans.byDevId = "&H" & txtSlave.Text
            I2cTrans.byType = I2C_TRANS_16ADR   '''2bytes
            I2cTrans.wMemAddr.hi = "&H" & Hex(i)
            I2cTrans.wMemAddr.lo = ("&H" & Hex(j)) * &H10
            
            lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
            If lWritten <> byteSend Then
                delay_ms 50
                lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
                If lWritten <> byteSend Then
                    GoTo Err
                End If
            End If
            
        End If
        
        shpBar.Width = (Add / ((SlvNo) * 256)) * shpBarMask.Width
        
        Add = Add + byteSend
        
        'downloadloop = 10000
        'e = 0
        'For n = 0 To downloadloop
        '    e = e + 1
        'Next n
        delay_ms mDelay
        
        If byteSend = 8 And k = 0 Then
            k = 1
            GoTo Send8ByteLoop
        End If
        
    Next j
Next i

shpBar.FillColor = vbGreen
lblSTATUS.ForeColor = vbGreen
lblSTATUS.Caption = "Write OK!"

Exit Sub

Err:

lblSTATUS.ForeColor = vbRed
lblSTATUS.Caption = "I2C Error!"
shpBar.FillColor = vbRed

End Sub

Private Sub cmdWrite1_Click()
Dim I2cTrans As I2C_TRANS                     ' Dimension an I2C_TRANS structure
Dim lWritten As Long
Dim MyStep As Integer
Dim byteSend As Integer
Dim Add As Integer
Dim mDelay As Long
Dim splitData() As String

lblStatus2.Caption = ""
DoEvents

If txtSubAdd1.Text <> "" And txtSubAdd2.Text <> "" Then
    I2cTrans.byType = I2C_TRANS_16ADR
    I2cTrans.wMemAddr.hi = CByte("&H" & txtSubAdd1.Text)
    I2cTrans.wMemAddr.lo = CByte("&H" & txtSubAdd2.Text)
ElseIf txtSubAdd1.Text <> "" And txtSubAdd2.Text = "" Then
    I2cTrans.byType = I2C_TRANS_8ADR
    I2cTrans.wMemAddr.hi = 0
    I2cTrans.wMemAddr.lo = CByte("&H" & txtSubAdd1.Text)
ElseIf txtSubAdd1.Text = "" And txtSubAdd2.Text = "" Then
    I2cTrans.byType = I2C_TRANS_NOADR
    I2cTrans.wMemAddr.hi = 0
    I2cTrans.wMemAddr.lo = 0
Else
    MsgBox "Syntax error!"
    Exit Sub
End If

byteSend = 0

If chkDataEnable(0).Value = 1 Then byteSend = byteSend + 1 Else GoTo skip1
If chkDataEnable(1).Value = 1 Then byteSend = byteSend + 1 Else GoTo skip1
If chkDataEnable(2).Value = 1 Then byteSend = byteSend + 1 Else GoTo skip1
If chkDataEnable(3).Value = 1 Then byteSend = byteSend + 1 Else GoTo skip1
    
skip1:

If byteSend = 0 Then MsgBox "Syntax Error!": Exit Sub

I2cTrans.wCount.hi = 0
I2cTrans.wCount.lo = byteSend
       
I2cTrans.byDevId = CByte("&H" & txtSlaveAdd.Text)
                     
For i = 0 To byteSend - 1
   I2cTrans.Data(i) = CByte("&H" & txtData(i).Text)
Next i
                     
If DAPI_WriteI2c(hDevInstance, I2cTrans) <> byteSend Then
    lblStatus2.ForeColor = vbRed
    lblStatus2.Caption = "I2C Error!"
    Exit Sub
End If

lblStatus2.ForeColor = vbGreen
lblStatus2.Caption = "Write OK!"

End Sub

Private Sub cmdWriteFF_Click()
Dim I2cTrans As I2C_TRANS                     ' Dimension an I2C_TRANS structure
Dim lWritten As Long
Dim MyStep As Integer
Dim byteSend As Integer
Dim Add As Integer
Dim mDelay As Long

shpBar.FillColor = vbBlue
Me.lblSTATUS.Caption = ""
lblChkSum.Caption = ""
DoEvents

'''max send byte:
'''2K   -   8 bytes
'''4K   -   16 bytes
'''16K  -   16 bytes
'''32K  -   32 bytes
'''64K  -   32 bytes
'''128K -   64 bytes

Select Case SlvNo

Case 1  '''2K
    MyStep = 1
    byteSend = 8
Case 2  '''4K
    MyStep = 1
    byteSend = 16
Case 4  '''8K
    MyStep = 1
    byteSend = 16
Case 8  '''16K
    MyStep = 1
    byteSend = 16
Case 16 '''32K
    MyStep = 2
    byteSend = 32
Case 32 '''64K
    MyStep = 2
    byteSend = 32
Case 64 '''128K
    MyStep = 2
    byteSend = 32
End Select

I2cTrans.wCount.hi = 0
I2cTrans.wCount.lo = byteSend

Add = 0
mDelay = txtWriteDelay.Text

For i = 0 To SlvNo - 1

    For j = 0 To 15 Step MyStep

        k = 0
        
Send8ByteLoop:
        
        For h = 0 To byteSend - 1
            I2cTrans.Data(h) = &HFF
        Next h
        
        If SlvNo <= 8 Then '''16K & below
                        
            I2cTrans.byDevId = "&HA" & Hex(i * 2)
            I2cTrans.byType = I2C_TRANS_8ADR
            I2cTrans.wMemAddr.hi = "&H" & Hex(i)
            If byteSend = 16 Then
                I2cTrans.wMemAddr.lo = ("&H" & Hex(j)) * &H10
            ElseIf byteSend = 8 Then
                If k = 0 Then I2cTrans.wMemAddr.lo = ("&H" & (Hex(j) & 0))
                If k = 1 Then I2cTrans.wMemAddr.lo = ("&H" & (Hex(j) & 8))
            End If
            
            lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
            If lWritten <> byteSend Then
                delay_ms 50
                lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
                If lWritten <> byteSend Then
                    GoTo Err
                End If
            End If
                        
        ElseIf SlvNo > 8 Then '''32K & above
        
            I2cTrans.byDevId = "&H" & txtSlave.Text
            I2cTrans.byType = I2C_TRANS_16ADR   '''2bytes
            I2cTrans.wMemAddr.hi = "&H" & Hex(i)
            I2cTrans.wMemAddr.lo = ("&H" & Hex(j)) * &H10
            
            lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
            If lWritten <> byteSend Then
                delay_ms 50
                lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
                If lWritten <> byteSend Then
                    GoTo Err
                End If
            End If
            
        End If
        
        shpBar.Width = (Add / ((SlvNo) * 256)) * shpBarMask.Width
        
        Add = Add + byteSend
        
        'downloadloop = 10000
        'e = 0
        'For n = 0 To downloadloop
        '    e = e + 1
        'Next n
        delay_ms mDelay
        
        If byteSend = 8 And k = 0 Then
            k = 1
            GoTo Send8ByteLoop
        End If
        
    Next j
Next i

shpBar.FillColor = vbGreen
lblSTATUS.ForeColor = vbGreen
lblSTATUS.Caption = "Write OK!"

Exit Sub

Err:

lblSTATUS.ForeColor = vbRed
lblSTATUS.Caption = "I2C Error!"
shpBar.FillColor = vbRed

End Sub

Private Sub Command1_Click()
Dim I2cTrans As I2C_TRANS                     ' Dimension an I2C_TRANS structure
Dim lWritten As Long
Dim MyStep As Integer
Dim byteSend As Integer
Dim Add As Integer
Dim mDelay As Long
Dim splitData() As String

Me.lblSTATUS.Caption = ""
DoEvents

I2cTrans.wCount.hi = 0
I2cTrans.wCount.lo = 2

I2cTrans.byDevId = &HE8
I2cTrans.byType = I2C_TRANS_8ADR

I2cTrans.wMemAddr.hi = 0
I2cTrans.wMemAddr.lo = 0
            
I2cTrans.Data(0) = &H43
I2cTrans.Data(1) = &H77
           
lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
If lWritten <> 2 Then
    delay_ms 50
    lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
    If lWritten <> byteSend Then
        GoTo Err
    End If
End If

lblSTATUS.ForeColor = vbGreen
lblSTATUS.Caption = "Write OK!"

Exit Sub

Err:

lblSTATUS.ForeColor = vbRed
lblSTATUS.Caption = "I2C Error!"

End Sub

Private Sub Form_Load()

'If DAPI_ConfigIoPorts(hDevInstance, &HFFFFE) = 0 Then
'    MsgBox "Config IO Port Error!"
'End If

InitGrdData
BusSelected = NoBus

optSRQ(0).Value = False
optSRQ(1).Value = True
optMemorySize(0).Value = True
shpBar.Width = 0
cmbReadSize.List(0) = "16"
cmbReadSize.List(1) = "256"
cmbReadSize.Text = "256"

txtReadDelay.Text = "10"
txtWriteDelay.Text = "10"

optPA0(0).Value = True
optPA0(1).Value = False

'VScroll1.Max = 7
End Sub

Private Sub Form_Terminate()
AppExit
End Sub

Private Sub Form_Unload(Cancel As Integer)
AppExit
End Sub
Private Sub AppExit()

optSRQ_Click (1)

End Sub

Private Sub Frame7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Text1.SetFocus
End Sub

Private Sub grdData_DblClick()
Dim CurrentRow As Integer
Dim CurrentCol As Integer
Dim t As String
Dim ChkSum As Long
Dim Data_Before As String
Dim Data_After As String

If grdData.Rows < 2 Then Exit Sub

CurrentRow = grdData.Row
CurrentCol = grdData.Col

Data_Before = grdData.TextMatrix(CurrentRow, CurrentCol)
t = InputBox(Data_Before & " ==> ? (Hex Data)", "Data Edit")
If t = "" Then Exit Sub
If IsNumeric("&H" & t) = False Then MsgBox "Not a valid number!", vbExclamation: Exit Sub
t = UCase(t)
If Len(t) = 1 Then t = "0" & t
If Len(t) > 2 Then Exit Sub

Data_After = t
grdData.TextMatrix(CurrentRow, CurrentCol) = Data_After

If lblChkSum.Caption = "" Then Exit Sub
ChkSum = Int("&H" & lblChkSum.Caption) - Int("&H" & Data_Before) + Int("&H" & Data_After)
lblChkSum.Caption = Hex(ChkSum)

End Sub

Private Sub grdData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim CurrentRow As Integer
Dim CurrentCol As Integer

'If grdData.Rows < 3 Then Exit Sub


'CurrentRow = grdData.Row
'CurrentCol = grdData.Col


End Sub


Private Sub optBus_Click(Index As Integer)
Dim I2cTrans As I2C_TRANS                     ' Dimension an I2C_TRANS structure
Dim lWritten As Long                          ' Dimension a long to hold the returned value

'Status = tvif_iic_wrt(tvif_add, &HE0, &H4, 1, Chr(4), 10)  '''IIC0
'Status = tvif_iic_wrt(tvif_add, &HE0, &H5, 1, Chr(5), 10)  '''IIC1
If BusSelected = Index Then Exit Sub

lblBusSelect.Caption = ""
I2cTrans.byDevId = &HE0
I2cTrans.byType = I2C_TRANS_TYPE.I2C_TRANS_8ADR
I2cTrans.wCount.hi = 0                        ' only writing 1 byte, so set to 0
I2cTrans.wCount.lo = 1                        ' writing 1 byte, so set to 1

Select Case Index

Case 0

    I2cTrans.wMemAddr.hi = 0
    I2cTrans.wMemAddr.lo = &H4
    I2cTrans.Data(0) = &H4
  
    ' call the function
    lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
    If lWritten = 1 Then
        lblBusSelect.ForeColor = vbGreen
        lblBusSelect.Caption = "I2C 0 OK!"
        BusSelected = I2C0
    Else
        lblBusSelect.ForeColor = vbRed
        lblBusSelect.Caption = "I2C 0 NG!"
        BusSelected = NoBus
    End If
    
Case 1

    I2cTrans.wMemAddr.hi = 0
    I2cTrans.wMemAddr.lo = &H5
    I2cTrans.Data(0) = &H5
  
    ' call the function
    lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
    If lWritten = 1 Then
        lblBusSelect.ForeColor = vbGreen
        lblBusSelect.Caption = "I2C 1 OK!"
        BusSelected = I2C1
    Else
        lblBusSelect.ForeColor = vbRed
        lblBusSelect.Caption = "I2C 1 NG!"
        BusSelected = NoBus
    End If
    

End Select

delay_ms 50

End Sub

Private Sub optMemorySize_Click(Index As Integer)

SlvNo = 2 ^ Index

End Sub

Private Sub optPA0_Click(Index As Integer)
Dim portMask As Long
Dim pushData As Long

Select Case Index

Case 0  ''A

'''0x000CBBAA
portMask = &H1
pushData = &H0

If DAPI_WriteIoPorts(hDevInstance, pushData, portMask) = 0 Then
    MsgBox "Write to PA0 NG!"
End If


Case 1  ''B

'''0x000CBBAA
portMask = &H1
pushData = &H1

If DAPI_WriteIoPorts(hDevInstance, pushData, portMask) = 0 Then
    MsgBox "Write to PA0 NG!"
End If

End Select

End Sub

Private Sub optSRQ_Click(Index As Integer)
Dim portMask As Long
Dim pushData As Long

Select Case Index

Case 0  ''Low

'''0x000CBBAA
portMask = &H80000
pushData = &H80000

If DAPI_WriteIoPorts(hDevInstance, pushData, portMask) = 0 Then
    MsgBox "SRQ Low Error!"
End If


Case 1  ''High

'''0x000CBBAA
portMask = &H80000
pushData = &H0

If DAPI_WriteIoPorts(hDevInstance, pushData, portMask) = 0 Then
    MsgBox "SRQ High Error!"
End If

End Select

End Sub

Private Sub Timer1_Timer()
Me.txtBKSVslave.Text = Hex(Val("&H" & txtBKSVslave.Text) + 1)
cmdReadBKSV_Click
End Sub

Private Sub txtBiDirectional_Change(Index As Integer)

If IsNumeric("&H" & txtBiDirectional(Index).Text) = False And Len(txtBiDirectional(Index).Text) = 2 Then
    MsgBox "Not a valid number!", vbExclamation, "Bi-Directional Command"
        txtBiDirectional(Index).SetFocus
        txtBiDirectional(Index).SelStart = 0
        txtBiDirectional(Index).SelLength = 2
    Exit Sub
End If

End Sub

Private Sub txtReadDelay_Change()

If IsNumeric(txtReadDelay.Text) = False Then
    MsgBox "Invalid number!", vbExclamation, "Read Delay"
    txtReadDelay.Text = ""
End If

End Sub


Private Sub txtWriteDelay_Change()

If IsNumeric(txtWriteDelay.Text) = False Then
    MsgBox "Invalid number!", vbExclamation, "Write Delay"
    txtWriteDelay.Text = ""
End If

End Sub

Private Sub VScroll1_Change()
Dim mAdd As String
Dim hAdd As String
Dim mValue As String

txtSlave.Text = Hex(VScroll1.Value)

Exit Sub
If VScroll1.Value = 0 Then VScroll1.Value = 1

mValue = Hex((VScroll1.Value - 1) * 2)
If Len(mValue) = 1 Then

    txtSlave.Text = "A" & Right(mValue, 1)
    
Else

    Select Case Left(mValue, 1)
    Case 1  'B
        txtSlave.Text = "B" & Right(mValue, 1)
    Case 2  'C
        txtSlave.Text = "C" & Right(mValue, 1)
    Case 3  'D
        txtSlave.Text = "D" & Right(mValue, 1)
    Case 4  'E
        txtSlave.Text = "E" & Right(mValue, 1)
    Case 5  'F
        txtSlave.Text = "F" & Right(mValue, 1)
    Case 6  'G
        txtSlave.Text = "G" & Right(mValue, 1)
    Case 7  'H
        txtSlave.Text = "H" & Right(mValue, 1)
    End Select

End If

Exit Sub

hAdd = Hex(VScroll1.Value)
If Len(hAdd) = 1 Then hAdd = "0" & hAdd
mAdd = hAdd & "00"
For i = 1 To grdData.Rows - 1
    If grdData.TextMatrix(i, 0) = mAdd Then
        grdData.TopRow = i
        Exit Sub
    End If
Next i

End Sub

Private Sub OpenDevice()
Dim i As Byte

For i = 0 To 127 Step 1
    If (OpenDiHandle(i)) = 1 Then
      ' succesfully opened a device
      Exit For
    End If
Next i
  
End Sub
Private Function OpenDiHandle(byDevInstance As Byte) As Byte
  ' this subroutine will handle opening the specified device instance
  
  ' make sure no handle is currently open
  CloseDiHandle
  
  ' now attempt to open handle to device instance
  hDevInstance = DAPI_OpenDeviceInstance(sDevSymName, byDevInstance)
  
  ' test result of function call and flag success or failure
  If (hDevInstance <> INVALID_HANDLE_VALUE) Then
    OpenDiHandle = 1
  Else
    OpenDiHandle = 0
  End If

End Function
Private Sub CloseDiHandle()

If hDevInstance <> INVALID_HANDLE_VALUE Then
    If DAPI_CloseDeviceInstance(hDevInstance) Then
      ' everythings zen
    Else
      ' SNAFU
    End If
    hDevInstance = INVALID_HANDLE_VALUE
End If
  
End Sub


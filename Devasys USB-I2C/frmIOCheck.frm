VERSION 5.00
Begin VB.Form frmIOCheck 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IO Check"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
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
   ScaleHeight     =   7560
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox shpBorder 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3000
      ScaleHeight     =   195
      ScaleWidth      =   1500
      TabIndex        =   40
      Top             =   2160
      Width           =   1560
      Begin VB.Shape shpBar 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.Timer tmrPoll 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3000
      Top             =   2520
   End
   Begin VB.CommandButton cmdPoll 
      Caption         =   "Start Poll"
      Height          =   375
      Left            =   3000
      TabIndex        =   39
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000E&
      Caption         =   "Mode"
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
      Left            =   3000
      TabIndex        =   34
      Top             =   0
      Width           =   1575
      Begin VB.OptionButton optIO 
         BackColor       =   &H8000000E&
         Caption         =   "Config"
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optIO 
         BackColor       =   &H8000000E&
         Caption         =   "Write"
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optIO 
         BackColor       =   &H8000000E&
         Caption         =   "Read"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblPoll 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   44
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Height          =   1935
      Left            =   3000
      TabIndex        =   31
      Top             =   5520
      Width           =   1575
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdLoadConfig 
         Caption         =   "Load IO Config"
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSaveConfig 
         Caption         =   "Save IO Config"
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Legend"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   27
      Top             =   4680
      Width           =   2775
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Write to Output"
         Height          =   255
         Left            =   600
         TabIndex        =   42
         Top             =   1875
         Width           =   2055
      End
      Begin VB.Shape shpWrite 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Input Detect"
         Height          =   255
         Left            =   600
         TabIndex        =   41
         Top             =   1395
         Width           =   2055
      End
      Begin VB.Shape shpDetect 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Shape           =   3  'Circle
         Top             =   1320
         Width           =   375
      End
      Begin VB.Shape shpError 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Shape           =   3  'Circle
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Error"
         Height          =   255
         Left            =   600
         TabIndex        =   30
         Top             =   2355
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Port Set as Output"
         Height          =   255
         Left            =   600
         TabIndex        =   29
         Top             =   920
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Port Set as Input"
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   420
         Width           =   2055
      End
      Begin VB.Shape shpOutput 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Shape           =   3  'Circle
         Top             =   840
         Width           =   375
      End
      Begin VB.Shape shpInput 
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Shape           =   3  'Circle
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Ports"
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
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   1920
         ScaleHeight     =   3975
         ScaleWidth      =   615
         TabIndex        =   19
         Top             =   480
         Width           =   615
         Begin VB.Label lblShpC 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   4
            Left            =   135
            TabIndex        =   23
            Top             =   2100
            Width           =   375
         End
         Begin VB.Label lblShpC 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   5
            Left            =   135
            TabIndex        =   22
            Top             =   2580
            Width           =   375
         End
         Begin VB.Label lblShpC 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   6
            Left            =   135
            TabIndex        =   21
            Top             =   3060
            Width           =   375
         End
         Begin VB.Label lblShpC 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   7
            Left            =   135
            TabIndex        =   20
            Top             =   3540
            Width           =   375
         End
         Begin VB.Shape shpC 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   4
            Left            =   120
            Shape           =   3  'Circle
            Top             =   2040
            Width           =   375
         End
         Begin VB.Shape shpC 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   5
            Left            =   120
            Shape           =   3  'Circle
            Top             =   2520
            Width           =   375
         End
         Begin VB.Shape shpC 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   6
            Left            =   120
            Shape           =   3  'Circle
            Top             =   3000
            Width           =   375
         End
         Begin VB.Shape shpC 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   7
            Left            =   120
            Shape           =   3  'Circle
            Top             =   3480
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   1080
         ScaleHeight     =   3975
         ScaleWidth      =   615
         TabIndex        =   10
         Top             =   480
         Width           =   615
         Begin VB.Label lblShpB 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   0
            Left            =   135
            TabIndex        =   18
            Top             =   180
            Width           =   375
         End
         Begin VB.Label lblShpB 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   17
            Top             =   660
            Width           =   375
         End
         Begin VB.Label lblShpB 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   2
            Left            =   135
            TabIndex        =   16
            Top             =   1140
            Width           =   375
         End
         Begin VB.Label lblShpB 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   135
            TabIndex        =   15
            Top             =   1620
            Width           =   375
         End
         Begin VB.Label lblShpB 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   4
            Left            =   135
            TabIndex        =   14
            Top             =   2100
            Width           =   375
         End
         Begin VB.Label lblShpB 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   5
            Left            =   135
            TabIndex        =   13
            Top             =   2580
            Width           =   375
         End
         Begin VB.Label lblShpB 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   6
            Left            =   135
            TabIndex        =   12
            Top             =   3060
            Width           =   375
         End
         Begin VB.Shape shpB 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   0
            Left            =   120
            Shape           =   3  'Circle
            Top             =   120
            Width           =   375
         End
         Begin VB.Shape shpB 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   1
            Left            =   120
            Shape           =   3  'Circle
            Top             =   600
            Width           =   375
         End
         Begin VB.Shape shpB 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   2
            Left            =   120
            Shape           =   3  'Circle
            Top             =   1080
            Width           =   375
         End
         Begin VB.Shape shpB 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   3
            Left            =   120
            Shape           =   3  'Circle
            Top             =   1560
            Width           =   375
         End
         Begin VB.Shape shpB 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   4
            Left            =   120
            Shape           =   3  'Circle
            Top             =   2040
            Width           =   375
         End
         Begin VB.Shape shpB 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   5
            Left            =   120
            Shape           =   3  'Circle
            Top             =   2520
            Width           =   375
         End
         Begin VB.Shape shpB 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   6
            Left            =   120
            Shape           =   3  'Circle
            Top             =   3000
            Width           =   375
         End
         Begin VB.Label lblShpB 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   7
            Left            =   135
            TabIndex        =   11
            Top             =   3540
            Width           =   375
         End
         Begin VB.Shape shpB 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   7
            Left            =   120
            Shape           =   3  'Circle
            Top             =   3480
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3975
         Left            =   240
         ScaleHeight     =   3975
         ScaleWidth      =   615
         TabIndex        =   1
         Top             =   480
         Width           =   615
         Begin VB.Label lblShpA 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   7
            Left            =   135
            TabIndex        =   9
            Top             =   3540
            Width           =   375
         End
         Begin VB.Shape shpA 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   7
            Left            =   120
            Shape           =   3  'Circle
            Top             =   3480
            Width           =   375
         End
         Begin VB.Label lblShpA 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   6
            Left            =   135
            TabIndex        =   8
            Top             =   3060
            Width           =   375
         End
         Begin VB.Shape shpA 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   6
            Left            =   120
            Shape           =   3  'Circle
            Top             =   3000
            Width           =   375
         End
         Begin VB.Label lblShpA 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   5
            Left            =   135
            TabIndex        =   7
            Top             =   2580
            Width           =   375
         End
         Begin VB.Shape shpA 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   5
            Left            =   120
            Shape           =   3  'Circle
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label lblShpA 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   4
            Left            =   135
            TabIndex        =   6
            Top             =   2100
            Width           =   375
         End
         Begin VB.Shape shpA 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   4
            Left            =   120
            Shape           =   3  'Circle
            Top             =   2040
            Width           =   375
         End
         Begin VB.Label lblShpA 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   135
            TabIndex        =   5
            Top             =   1620
            Width           =   375
         End
         Begin VB.Shape shpA 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   3
            Left            =   120
            Shape           =   3  'Circle
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label lblShpA 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   2
            Left            =   135
            TabIndex        =   4
            Top             =   1140
            Width           =   375
         End
         Begin VB.Shape shpA 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   2
            Left            =   120
            Shape           =   3  'Circle
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label lblShpA 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   1
            Left            =   135
            TabIndex        =   3
            Top             =   660
            Width           =   375
         End
         Begin VB.Shape shpA 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   1
            Left            =   120
            Shape           =   3  'Circle
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblShpA 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "A0"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   0
            Left            =   135
            TabIndex        =   2
            Top             =   180
            Width           =   375
         End
         Begin VB.Shape shpA 
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   0
            Left            =   120
            Shape           =   3  'Circle
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.Label lblGetPort 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   43
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "C"
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
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "B"
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
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "A"
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
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmIOCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myPoll As Boolean
Dim barCount As Integer

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPoll_Click()

If cmdPoll.Caption = "Start Poll" Then
    cmdPoll.Caption = "End Poll"
    myPoll = True
    tmrPoll.Enabled = True
    optIO(2).Enabled = False
Else
    cmdPoll.Caption = "Start Poll"
    myPoll = False
    tmrPoll.Enabled = False
    shpBar.Width = 0
    optIO(2).Enabled = True
    
    'For i = 0 To 7
    '    If shpA(i).FillColor = vbYellow Then shpA(i).FillColor = vbBlue: lblShpA(i).ForeColor = vbWhite
    '    If shpB(i).FillColor = vbYellow Then shpB(i).FillColor = vbBlue: lblShpB(i).ForeColor = vbWhite
    '    If i >= 4 Then If shpC(i).FillColor = vbYellow Then shpC(i).FillColor = vbBlue:: lblShpC(i).ForeColor = vbWhite
    'Next i
    
    GetPort
    lblPoll.Caption = ""
End If

End Sub

Private Sub Form_Load()

For i = 0 To 7
    lblShpA(i).Caption = "A" & i
    lblShpB(i).Caption = "B" & i
    If i >= 4 Then lblShpC(i).Caption = "C" & i
Next i

shpInput.FillColor = vbGreen
shpOutput.FillColor = vbBlue
shpError.FillColor = vbRed
shpDetect.FillColor = vbWhite
shpWrite.FillColor = vbYellow

optIO(2).Value = True
cmdPoll.Enabled = False
myPoll = False
shpBar.Width = 0
shpBar.FillColor = vbBlue
barCount = PLUS

GetPort
'''SetPort '''Set all as input

End Sub
Private Sub GetPort()
Dim readIOConfig As Long
Dim readByte As String
Dim rPA, rPB, rPC As Integer
Dim tmpBitA, tmpBitB, tmpBitC As Integer

If DAPI_GetIoConfig(hDevInstance, readIOConfig) = 0 Then
    For i = 0 To 7
        shpA(i).FillColor = vbRed
        shpB(i).FillColor = vbRed
        If i >= 4 Then shpC(i).FillColor = vbRed
    Next i
    Exit Sub
Else
    readByte = Hex(readIOConfig)
    If Len(readByte) >= 5 Then readByte = Right(readByte, 5)
    Do While Len(readByte) < 5
        readByte = "0" & readByte
    Loop
    rPA = "&H" & Mid(readByte, 4, 2)
    rPB = "&H" & Mid(readByte, 2, 2)
    rPC = "&H" & Mid(readByte, 1, 1)
    
    lblGetPort.Caption = readByte
    
    For i = 0 To 7
    
        tmpBitA = rPA Mod 2
        If tmpBitA = 0 Then shpA(i).FillColor = vbBlue: lblShpA(i).ForeColor = vbWhite
        If tmpBitA = 1 Then shpA(i).FillColor = vbGreen: lblShpA(i).ForeColor = vbBlack
        rPA = rPA \ 2
    
        tmpBitB = rPB Mod 2
        If tmpBitB = 0 Then shpB(i).FillColor = vbBlue: lblShpB(i).ForeColor = vbWhite
        If tmpBitB = 1 Then shpB(i).FillColor = vbGreen: lblShpB(i).ForeColor = vbBlack
        rPB = rPB \ 2
    
        If i >= 4 Then
            tmpBitC = rPC Mod 2
            If tmpBitC = 0 Then shpC(i).FillColor = vbBlue: lblShpC(i).ForeColor = vbWhite
            If tmpBitC = 1 Then shpC(i).FillColor = vbGreen: lblShpC(i).ForeColor = vbBlack
            rPC = rPC \ 2
        End If
    
    Next i
End If

End Sub


Private Sub SetPort()
  ' Configures the IoPorts
  ' Io Port bit mapping is 0x000CBBAA
  ' individual bits are configured by writing 0 = output, 1 = input
  
  ' configure IO Ports for C = input, B = output A = input
  ' 0000 0000 0000 CCCC BBBB BBBB AAAA AAAA
  ' 0    0    0    1111 0000 0000 1111 1111
  ' 0    0    0    F    0    0    F    F
  
  '''Configure all Port as input
  ' 0000 0000 0000 1111 1111 1111 1111 1111
  ' 0    0    0    F    F    F    F    F

  If DAPI_ConfigIoPorts(hDevInstance, &HFF00) = 1 Then
    For i = 0 To 7
        shpA(i).FillColor = vbBlue
        shpB(i).FillColor = vbGreen
        lblShpB(i).ForeColor = vbBlack
        If i >= 4 Then shpC(i).FillColor = vbBlue
    Next i
  Else
    For i = 0 To 7
        shpA(i).FillColor = vbRed
        shpB(i).FillColor = vbRed
        If i >= 4 Then shpC(i).FillColor = vbRed
    Next i
  End If


End Sub


Private Sub lblShpA_DblClick(Index As Integer)
Dim readIOConfig As Long
Dim REreadIOconfig As Long
Dim i, j As Long

If optIO(0).Value = True Then Exit Sub
If optIO(1).Value = True And myPoll = True Then
    If shpA(Index).FillColor = vbBlue Then shpA(Index).FillColor = vbYellow: lblShpA(Index).ForeColor = vbBlack: Exit Sub
    If shpA(Index).FillColor = vbYellow Then shpA(Index).FillColor = vbBlue: lblShpA(Index).ForeColor = vbWhite: Exit Sub
ElseIf optIO(1).Value = True And myPoll = False Then
    Exit Sub
End If
If shpA(Index).FillColor = vbRed Then Exit Sub
If myPoll = True And shpA(Index).FillColor = vbGreen Then Exit Sub

lblGetPort.Caption = ""

If DAPI_GetIoConfig(hDevInstance, readIOConfig) = 0 Then Exit Sub

t = Hex(readIOConfig)

i = 2 ^ Index
If shpA(Index).FillColor = vbGreen Then
    j = readIOConfig - i
ElseIf shpA(Index).FillColor = vbBlue Then
    j = readIOConfig + i
End If

Debug.Print Hex(j)
If DAPI_ConfigIoPorts(hDevInstance, j) = 0 Then Exit Sub

If DAPI_GetIoConfig(hDevInstance, REreadIOconfig) = 0 Then Exit Sub

If REreadIOconfig = j Then
    If shpA(Index).FillColor = vbGreen Then
        shpA(Index).FillColor = vbBlue
        lblShpA(Index).ForeColor = vbWhite
    ElseIf shpA(Index).FillColor = vbBlue Then
        shpA(Index).FillColor = vbGreen
        lblShpA(Index).ForeColor = vbBlack
    End If
End If

lblGetPort.Caption = Right(Hex(REreadIOconfig), 5)

End Sub

Private Sub lblShpB_DblClick(Index As Integer)

Dim readIOConfig As Long
Dim REreadIOconfig As Long
Dim i, j As Long

If optIO(0).Value = True Then Exit Sub
If optIO(1).Value = True And myPoll = True Then
    If shpB(Index).FillColor = vbBlue Then shpB(Index).FillColor = vbYellow: lblShpB(Index).ForeColor = vbBlack: Exit Sub
    If shpB(Index).FillColor = vbYellow Then shpB(Index).FillColor = vbBlue: lblShpB(Index).ForeColor = vbWhite: Exit Sub
ElseIf optIO(1).Value = True And myPoll = False Then
    Exit Sub
End If
If shpB(Index).FillColor = vbRed Then Exit Sub
If myPoll = True And shpB(Index).FillColor = vbGreen Then Exit Sub

lblGetPort.Caption = ""

If DAPI_GetIoConfig(hDevInstance, readIOConfig) = 0 Then Exit Sub

i = 2 ^ (Index + 8)
If shpB(Index).FillColor = vbGreen Then
    j = readIOConfig - i
ElseIf shpB(Index).FillColor = vbBlue Then
    j = readIOConfig + i
End If

Debug.Print Hex(j)
If DAPI_ConfigIoPorts(hDevInstance, j) = 0 Then Exit Sub

If DAPI_GetIoConfig(hDevInstance, REreadIOconfig) = 0 Then Exit Sub

If REreadIOconfig = j Then
    If shpB(Index).FillColor = vbGreen Then
        shpB(Index).FillColor = vbBlue
        lblShpB(Index).ForeColor = vbWhite
    ElseIf shpB(Index).FillColor = vbBlue Then
        shpB(Index).FillColor = vbGreen
        lblShpB(Index).ForeColor = vbBlack
    End If
End If

lblGetPort.Caption = Right(Hex(REreadIOconfig), 5)

End Sub

Private Sub lblShpC_DblClick(Index As Integer)

Dim readIOConfig As Long
Dim REreadIOconfig As Long
Dim i, j As Long

If optIO(0).Value = True Then Exit Sub
If optIO(1).Value = True And myPoll = True Then
    If shpC(Index).FillColor = vbBlue Then shpC(Index).FillColor = vbYellow: lblShpC(Index).ForeColor = vbBlack: Exit Sub
    If shpC(Index).FillColor = vbYellow Then shpC(Index).FillColor = vbBlue: lblShpC(Index).ForeColor = vbWhite: Exit Sub
ElseIf optIO(1).Value = True And myPoll = False Then
    Exit Sub
End If
If shpC(Index).FillColor = vbRed Then Exit Sub
If myPoll = True And shpC(Index).FillColor = vbGreen Then Exit Sub

lblGetPort.Caption = ""

If DAPI_GetIoConfig(hDevInstance, readIOConfig) = 0 Then Exit Sub

i = 2 ^ (Index + 16 - 4)
If shpC(Index).FillColor = vbGreen Then
    j = readIOConfig - i
ElseIf shpC(Index).FillColor = vbBlue Then
    j = readIOConfig + i
End If

Debug.Print Hex(j)
If DAPI_ConfigIoPorts(hDevInstance, j) = 0 Then Exit Sub

If DAPI_GetIoConfig(hDevInstance, REreadIOconfig) = 0 Then Exit Sub

If REreadIOconfig = j Then
    If shpC(Index).FillColor = vbGreen Then
        shpC(Index).FillColor = vbBlue
        lblShpC(Index).ForeColor = vbWhite
    ElseIf shpC(Index).FillColor = vbBlue Then
        shpC(Index).FillColor = vbGreen
        lblShpC(Index).ForeColor = vbBlack
    End If
End If

lblGetPort.Caption = Right(Hex(REreadIOconfig), 5)

End Sub

Private Sub optIO_Click(Index As Integer)

If Index = 2 Then cmdPoll.Enabled = False Else cmdPoll.Enabled = True

End Sub

Private Sub tmrPoll_Timer()
Dim t As Long
Dim barSTEP As Integer
Dim pullData As Long
Dim pushData As Long
Dim portMask As Long
Dim splitData(2) As Integer
Dim tA, tB, tC As Integer

barSTEP = 200

If barCount = PLUS Then
    shpBar.Width = shpBar.Width + barSTEP
    If shpBar.Width >= shpBorder.Width Then barCount = MINUS
ElseIf barCount = MINUS Then
    t = shpBar.Width
    If (t - barSTEP) <= 0 Then
        barCount = PLUS
    Else
        shpBar.Width = shpBar.Width - barSTEP
    End If
End If

If optIO(0).Value = True Then   '''Read
    If DAPI_ReadIoPorts(hDevInstance, pullData) = 1 Then
        lblPoll.Caption = Hex("&H" & Hex(pullData) And &HFFFFF)
        'If lblpoll.caption <> "FFFFF" Then
            splitData(0) = "&H" & Mid(lblPoll.Caption, 4, 2)
            tA = splitData(0)
            For i = 0 To 7
                If shpA(i).FillColor <> vbBlue Then
                    If (tA Mod 2) = 1 Then
                        shpA(i).FillColor = vbGreen
                    ElseIf (tA Mod 2) = 0 Then
                        shpA(i).FillColor = vbWhite
                    End If
                End If
                tA = tA \ 2
            Next i
            splitData(1) = "&H" & Mid(lblPoll.Caption, 2, 2)
            tB = splitData(1)
            For i = 0 To 7
                If shpB(i).FillColor <> vbBlue Then
                    If (tB Mod 2) = 1 Then
                        shpB(i).FillColor = vbGreen
                    ElseIf (tB Mod 2) = 0 Then
                        shpB(i).FillColor = vbWhite
                    End If
                    
                End If
                tB = tB \ 2
            Next i
            splitData(2) = "&H" & Mid(lblPoll.Caption, 1, 1)
            tC = splitData(2)
            For i = 4 To 7
                If shpC(i).FillColor <> vbBlue Then
                    If (tC Mod 2) = 1 Then
                        shpC(i).FillColor = vbGreen
                    ElseIf (tC Mod 2) = 0 Then
                        shpC(i).FillColor = vbWhite
                    End If
                End If
                tC = tC \ 2
            Next i
        'End If
    
    End If
ElseIf optIO(1).Value = True Then   '''Write
    portMask = "&HFFFFFFFF" ''& lblGetPort.Caption
    pushData = 0
    For i = 0 To 7
        If shpA(i).FillColor = vbGreen Or shpA(i).FillColor = vbYellow Then pushData = pushData + (2 ^ i)
        If shpB(i).FillColor = vbGreen Or shpA(i).FillColor = vbYellow Then pushData = pushData + (2 ^ (i + 8))
        If i >= 4 Then
            If shpC(i).FillColor = vbGreen Or shpC(i).FillColor = vbYellow Then pushData = pushData + (2 ^ (i + 16 - 4))
        End If
    Next i
    Debug.Print Hex(pushData)
    If DAPI_WriteIoPorts(hDevInstance, pushData, portMask) = 1 Then
    End If
End If


End Sub

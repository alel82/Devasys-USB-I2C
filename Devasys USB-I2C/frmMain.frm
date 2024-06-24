VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7l.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000E&
      Height          =   2055
      Left            =   120
      TabIndex        =   54
      Top             =   0
      Width           =   5055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000014&
      Height          =   2055
      Left            =   5280
      TabIndex        =   51
      Top             =   0
      Width           =   4455
      Begin VB.PictureBox picSTATUS 
         Height          =   1575
         Left            =   120
         ScaleHeight     =   1515
         ScaleWidth      =   4155
         TabIndex        =   52
         Top             =   300
         Width           =   4215
         Begin VB.Label lblSTATUS 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PASS"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   72
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   1575
            Left            =   240
            TabIndex        =   53
            Top             =   -120
            Width           =   3615
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000E&
      Caption         =   "Delay Settings"
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
      Height          =   3375
      Left            =   9840
      TabIndex        =   26
      Top             =   720
      Width           =   2055
      Begin VB.TextBox txtPROCESS 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         MaxLength       =   5
         TabIndex        =   48
         Text            =   "200"
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox txtEND 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         MaxLength       =   5
         TabIndex        =   45
         Text            =   "0"
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox txtSTART 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         MaxLength       =   5
         TabIndex        =   42
         Text            =   "1000"
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtVERIFY 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         MaxLength       =   5
         TabIndex        =   40
         Text            =   "50"
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox txtWRITE 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         MaxLength       =   5
         TabIndex        =   37
         Text            =   "50"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtBUS 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         MaxLength       =   5
         TabIndex        =   34
         Text            =   "500"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtSRQHI 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         MaxLength       =   5
         TabIndex        =   31
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtSRQLO 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         MaxLength       =   5
         TabIndex        =   28
         Text            =   "2000"
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label25 
         BackColor       =   &H8000000E&
         Caption         =   "ms"
         Height          =   255
         Left            =   1680
         TabIndex        =   50
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label24 
         BackColor       =   &H8000000E&
         Caption         =   "PROCESS"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label23 
         BackColor       =   &H8000000E&
         Caption         =   "END"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label22 
         BackColor       =   &H8000000E&
         Caption         =   "ms"
         Height          =   255
         Left            =   1680
         TabIndex        =   46
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label Label21 
         BackColor       =   &H8000000E&
         Caption         =   "START"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000E&
         Caption         =   "ms"
         Height          =   255
         Left            =   1680
         TabIndex        =   43
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000E&
         Caption         =   "ms"
         Height          =   255
         Left            =   1680
         TabIndex        =   41
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000E&
         Caption         =   "VERIFY"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000E&
         Caption         =   "ms"
         Height          =   255
         Left            =   1680
         TabIndex        =   38
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000E&
         Caption         =   "WRITE"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000E&
         Caption         =   "ms"
         Height          =   255
         Left            =   1680
         TabIndex        =   35
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000E&
         Caption         =   "BUS"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000E&
         Caption         =   "ms"
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000E&
         Caption         =   "SRQ HI"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000E&
         Caption         =   "ms"
         Height          =   255
         Left            =   1680
         TabIndex        =   29
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         Caption         =   "SRQ LO"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "Process Settings"
      Height          =   495
      Left            =   9840
      TabIndex        =   25
      Top             =   120
      Width           =   2055
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   5300
      Left            =   5280
      TabIndex        =   23
      Top             =   2040
      Width           =   4455
      Begin VB.ListBox List1 
         Height          =   4860
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.PictureBox shpBar 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   11700
      TabIndex        =   15
      Top             =   8880
      Width           =   11760
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   11775
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H8000000E&
         Height          =   1095
         Left            =   10680
         Picture         =   "frmMain.frx":4A9A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Load File"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H8000000E&
         Default         =   -1  'True
         Height          =   1095
         Left            =   120
         Picture         =   "frmMain.frx":4F1F
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Load File"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "INFO :"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   59
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "WRITE:A4-aa-bb-cc"
         Height          =   255
         Left            =   2400
         TabIndex        =   58
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "WP"
         Height          =   255
         Left            =   2400
         TabIndex        =   57
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "BUS:IIC0/IIC1"
         Height          =   255
         Left            =   2400
         TabIndex        =   56
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "SRQ:LOW/HiGH"
         Height          =   255
         Left            =   2400
         TabIndex        =   55
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Process"
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
      Height          =   5295
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   5085
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   4905
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   4845
         _cx             =   8546
         _cy             =   8652
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Index           =   1
      Left            =   9840
      TabIndex        =   6
      Top             =   4080
      Width           =   2055
      Begin VB.CommandButton cmdProp 
         Caption         =   "&USB Properties"
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check"
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
         Height          =   1215
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   1575
         Begin VB.CommandButton cmdI2CCheck 
            Caption         =   "I2C &Check"
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1320
         End
         Begin VB.CommandButton cmdIO 
            Caption         =   "&IO Check"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Left            =   900
            TabIndex        =   22
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdModelChange 
         Caption         =   "&Model Change"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "USB I2C"
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
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   150
         Width           =   1575
         Begin VB.CommandButton Command1 
            Caption         =   ".."
            Height          =   255
            Left            =   1200
            TabIndex        =   18
            Top             =   960
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Timer tmrDetect 
            Interval        =   100
            Left            =   0
            Top             =   480
         End
         Begin VB.Label lblA2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Left            =   900
            TabIndex        =   17
            Top             =   600
            Width           =   495
         End
         Begin VB.Shape shpA2 
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   900
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000E&
            Caption         =   "A2 Check"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblSerialID 
            Alignment       =   2  'Center
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
            TabIndex        =   14
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label lblDetect 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
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
            Left            =   900
            TabIndex        =   13
            Top             =   240
            Width           =   495
         End
         Begin VB.Shape shpDetect 
            FillStyle       =   0  'Solid
            Height          =   255
            Left            =   900
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000E&
            Caption         =   "Detect"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000009&
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
         Height          =   1095
         Index           =   2
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   1515
         Begin VB.PictureBox Picture1 
            Height          =   735
            Left            =   150
            ScaleHeight     =   675
            ScaleWidth      =   1155
            TabIndex        =   8
            Top             =   240
            Width           =   1215
            Begin VB.Label lblMemorySize 
               Alignment       =   2  'Center
               Caption         =   "128 K"
               BeginProperty Font 
                  Name            =   "Trebuchet MS"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   0
               TabIndex        =   9
               Top             =   120
               Width           =   1095
            End
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuTroubleshoot 
      Caption         =   "&Troubleshoot"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bDevicePresent As Boolean     ' flag to indicate presence of device
Dim SRQLOactive As Boolean
Dim WPactive As Boolean

Private Sub cmdExit_Click()
AppExit
End
End Sub

Private Sub cmdI2CCheck_Click()
frmI2CCheck.Show vbModal
End Sub

Private Sub cmdIO_Click()
If shpDetect.FillColor = vbRed Then Exit Sub
frmIOCheck.Show vbModal
End Sub

Private Sub cmdLoadFile_Click()
Dim textline As String
Dim tmpstr As String

CommonDialog1.FileName = ""
CommonDialog1.Filter = "Data Files (*.dat,*.ver)|*.dat;*.ver|All Files (*.*)|*.*"
CommonDialog1.DialogTitle = "Load Data from File"
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then Exit Sub

lblMemorySize.Caption = ""
Me.MousePointer = vbHourglass
DoEvents

InitGrdData

Open CommonDialog1.FileName For Input As #1

Do While Not EOF(1) ' Loop until end of file.
    Line Input #1, textline ' Read line into variable.
    tmpstr = Mid(textline, 7)
    If Len(tmpstr) > 40 Then
        grdData.Rows = grdData.Rows + 1
        grdData.TextMatrix(grdData.Rows - 1, 0) = Mid(textline, 2, 4)
        j = 1
        For i = 1 To 48 Step 3
            grdData.TextMatrix(grdData.Rows - 1, j) = Mid(tmpstr, i, 2)
            j = j + 1
        Next i
    End If
Loop

Close #1

Me.MousePointer = vbNormal
lblMemorySize.Caption = (grdData.Rows - 1) / 8 & " K"

End Sub

Private Sub cmdModelChange_Click()
frmModelChange.Show vbModal
End Sub

Private Sub cmdProcess_Click()
frmProcess.Show vbModal
End Sub

Private Sub cmdProp_Click()
If shpDetect.FillColor = vbRed Then Exit Sub
frmUSBProp.Show vbModal
A2Detect
End Sub

Private Sub ShowGO()

lblSTATUS.Caption = "PASS"
picSTATUS.BackColor = vbGreen

End Sub
Private Sub ShowNG()

lblSTATUS.Caption = "FAIL"
picSTATUS.BackColor = vbRed

End Sub

Private Sub cmdStart_Click()
Dim i As Integer
Dim t As Integer
Dim ProcessItem As String
Dim processDATA As String
Dim tacttime As Single
Dim processtime As Single

tacttime = Timer

picSTATUS.BackColor = vbCyan
lblSTATUS.Caption = ""
ButtonEnable False
InitGrdData
List1.Clear
List1.AddItem "[START]"
DoEvents

WPactive = False

delay_ms Val(txtSTART.Text)

For i = 1 To grdData.Rows - 1
    processtime = Timer
    grdData.Row = i
    grdData.Col = 1
    grdData.CellBackColor = vbCyan
    grdData.Col = 0
    DoEvents
    
    
    If InStr(1, grdData.TextMatrix(i, 1), ":") > 0 Then
        ProcessItem = Mid(grdData.TextMatrix(i, 1), 1, InStr(1, grdData.TextMatrix(i, 1), ":") - 1)
        processDATA = Mid(grdData.TextMatrix(i, 1), InStr(1, grdData.TextMatrix(i, 1), ":") + 1)
    Else
        ProcessItem = grdData.TextMatrix(i, 1)
        processDATA = ""
    End If
    'Debug.Print ProcessItem & " " & processDATA
    
    Select Case ProcessItem
        Case "DELAY"
            delay_ms processDATA
            
        Case "SRQ"
            If processDATA = "LOW" Then t = SRQ(SrqLow)
            If processDATA = "HIGH" Then t = SRQ(SrqHigh)
            
            grdData.Col = 2
            If t = OK Then
                grdData.TextMatrix(i, 2) = "OK"
                grdData.CellBackColor = vbGreen
            Else
                grdData.TextMatrix(i, 2) = "NG"
                grdData.CellBackColor = vbRed
                GoTo ProcessErr
            End If
            DoEvents
            
        Case "BUS"
            If processDATA = "IIC0" Then t = BUS(I2C0)
            If processDATA = "IIC1" Then t = BUS(I2C1)
            
            grdData.Col = 2
            If t = OK Then
                grdData.TextMatrix(i, 2) = "OK"
                grdData.CellBackColor = vbGreen
            Else
                grdData.TextMatrix(i, 2) = "NG"
                grdData.CellBackColor = vbRed
                GoTo ProcessErr
            End If
            DoEvents
        
        Case "WP"
            If WPactive = False Then
                t = WriteProtect
                grdData.Col = 2
                If t = OK Then
                    grdData.TextMatrix(i, 2) = "OK"
                    grdData.CellBackColor = vbGreen
                Else
                    grdData.TextMatrix(i, 2) = "NG"
                    grdData.CellBackColor = vbRed
                    GoTo ProcessErr
                End If
                DoEvents
            End If
        
        Case "WRITE"
            t = WriteData(processDATA)
            If t <> OK Then
                grdData.Col = 2
                grdData.TextMatrix(i, 2) = "NG"
                grdData.CellBackColor = vbRed
                GoTo ProcessErr
            End If
            
            grdData.Col = 2
            If t = OK Then
                grdData.TextMatrix(i, 2) = "OK"
                grdData.CellBackColor = vbGreen
            Else
                grdData.TextMatrix(i, 2) = "NG"
                grdData.CellBackColor = vbRed
                GoTo ProcessErr
            End If
            DoEvents
            
        Case "VERIFY"
    
    End Select
    
    delay_ms Val(txtPROCESS.Text)
    grdData.Col = 1
    grdData.CellBackColor = vbGray
    grdData.TextMatrix(i, 3) = Format(Timer - processtime, "0.00") & " s"
    grdData.TopRow = i - 1
    DoEvents
Next i

delay_ms Val(txtEND.Text)

grdData.Row = 0
grdData.Col = 0

List1.AddItem ""
List1.AddItem "Total Time : " & Format(Timer - tacttime, "0.00") & " s"
List1.AddItem "[END]"
List1.TopIndex = List1.ListCount - 1

ButtonEnable True
ShowGO

Exit Sub

ProcessErr:

If SRQLOactive = True And t <> IICError Then SRQ SrqHigh
If t = IICError Then ResetUSB

grdData.Row = 0
grdData.Col = 0

ButtonEnable True
ShowNG

List1.AddItem ""
List1.AddItem "Total Time : " & Format(Timer - tacttime, "0.00") & " s"
List1.AddItem "[END]"

End Sub

Private Sub Command1_Click()
Dim SerialID As String

If DAPI_GetSerialId(hDevInstance, SerialID) Then
    lblSerialID.Caption = SerialID
End If

End Sub

Private Sub Form_Load()

lblSTATUS.Caption = ""
picSTATUS.BackColor = vbGray

Me.Caption = " " & App.Title & "   [ver " & App.Major & "." & App.Minor & "." & App.Revision & "]"

sDevSymName = "UsbI2cIo"                        ' UsbI2cIo device symbolic name
byDevInstance = 255                             ' initial device instance (255 = no device)
hDevInstance = INVALID_HANDLE_VALUE             ' initialize file handle

ReadProcess
InitGrdData

Show
DoEvents

OpenDevice
cmdStart.SetFocus

End Sub
Private Sub ReadProcess()
Dim i As Integer

lpFileName$ = App.Path + "\Current.ini"

lpApplicationName$ = "Process Settings"
lpDefault = 0
nSize = 256

lpReturnedString$ = Space$(nSize)
lpKeyName$ = "Count"
n% = GetPrivateProfileString%(lpApplicationName$, lpKeyName$, lpDefault, lpReturnedString$, nSize, lpFileName$)
lpReturnedString$ = Left$(lpReturnedString$, InStr(lpReturnedString$, Chr$(0)) - 1)
ProcessCount = lpReturnedString$

If ProcessCount > 0 Then
    ReDim ProcessItem(ProcessCount - 1)
    
    For i = 0 To ProcessCount - 1
        lpReturnedString$ = Space$(nSize)
        lpKeyName$ = "Process" & i + 1
        n% = GetPrivateProfileString%(lpApplicationName$, lpKeyName$, lpDefault, lpReturnedString$, nSize, lpFileName$)
        lpReturnedString$ = Left$(lpReturnedString$, InStr(lpReturnedString$, Chr$(0)) - 1)
        ProcessItem(i) = lpReturnedString$
    Next i
End If

End Sub
Private Sub A2Detect()
Dim I2CProp As Byte
Dim bitCheck As Integer
Dim tmpByte As Integer

If DAPI_GetProperty(hDevInstance, I2CProp, &H1) = False Then
    MsgBox "A2 Bit NG!", vbExclamation, Me.Caption
    shpA2.FillColor = vbRed
    Exit Sub
Else
    shpA2.FillColor = vbRed
    tmpByte = "&H" & Hex(I2CProp)
    For i = 0 To 7
        bitCheck = tmpByte Mod 2
        tmpByte = tmpByte \ 2
        If bitCheck = 0 And i = 5 Then
            shpA2.FillColor = vbGreen
            Exit For
        End If
    Next i
    lblA2.Caption = Hex(I2CProp)
    If Len(lblA2.Caption) = 1 Then lblA2.Caption = "0" & lblA2.Caption
End If

End Sub
Private Sub OpenDevice()
Dim i As Byte

For i = 0 To 127 Step 1
    If (OpenDiHandle(i)) = 1 Then
      ' succesfully opened a device
      Exit For
    End If
Next i
  
' indicate success or notify of failure
If (i = 128) Then
    ' no device was found
    'Call MsgBox("No UUSB-I2CIO Devices were detected", vbOKOnly, "Device Open Error")
    shpDetect.FillColor = vbRed
    lblDetect.Caption = "NG"
Else
    ' found a device
    byDevInstance = i
    tmrDetect.Enabled = True
    shpDetect.FillColor = vbGreen
    lblDetect.Caption = "OK"
End If

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
    A2Detect
  Else
    OpenDiHandle = 0
    lblA2.Caption = ""
    shpA2.FillColor = vbRed
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
Public Function CheckDevice() As Boolean

    If hDevInstance = INVALID_HANDLE_VALUE Then
        If OpenDiHandle(byDevInstance) = 1 Then
            CheckDevice = True
        Else
            CheckDevice = False
        End If
    ElseIf DAPI_DetectDevice(hDevInstance) Then
        CheckDevice = True
    Else
        CloseDiHandle
        CheckDevice = False
    End If
  
End Function

Private Sub Form_Terminate()
AppExit
End Sub

Private Sub Form_Unload(Cancel As Integer)
AppExit
End Sub

Private Sub lblA2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If shpA2.FillColor = vbGreen Then
    lblA2.ToolTipText = "A2 bit OK!"
Else
    If lblA2.Caption <> "" Then lblA2.ToolTipText = "A2 bit NG!" Else lblA2.ToolTipText = "Devasys not found!"
End If

End Sub

Private Sub lblDetect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If shpDetect.FillColor = vbGreen Then
    lblDetect.ToolTipText = "Devasys detected!"
Else
    lblDetect.ToolTipText = "Devasys not found!"
End If

End Sub

Private Sub List1_DblClick()
List1.Clear

End Sub

Private Sub tmrDetect_Timer()
' perform periodic tasks for the application
Static bDevPrevPresent As Boolean
  
  ' check if UsbI2cIo device is still present
bDevicePresent = CheckDevice
If (bDevicePresent) = True Then
    shpDetect.FillColor = vbGreen
    lblDetect.Caption = "OK"
Else
    shpDetect.FillColor = vbRed
    lblDetect.Caption = "NG"
    OpenDevice
End If
bDevPrevPresent = bDevicePresent

End Sub
Private Sub InitGrdData()
Dim i As Integer

grdData.Clear
grdData.Cols = 4
grdData.Rows = ProcessCount + 1

grdData.ColWidth(0) = 450
grdData.ColWidth(1) = 2600
grdData.ColWidth(2) = 700
grdData.ColWidth(3) = 700

For i = 0 To grdData.Cols - 1
    grdData.ColAlignment(i) = flexAlignCenterCenter
Next i
grdData.ColAlignment(1) = flexAlignLeftCenter

grdData.TextMatrix(0, 0) = "No"
grdData.TextMatrix(0, 1) = "Item"
grdData.TextMatrix(0, 2) = "Status"
grdData.TextMatrix(0, 3) = "Time"

For i = 1 To grdData.Rows - 1
    grdData.TextMatrix(i, 0) = i
    grdData.TextMatrix(i, 1) = ProcessItem(i - 1)
Next i

grdData.Row = 0
grdData.Col = 0

End Sub

Public Sub AppExit()
  ' this subroutine performs cleanup tasks and then exits the application
CloseDiHandle
End
End Sub

Private Function SRQ(Index As Integer)
Dim portMask As Long
Dim pushData As Long

SRQLOactive = False

Select Case Index

    Case SrqLow  ''Low
    
    '''0x000CBBAA
    portMask = &H80000
    pushData = &H80000
    
    If DAPI_WriteIoPorts(hDevInstance, pushData, portMask) = 0 Then
        List1.AddItem "SRQ : LOW NG!"
        SRQ = NG
        Exit Function
    End If
    List1.AddItem "SRQ : LOW OK!"
    DoEvents
    
    delay_ms Val(txtSRQLO.Text)
    SRQLOactive = True
    
    Case SrqHigh  ''High
    
    '''0x000CBBAA
    portMask = &H80000
    pushData = &H0
    
    If DAPI_WriteIoPorts(hDevInstance, pushData, portMask) = 0 Then
        List1.AddItem "SRQ : HIGH NG!"
        SRQ = NG
        Exit Function
    End If
    List1.AddItem "SRQ : HIGH OK!"
    DoEvents
    
    delay_ms Val(txtSRQHI.Text)
     
End Select

SRQ = OK

End Function

Private Function BUS(Index As Integer)
Dim I2cTrans As I2C_TRANS                     ' Dimension an I2C_TRANS structure
Dim lWritten As Long                          ' Dimension a long to hold the returned value

I2cTrans.byDevId = &HE0
I2cTrans.byType = I2C_TRANS_TYPE.I2C_TRANS_8ADR
I2cTrans.wCount.hi = 0                        ' only writing 1 byte, so set to 0
I2cTrans.wCount.lo = 1                        ' writing 1 byte, so set to 1

Select Case Index

Case I2C0

    I2cTrans.wMemAddr.hi = 0
    I2cTrans.wMemAddr.lo = &H4
    I2cTrans.Data(0) = &H4
  
    ' call the function
    lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
    If lWritten = 1 Then
        List1.AddItem "BUS SELECT : IIC0 OK!"
        BUS = OK
    Else
        List1.AddItem "BUS SELECT : IIC0 NG!"
        BUS = IICError
    End If
    
Case I2C1

    I2cTrans.wMemAddr.hi = 0
    I2cTrans.wMemAddr.lo = &H5
    I2cTrans.Data(0) = &H5
  
    ' call the function
    lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
    If lWritten = 1 Then
        List1.AddItem "BUS SELECT : IIC1 OK!"
        BUS = OK
    Else
        List1.AddItem "BUS SELECT : IIC1 NG!"
        BUS = IICError
    End If

End Select

DoEvents
delay_ms Val(txtBUS.Text)

End Function
Private Function WriteProtect()
Dim I2cTrans As I2C_TRANS
Dim lWritten As Long

I2cTrans.byType = I2C_TRANS_8ADR
I2cTrans.byDevId = &H70
I2cTrans.wMemAddr.hi = 0
I2cTrans.wMemAddr.lo = &H8B

I2cTrans.wCount.hi = 0
I2cTrans.wCount.lo = 1

I2cTrans.Data(0) = 0

lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
If lWritten <> 1 Then
    List1.AddItem "Write Protect...IIC Error!"
    WriteProtect = IICError
    Exit Function
End If
delay_ms 50

I2cTrans.byType = I2C_TRANS_8ADR
I2cTrans.byDevId = &H70
I2cTrans.wMemAddr.hi = 0
I2cTrans.wMemAddr.lo = &H88

I2cTrans.wCount.hi = 0
I2cTrans.wCount.lo = 1

I2cTrans.Data(0) = 0

lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
If lWritten <> 1 Then
    List1.AddItem "Write Protect...IIC Error!"
    WriteProtect = IICError
    Exit Function
End If
delay_ms 50

List1.AddItem "Write Protect...OK!"
DoEvents
WriteProtect = OK
WPactive = True

End Function

Private Function WriteData(add_data As String)
Dim I2cTrans As I2C_TRANS
Dim I2CTrans_READ As I2C_TRANS
Dim lWritten As Long
Dim lRead As Long
Dim splitData() As String
Dim splitADD() As String

splitData = Split(add_data, ",")
splitADD = Split(splitData(0), "-")

'''WRITE
If UBound(splitADD) = 1 Then
    I2cTrans.byType = I2C_TRANS_8ADR
    I2cTrans.byDevId = Val("&H" & splitADD(0))
    I2cTrans.wMemAddr.hi = 0
    I2cTrans.wMemAddr.lo = Val("&H" & splitADD(1))
Else
    I2cTrans.byType = I2C_TRANS_16ADR
    I2cTrans.byDevId = Val("&H" & splitADD(0))
    I2cTrans.wMemAddr.hi = Val("&H" & splitADD(1))
    I2cTrans.wMemAddr.lo = Val("&H" & splitADD(2))
End If

I2cTrans.wCount.hi = 0
I2cTrans.wCount.lo = 1

I2cTrans.Data(0) = Val("&H" & splitData(1))

lWritten = DAPI_WriteI2c(hDevInstance, I2cTrans)
If lWritten <> 1 Then
    List1.AddItem "Write Data [ADD:" & splitData(0) & " DATA:" & splitData(1) & "...IIC Error!"
    WriteData = IICError
    Exit Function
End If
delay_ms Val(txtWRITE.Text)


'''READ
If UBound(splitADD) = 1 Then
    I2CTrans_READ.byType = I2C_TRANS_8ADR
    I2CTrans_READ.byDevId = Val("&H" & splitADD(0))
    I2CTrans_READ.wMemAddr.hi = 0
    I2CTrans_READ.wMemAddr.lo = Val("&H" & splitADD(1))
Else
    I2CTrans_READ.byType = I2C_TRANS_16ADR
    I2CTrans_READ.byDevId = Val("&H" & splitADD(0))
    I2CTrans_READ.wMemAddr.hi = Val("&H" & splitADD(1))
    I2CTrans_READ.wMemAddr.lo = Val("&H" & splitADD(2))
End If

I2CTrans_READ.wCount.hi = 0
I2CTrans_READ.wCount.lo = 1

lRead = DAPI_ReadI2c(hDevInstance, I2CTrans_READ)
If lRead <> 1 Then
    List1.AddItem "Read Data [ADD:" & splitData(0) & " DATA:" & splitData(1) & "...IIC Error!"
    WriteData = IICError
    Exit Function
End If
delay_ms Val(txtVERIFY.Text)

If Val("&H" & splitData(1)) <> I2CTrans_READ.Data(0) Then
    List1.AddItem "Verify Data [ADD:" & splitData(0) & " DATA:" & splitData(1) & "...NG!"
    WriteData = NG
    Exit Function
End If

List1.AddItem "Write + Verify Data [ADD:" & splitData(0) & " DATA:" & splitData(1) & "...OK!"
DoEvents
WriteData = OK

End Function

Private Sub ButtonEnable(stat As Boolean)

Me.cmdExit.Enabled = stat
Me.cmdI2CCheck.Enabled = stat
Me.cmdIO.Enabled = stat
Me.cmdModelChange.Enabled = stat
Me.cmdProcess.Enabled = stat
Me.cmdProp.Enabled = stat
Me.cmdStart.Enabled = stat

End Sub

Private Sub ResetUSB()
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
    List1.AddItem "Board was successfully reset!"
Else
    List1.AddItem "Board was not successfully reset!"
End If

SRQ SrqHigh

End Sub

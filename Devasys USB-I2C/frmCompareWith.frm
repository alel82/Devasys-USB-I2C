VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCompareWith 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compare Table With >>>"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8760
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
   ScaleHeight     =   1935
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse File"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   4
      ToolTipText     =   "Browse"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<< Back"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      ToolTipText     =   "Browse"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Proceed >>>"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      ToolTipText     =   "Browse"
      Top             =   720
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "File Name"
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
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.TextBox txtFilename 
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   6135
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6000
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCompareWith"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
Unload frmCompareResult
Unload Me
End Sub

Private Sub cmdBrowse_Click()

CommonDialog1.FileName = ""
CommonDialog1.Filter = "Data Files (*.dat,*.ver)|*.dat;*.ver|All Files (*.*)|*.*"
CommonDialog1.DialogTitle = "Load Data from File"
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowOpen

If CommonDialog1.FileName = "" Then Exit Sub

cmdProceed.Enabled = True
txtFilename.Text = CommonDialog1.FileTitle
CompareFilePath = CommonDialog1.FileName
CompareFileName = txtFilename.Text

End Sub


Private Sub cmdProceed_Click()
Me.Hide
frmCompareResult.Show vbModal

End Sub

Private Sub Form_Load()
cmdProceed.Enabled = False
End Sub

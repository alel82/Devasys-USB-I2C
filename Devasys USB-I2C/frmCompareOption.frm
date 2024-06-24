VERSION 5.00
Begin VB.Form frmCompareOption 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "     Compare Table With >>>"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2910
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
   ScaleHeight     =   2775
   ScaleWidth      =   2910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "&BACK"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   2055
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton cmdFile 
         Caption         =   "&FILE"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton cmdEEPROM 
         Caption         =   "EEPROM &2"
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton cmdEEPROM 
         Caption         =   "EEPROM &1"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmCompareOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub cmdEEPROM_Click(Index As Integer)
If Index = 0 Then CompareWith = CompareEEPROM1
If Index = 1 Then CompareWith = CompareEEPROM2
Unload Me
End Sub

Private Sub cmdFile_Click()

CompareWith = CompareFile
CompareFilePath = ""
Unload Me
frmCompareFile.Show vbModal

End Sub

Private Sub Form_Load()

For i = 0 To 1
If EEPROMEnable(i) = False Then cmdEEPROM(i).Enabled = False Else cmdEEPROM(i).Enabled = True
Next i

End Sub

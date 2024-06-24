VERSION 5.00
Begin VB.Form frmCompareFile 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compare With File"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
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
   ScaleHeight     =   3705
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "&Proceed"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtFilePath 
      Height          =   645
      Left            =   7080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.FileListBox File1 
         Height          =   2490
         Left            =   3360
         Pattern         =   "*.ver;*.dat"
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
      Begin VB.DirListBox Dir1 
         Height          =   2250
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3135
      End
      Begin VB.DriveListBox Drive1 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmCompareFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
CompareFilePath = ""
Unload Me
End Sub

Private Sub cmdProceed_Click()
CompareFilePath = txtFilePath.Text
Unload Me
End Sub

Private Sub Dir1_Change()

txtFilePath.Text = ""
File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()

On Error GoTo NextDrive

txtFilePath.Text = ""
Dir1.Path = Drive1.Drive

Exit Sub

NextDrive:
Drive1.Drive = "C:\"

End Sub

Private Sub File1_Click()

Me.txtFilePath.Text = File1.Path & "\" & File1.FileName

End Sub


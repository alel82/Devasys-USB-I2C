VERSION 5.00
Begin VB.Form frmModelChange 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Model Change"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
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
   ScaleHeight     =   7335
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox DirChk 
      Height          =   900
      Left            =   10560
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2415
   End
   Begin VB.DirListBox DirChk2 
      Height          =   900
      Left            =   10560
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10560
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   9975
      Begin VB.CommandButton cmdSkip 
         Caption         =   "&Skip"
         Height          =   495
         Left            =   240
         TabIndex        =   21
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000009&
         Height          =   1455
         Left            =   120
         TabIndex        =   18
         Top             =   5520
         Width           =   1455
         Begin VB.CommandButton cmdOK 
            Caption         =   "&OK"
            Height          =   495
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   495
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000009&
         Caption         =   "EEPROM 2"
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
         Height          =   3330
         Left            =   1680
         TabIndex        =   11
         Top             =   3600
         Width           =   8175
         Begin VB.CheckBox chkEnable 
            BackColor       =   &H80000009&
            Caption         =   "Enable"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.DirListBox dirEEPROM2 
            Height          =   1980
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   4335
         End
         Begin VB.DriveListBox drvEEPROM2 
            Height          =   360
            Left            =   120
            TabIndex        =   4
            Top             =   2880
            Width           =   4335
         End
         Begin VB.FileListBox filEEPROM 
            Height          =   2250
            Index           =   1
            Left            =   4560
            Pattern         =   "*.ver"
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox txtEEPROM 
            Height          =   315
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   480
            Width           =   7935
         End
      End
      Begin VB.Frame frmEEPROM1 
         BackColor       =   &H80000009&
         Caption         =   "EEPROM 1"
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
         Height          =   3330
         Left            =   1680
         TabIndex        =   9
         Top             =   120
         Width           =   8175
         Begin VB.CheckBox chkEnable 
            BackColor       =   &H80000009&
            Caption         =   "Enable"
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txtEEPROM 
            Height          =   315
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   480
            Width           =   7935
         End
         Begin VB.FileListBox filEEPROM 
            Height          =   2250
            Index           =   0
            Left            =   4560
            Pattern         =   "*.ver"
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   840
            Width           =   3495
         End
         Begin VB.DriveListBox drvEEPROM1 
            Height          =   360
            Left            =   120
            TabIndex        =   1
            Top             =   2880
            Width           =   4335
         End
         Begin VB.DirListBox dirEEPROM1 
            Height          =   1980
            Left            =   120
            TabIndex        =   2
            Top             =   840
            Width           =   4335
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
         Caption         =   "Chassis"
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
         TabIndex        =   8
         Top             =   120
         Width           =   1455
         Begin VB.ComboBox cmbChassis 
            Height          =   360
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmModelChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkEnable_Click(Index As Integer)

If chkEnable(Index).Caption = "Enable" Then
    chkEnable(Index).Caption = "Disable"
    EEPROMEnable(Index) = False
    Exit Sub
ElseIf chkEnable(Index).Caption = "Disable" Then
    chkEnable(Index).Caption = "Enable"
    EEPROMEnable(Index) = True
    Exit Sub
End If

End Sub

Private Sub cmbChassis_Click()
Dim EEPROM1_Path As String
Dim EEPROM2_Path As String
Dim foundE1 As Integer
Dim foundE2 As Integer

If cmbChassis.Text = "" Then Exit Sub

foundE1 = 0
foundE2 = 0

EEPROM1_Path = App.Path & "\Verify1\" & cmbChassis.Text
EEPROM2_Path = App.Path & "\Verify2\" & cmbChassis.Text

'Show

reTRY:
DirChk.Path = App.Path
DirChk.Refresh
If DirChk.ListCount > 0 Then
    For i = 0 To DirChk.ListCount - 1
        If DirChk.List(i) = App.Path & "\Verify1" Then
reTRY2:
            DirChk2.Path = DirChk.List(i)
            DirChk2.Refresh
            If DirChk2.ListCount > 0 Then
                For j = 0 To DirChk2.ListCount - 1
                    If DirChk2.List(j) = DirChk.List(i) & "\" & cmbChassis.Text Then foundE1 = 1: GoTo nextI
                Next j
                ChDir DirChk2.Path
                MkDir cmbChassis.Text
                foundE1 = 1
            Else
                ChDir DirChk2.Path
                MkDir cmbChassis.Text
                GoTo reTRY2
            End If
            
        ElseIf DirChk.List(i) = App.Path & "\Verify2" Then
reTRY3:
            DirChk2.Path = DirChk.List(i)
            DirChk2.Refresh
            If DirChk2.ListCount > 0 Then
                For j = 0 To DirChk2.ListCount - 1
                    If DirChk2.List(j) = DirChk.List(i) & "\" & cmbChassis.Text Then foundE2 = 1: GoTo nextI
                Next j
                ChDir DirChk2.Path
                MkDir cmbChassis.Text
                foundE2 = 1
            Else
                ChDir DirChk2.Path
                MkDir cmbChassis.Text
                GoTo reTRY3
            End If
          
        End If
        
nextI:
    Next i
    
Else

reTRY4:

    ChDrive Mid(App.Path, 1, 1)
    If foundE1 = 0 Then MkDir App.Path & "\Verify1\"
    If foundE2 = 0 Then MkDir App.Path & "\Verify2\"
    GoTo reTRY
    
End If

If foundE1 + foundE2 <> 2 Then GoTo reTRY4

ChDir App.Path

dirEEPROM1.Path = EEPROM1_Path
dirEEPROM2.Path = EEPROM2_Path


End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()

If txtEEPROM(0).Text = "" And EEPROMEnable(0) = True Then MsgBox "EEPROM 1 not selected!", vbExclamation, Me.Caption: Exit Sub
If txtEEPROM(1).Text = "" And EEPROMEnable(1) = True Then MsgBox "EEPROM 2 not selected!", vbExclamation, Me.Caption: Exit Sub

Chassis = cmbChassis.Text
lpKeyName = "Chassis"
lpKeyValue = Chassis
WritePrivateProfileString "Current Settings", lpKeyName, lpKeyValue, App.Path & "\" & "Current.ini"

For i = 0 To 1
    If EEPROMEnable(i) = True Then
        EEPROM_Filename(i) = filEEPROM(i).FileName
        VerifyPath(i) = txtEEPROM(i).Text
    Else
        EEPROM_Filename(i) = "-"
    End If
Next i

Unload Me
frmMain.Show

End Sub

Private Sub dirEEPROM1_Change()

filEEPROM(0).Path = dirEEPROM1.Path

End Sub
Private Sub dirEEPROM2_Change()

filEEPROM(1).Path = dirEEPROM2.Path

End Sub


Private Sub filEEPROM_Click(Index As Integer)

txtEEPROM(Index).Text = filEEPROM(Index).Path & "\" & filEEPROM(Index).FileName

End Sub

Private Sub Form_Activate()

Text1.SetFocus

End Sub

Private Sub Form_Load()

Load_Current
EEPROMEnable(0) = True
EEPROMEnable(1) = True

End Sub
Private Sub Load_Current()

Dim ChassisCount As Integer

cmbChassis.Clear

lpFileName$ = App.Path & "\Current.ini"

lpApplicationName$ = "Chassis"
lpDefault = ""
nSize = 128
      
lpReturnedString$ = Space$(128)
lpKeyName$ = "Count"
n% = GetPrivateProfileString%(lpApplicationName$, lpKeyName$, lpDefault, lpReturnedString$, nSize, lpFileName$)
lpReturnedString$ = Left$(lpReturnedString$, InStr(lpReturnedString$, Chr$(0)) - 1)
ChassisCount = (lpReturnedString$)

For i = 1 To ChassisCount

    lpReturnedString$ = Space$(128)
    lpKeyName$ = "Chassis" & i
    n% = GetPrivateProfileString%(lpApplicationName$, lpKeyName$, lpDefault, lpReturnedString$, nSize, lpFileName$)
    lpReturnedString$ = Left$(lpReturnedString$, InStr(lpReturnedString$, Chr$(0)) - 1)
    If lpReturnedString$ <> "" Then cmbChassis.AddItem (lpReturnedString$)

Next i

lpApplicationName$ = "Current Settings"
lpDefault = ""
nSize = 128
      
lpReturnedString$ = Space$(128)
lpKeyName$ = "Chassis"
n% = GetPrivateProfileString%(lpApplicationName$, lpKeyName$, lpDefault, lpReturnedString$, nSize, lpFileName$)
lpReturnedString$ = Left$(lpReturnedString$, InStr(lpReturnedString$, Chr$(0)) - 1)
For i = 1 To ChassisCount
    If lpReturnedString$ = cmbChassis.List(i - 1) Then
        cmbChassis.Text = (lpReturnedString$)
    End If
Next i

End Sub

VERSION 5.00
Begin VB.Form frmUSBProp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USB12CI0 Properties"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
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
   ScaleHeight     =   5535
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoadKM 
      Caption         =   "Load KM Settings"
      Height          =   495
      Left            =   5880
      TabIndex        =   69
      Top             =   120
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "IO Properties"
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
      Height          =   5415
      Left            =   3360
      TabIndex        =   34
      Top             =   0
      Width           =   2415
      Begin VB.CheckBox Check2 
         Caption         =   "Output [ 0 ]"
         Height          =   255
         Left            =   960
         TabIndex        =   67
         Top             =   5040
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Input    [ 1 ]"
         Height          =   255
         Left            =   960
         TabIndex        =   65
         Top             =   4800
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkPortC 
         BackColor       =   &H8000000E&
         Caption         =   "PC7"
         Height          =   240
         Index           =   7
         Left            =   1680
         TabIndex        =   64
         Top             =   4320
         Width           =   700
      End
      Begin VB.CheckBox chkPortC 
         BackColor       =   &H8000000E&
         Caption         =   "PC6"
         Height          =   240
         Index           =   6
         Left            =   1680
         TabIndex        =   63
         Top             =   4080
         Width           =   700
      End
      Begin VB.CheckBox chkPortC 
         BackColor       =   &H8000000E&
         Caption         =   "PC5"
         Height          =   240
         Index           =   5
         Left            =   1680
         TabIndex        =   62
         Top             =   3840
         Width           =   700
      End
      Begin VB.CheckBox chkPortC 
         BackColor       =   &H8000000E&
         Caption         =   "PC4"
         Height          =   240
         Index           =   4
         Left            =   1680
         TabIndex        =   61
         Top             =   3600
         Width           =   700
      End
      Begin VB.CheckBox chkPortC 
         BackColor       =   &H8000000E&
         Caption         =   "PC3"
         Enabled         =   0   'False
         Height          =   225
         Index           =   3
         Left            =   960
         TabIndex        =   60
         Top             =   4320
         Width           =   975
      End
      Begin VB.CheckBox chkPortC 
         BackColor       =   &H8000000E&
         Caption         =   "PC2"
         Enabled         =   0   'False
         Height          =   225
         Index           =   2
         Left            =   960
         TabIndex        =   59
         Top             =   4080
         Width           =   975
      End
      Begin VB.CheckBox chkPortC 
         BackColor       =   &H8000000E&
         Caption         =   "PC1"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   960
         TabIndex        =   58
         Top             =   3840
         Width           =   975
      End
      Begin VB.CheckBox chkPortC 
         BackColor       =   &H8000000E&
         Caption         =   "PC0"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   960
         TabIndex        =   57
         Top             =   3600
         Width           =   975
      End
      Begin VB.CheckBox chkPortB 
         BackColor       =   &H8000000E&
         Caption         =   "PB7"
         Height          =   240
         Index           =   7
         Left            =   1680
         TabIndex        =   56
         Top             =   2880
         Width           =   700
      End
      Begin VB.CheckBox chkPortB 
         BackColor       =   &H8000000E&
         Caption         =   "PB6"
         Height          =   240
         Index           =   6
         Left            =   1680
         TabIndex        =   55
         Top             =   2640
         Width           =   700
      End
      Begin VB.CheckBox chkPortB 
         BackColor       =   &H8000000E&
         Caption         =   "PB5"
         Height          =   240
         Index           =   5
         Left            =   1680
         TabIndex        =   54
         Top             =   2400
         Width           =   700
      End
      Begin VB.CheckBox chkPortB 
         BackColor       =   &H8000000E&
         Caption         =   "PB4"
         Height          =   240
         Index           =   4
         Left            =   1680
         TabIndex        =   53
         Top             =   2160
         Width           =   700
      End
      Begin VB.CheckBox chkPortB 
         BackColor       =   &H8000000E&
         Caption         =   "PB3"
         Height          =   225
         Index           =   3
         Left            =   960
         TabIndex        =   52
         Top             =   2880
         Width           =   975
      End
      Begin VB.CheckBox chkPortB 
         BackColor       =   &H8000000E&
         Caption         =   "PB2"
         Height          =   225
         Index           =   2
         Left            =   960
         TabIndex        =   51
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox chkPortB 
         BackColor       =   &H8000000E&
         Caption         =   "PB1"
         Height          =   225
         Index           =   1
         Left            =   960
         TabIndex        =   50
         Top             =   2400
         Width           =   975
      End
      Begin VB.CheckBox chkPortB 
         BackColor       =   &H8000000E&
         Caption         =   "PB0"
         Height          =   225
         Index           =   0
         Left            =   960
         TabIndex        =   49
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox chkPortA 
         BackColor       =   &H8000000E&
         Caption         =   "PA7"
         Height          =   240
         Index           =   7
         Left            =   1680
         TabIndex        =   48
         Top             =   1440
         Width           =   700
      End
      Begin VB.CheckBox chkPortA 
         BackColor       =   &H8000000E&
         Caption         =   "PA6"
         Height          =   240
         Index           =   6
         Left            =   1680
         TabIndex        =   47
         Top             =   1200
         Width           =   700
      End
      Begin VB.CheckBox chkPortA 
         BackColor       =   &H8000000E&
         Caption         =   "PA5"
         Height          =   240
         Index           =   5
         Left            =   1680
         TabIndex        =   46
         Top             =   960
         Width           =   700
      End
      Begin VB.CheckBox chkPortA 
         BackColor       =   &H8000000E&
         Caption         =   "PA4"
         Height          =   240
         Index           =   4
         Left            =   1680
         TabIndex        =   45
         Top             =   720
         Width           =   700
      End
      Begin VB.CheckBox chkPortA 
         BackColor       =   &H8000000E&
         Caption         =   "PA3"
         Height          =   225
         Index           =   3
         Left            =   960
         TabIndex        =   44
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkPortA 
         BackColor       =   &H8000000E&
         Caption         =   "PA2"
         Height          =   225
         Index           =   2
         Left            =   960
         TabIndex        =   43
         Top             =   1200
         Width           =   975
      End
      Begin VB.CheckBox chkPortA 
         BackColor       =   &H8000000E&
         Caption         =   "PA1"
         Height          =   225
         Index           =   1
         Left            =   960
         TabIndex        =   42
         Top             =   960
         Width           =   975
      End
      Begin VB.CheckBox chkPortA 
         BackColor       =   &H8000000E&
         Caption         =   "PA0"
         Height          =   225
         Index           =   0
         Left            =   960
         TabIndex        =   41
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtPortC 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   3240
         Width           =   495
      End
      Begin VB.TextBox txtPortB 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtPortA 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "Note :"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Port C"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   3285
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Port B"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1845
         Width           =   735
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Port A"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   400
         Width           =   735
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Height          =   495
         Left            =   120
         TabIndex        =   68
         Top             =   4800
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5880
      TabIndex        =   9
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton cmdLoadDefault 
      Caption         =   "Load Default"
      Height          =   495
      Left            =   7680
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSaveProp 
      Caption         =   "Save Properties"
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   360
      Left            =   8160
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   720
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000E&
         Caption         =   "I2C Properties"
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
         Height          =   3135
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   2895
         Begin VB.Frame Frame4 
            BackColor       =   &H8000000E&
            Height          =   975
            Left            =   120
            TabIndex        =   25
            Top             =   2040
            Width           =   2655
            Begin VB.TextBox txtHex 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   960
               MaxLength       =   2
               TabIndex        =   28
               Top             =   200
               Width           =   1455
            End
            Begin VB.Label lblBit 
               Alignment       =   2  'Center
               Caption         =   "00"
               Height          =   255
               Index           =   4
               Left            =   960
               TabIndex        =   33
               Top             =   600
               Width           =   255
            End
            Begin VB.Label lblBit 
               Alignment       =   2  'Center
               Caption         =   "0"
               Height          =   255
               Index           =   3
               Left            =   1320
               TabIndex        =   32
               Top             =   600
               Width           =   135
            End
            Begin VB.Label lblBit 
               Alignment       =   2  'Center
               Caption         =   "0"
               Height          =   255
               Index           =   2
               Left            =   1560
               TabIndex        =   31
               Top             =   600
               Width           =   135
            End
            Begin VB.Label lblBit 
               Alignment       =   2  'Center
               Caption         =   "0"
               Height          =   255
               Index           =   1
               Left            =   1800
               TabIndex        =   30
               Top             =   600
               Width           =   135
            End
            Begin VB.Label lblBit 
               Alignment       =   2  'Center
               Caption         =   "000"
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   29
               Top             =   600
               Width           =   375
            End
            Begin VB.Label Label16 
               Caption         =   " Binary"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label15 
               Caption         =   " Hex"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.Label lblI2CProp 
            Alignment       =   2  'Center
            Caption         =   "RESERVED FIELD"
            Height          =   255
            Index           =   4
            Left            =   960
            TabIndex        =   24
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label lblI2CProp 
            Alignment       =   2  'Center
            Caption         =   "AUTO REDIRECT A2"
            Height          =   255
            Index           =   3
            Left            =   960
            TabIndex        =   23
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblI2CProp 
            Alignment       =   2  'Center
            Caption         =   "POLL EEPROM ACK"
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   22
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lblI2CProp 
            Alignment       =   2  'Center
            Caption         =   "IGNORE NAK"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   21
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblI2CProp 
            Alignment       =   2  'Center
            Caption         =   "RETRY FIELD"
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   20
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label9 
            Caption         =   " Bit 6 ~ 7"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   " Bit 5"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   " Bit 4"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   " Bit 3"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   " Bit 0 ~ 2"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.ComboBox cmbChannel 
         Height          =   360
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   200
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "I2C Clock Rate"
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
         TabIndex        =   2
         Top             =   600
         Width           =   2535
         Begin VB.ComboBox cmbChannelClock 
            Height          =   360
            Index           =   2
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1040
            Width           =   1335
         End
         Begin VB.ComboBox cmbChannelClock 
            Height          =   360
            Index           =   1
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   680
            Width           =   1335
         End
         Begin VB.ComboBox cmbChannelClock 
            Height          =   360
            Index           =   0
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   320
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Channel 2"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Channel 1"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Channel 0"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "I2C Default Channel"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmUSBProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FormLoad As Boolean

Private Sub Save_Properties()
Dim propvalue As Byte

propvalue = USBI2CIO_PROPERTY_COMMAND.CMD_STORE_TABLE_TO_EEPROM

If DAPI_SetProperty(hDevInstance, USBI2CIO_PROPERTY_INDEX.WRCOMMAND_RDCOUNT, propvalue) = False Then
    MsgBox "Error!", vbExclamation, "Save Properties Error"
End If


End Sub

Private Sub Check1_Click()

Check1.Value = 1

End Sub

Private Sub Check2_Click()
Check2.Value = 0
End Sub

Private Sub chkPortA_Click(Index As Integer)
Dim t As Integer

If FormLoad = True Then Exit Sub

t = 0
For i = 0 To 7
    t = t + (2 ^ i) * (chkPortA(i).Value)
Next i

txtPortA.Text = Hex(t)
If Len(txtPortA.Text) = 1 Then txtPortA.Text = "0" & txtPortA.Text

If DAPI_SetProperty(hDevInstance, USBI2CIO_PROPERTY_INDEX.IO_CONFIG_PORTA, t) = False Then
    MsgBox "Error!", vbExclamation, "Set Property Error"
    Exit Sub
End If

End Sub

Private Sub chkPortB_Click(Index As Integer)
Dim t As Integer

If FormLoad = True Then Exit Sub

t = 0
For i = 0 To 7
    t = t + (2 ^ i) * (chkPortB(i).Value)
Next i

txtPortB.Text = Hex(t)
If Len(txtPortB.Text) = 1 Then txtPortB.Text = "0" & txtPortB.Text

If DAPI_SetProperty(hDevInstance, USBI2CIO_PROPERTY_INDEX.IO_CONFIG_PORTB, t) = False Then
    MsgBox "Error!", vbExclamation, "Set Property Error"
    Exit Sub
End If

End Sub

Private Sub chkPortC_Click(Index As Integer)
Dim t As Integer

If FormLoad = True Then Exit Sub

t = 0
For i = 4 To 7
    t = t + (2 ^ (i - 4)) * (chkPortC(i).Value)
Next i

txtPortC.Text = "F" & Hex(t)
If Len(txtPortC.Text) = 1 Then txtPortC.Text = txtPortC.Text

t = "&H" & (txtPortC.Text)
'''Debug.Print Hex(t)
If DAPI_SetProperty(hDevInstance, USBI2CIO_PROPERTY_INDEX.IO_CONFIG_PORTC, t) = False Then
    MsgBox "Error!", vbExclamation, "Set Property Error"
    Exit Sub
End If


End Sub

Private Sub cmbChannel_Click()

If FormLoad = True Then Exit Sub

If DAPI_SetProperty(hDevInstance, USBI2CIO_PROPERTY_INDEX.I2C_DEFAULT_CHANNEL, cmbChannel.ListIndex) = False Then
    MsgBox "Error!", vbExclamation, "Set Property Error"
    Exit Sub
End If

End Sub
Private Sub cmbChannelClock_Click(Index As Integer)
Dim propvalue As Byte

If FormLoad = True Then Exit Sub

Select Case cmbChannelClock(Index).ListIndex

Case 0  '90 kHz
propvalue = I2C_CLOCK.PROP_I2C_90kHz

Case 1  '100 kHz
propvalue = I2C_CLOCK.PROP_I2C_100kHz

Case 2  '400 kHz
propvalue = I2C_CLOCK.PROP_I2C_400kHz

Case 3  '1 MHz
propvalue = I2C_CLOCK.PROP_I2C_1MHz

End Select

Select Case Index

Case 0
    
If DAPI_SetProperty(hDevInstance, USBI2CIO_PROPERTY_INDEX.I2C_CHAN0_CLK_LO, propvalue) = False Then
    MsgBox "Error!", vbExclamation, "Set Property Error"
    Exit Sub
End If
    
Case 1
    
If DAPI_SetProperty(hDevInstance, USBI2CIO_PROPERTY_INDEX.I2C_CHAN1_CLK_LO, propvalue) = False Then
    MsgBox "Error!", vbExclamation, "Set Property Error"
    Exit Sub
End If
    
Case 2
    
If DAPI_SetProperty(hDevInstance, USBI2CIO_PROPERTY_INDEX.I2C_CHAN2_CLK_LO, propvalue) = False Then
    MsgBox "Error!", vbExclamation, "Set Property Error"
    Exit Sub
End If
    
End Select

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdLoadKM_Click()

cmbChannel.Text = cmbChannel.List(1)

For i = 0 To 2
    cmbChannelClock(i).Text = cmbChannelClock(i).List(1)
Next i

txtHex.Text = "15"

For i = 0 To 7
    chkPortA(i).Value = 0
    If i >= 4 Then chkPortC(i).Value = 0
Next i

DoEvents

cmdSaveProp_Click

End Sub

Private Sub cmdSaveProp_Click()

t = MsgBox("Save current settings?", vbYesNo, "Save Properties")
If t = vbYes Then Save_Properties

Unload Me

End Sub

Private Sub Form_Load()
FormLoad = True

Load_Combo
Load_Properties

For i = 0 To 4
    lblBit(i).ToolTipText = lblI2CProp(i).Caption
Next i

FormLoad = False
End Sub
Private Sub Load_Combo()

For i = 0 To 2
    cmbChannel.List(i) = "Channel " & i
    cmbChannelClock(i).List(0) = "90 kHz"
    cmbChannelClock(i).List(1) = "100 kHz"
    cmbChannelClock(i).List(2) = "400 kHz"
    cmbChannelClock(i).List(3) = "1 MHz"
Next i

End Sub
Private Sub Load_Properties()
Dim propvalue As Byte
Dim tmpPA, tmpPB, tmpPC, tmpBin As Integer

If DAPI_GetProperty(hDevInstance, propvalue, USBI2CIO_PROPERTY_INDEX.I2C_DEFAULT_CHANNEL) = False Then
    MsgBox "Error!", vbExclamation, "Get Property Error"
    Exit Sub
Else
    cmbChannel.Text = "Channel " & Hex(propvalue)
End If

If DAPI_GetProperty(hDevInstance, propvalue, USBI2CIO_PROPERTY_INDEX.I2C_CHAN0_CLK_LO) = False Then
    MsgBox "Error!", vbExclamation, "Get Property Error"
    Exit Sub
Else
    Select Case propvalue
    Case I2C_CLOCK.PROP_I2C_90kHz
        cmbChannelClock(0).Text = "90 kHz"
    Case I2C_CLOCK.PROP_I2C_100kHz
        cmbChannelClock(0).Text = "100 kHz"
    Case I2C_CLOCK.PROP_I2C_400kHz
        cmbChannelClock(0).Text = "400 kHz"
    Case I2C_CLOCK.PROP_I2C_1MHz
        cmbChannelClock(0).Text = "1 MHz"
    End Select
End If

If DAPI_GetProperty(hDevInstance, propvalue, USBI2CIO_PROPERTY_INDEX.I2C_CHAN1_CLK_LO) = False Then
    MsgBox "Error!", vbExclamation, "Get Property Error"
    Exit Sub
Else
    Select Case propvalue
    Case I2C_CLOCK.PROP_I2C_90kHz
        cmbChannelClock(1).Text = "90 kHz"
    Case I2C_CLOCK.PROP_I2C_100kHz
        cmbChannelClock(1).Text = "100 kHz"
    Case I2C_CLOCK.PROP_I2C_400kHz
        cmbChannelClock(1).Text = "400 kHz"
    Case I2C_CLOCK.PROP_I2C_1MHz
        cmbChannelClock(1).Text = "1 MHz"
    End Select
End If

If DAPI_GetProperty(hDevInstance, propvalue, USBI2CIO_PROPERTY_INDEX.I2C_CHAN2_CLK_LO) = False Then
    MsgBox "Error!", vbExclamation, "Get Property Error"
    Exit Sub
Else
    Select Case propvalue
    Case I2C_CLOCK.PROP_I2C_90kHz
        cmbChannelClock(2).Text = "90 kHz"
    Case I2C_CLOCK.PROP_I2C_100kHz
        cmbChannelClock(2).Text = "100 kHz"
    Case I2C_CLOCK.PROP_I2C_400kHz
        cmbChannelClock(2).Text = "400 kHz"
    Case I2C_CLOCK.PROP_I2C_1MHz
        cmbChannelClock(2).Text = "1 MHz"
    End Select
End If

txtHex.Text = ""
If DAPI_GetProperty(hDevInstance, propvalue, USBI2CIO_PROPERTY_INDEX.I2C_CONFIG) = False Then
    MsgBox "Error!", vbExclamation, "Get Property Error"
    Exit Sub
Else
    txtHex.Text = Hex(propvalue)
    If Len(txtHex.Text) = 1 Then txtHex.Text = "0" & txtHex.Text
    Hex2Bin "&H" & Hex(propvalue)
End If

txtPortA.Text = ""
If DAPI_GetProperty(hDevInstance, propvalue, USBI2CIO_PROPERTY_INDEX.IO_CONFIG_PORTA) = False Then
    MsgBox "Error!", vbExclamation, "Get Property Error"
    Exit Sub
Else
    txtPortA.Text = Hex(propvalue)
    If Len(txtPortA.Text) = 1 Then txtPortA.Text = "0" & txtPortA.Text
    tmpPA = propvalue
    For i = 0 To 7
        tmpBin = tmpPA Mod 2
        chkPortA(i).Value = tmpBin
        tmpPA = tmpPA \ 2
    Next i
End If


txtPortB.Text = ""
If DAPI_GetProperty(hDevInstance, propvalue, USBI2CIO_PROPERTY_INDEX.IO_CONFIG_PORTB) = False Then
    MsgBox "Error!", vbExclamation, "Get Property Error"
    Exit Sub
Else
    txtPortB.Text = Hex(propvalue)
    If Len(txtPortB.Text) = 1 Then txtPortB.Text = "0" & txtPortB.Text
    tmpPB = propvalue
    For i = 0 To 7
        tmpBin = tmpPB Mod 2
        chkPortB(i).Value = tmpBin
        tmpPB = tmpPB \ 2
    Next i
End If

txtPortC.Text = ""
If DAPI_GetProperty(hDevInstance, propvalue, USBI2CIO_PROPERTY_INDEX.IO_CONFIG_PORTC) = False Then
    MsgBox "Error!", vbExclamation, "Get Property Error"
    Exit Sub
Else
    txtPortC.Text = Hex(propvalue)
    If Len(txtPortC.Text) = 1 Then txtPortC.Text = "0" & txtPortC.Text
    tmpPC = Val("&H" & (Mid(txtPortC.Text, 2, 1)))
    For i = 0 To 7
        If i >= 4 Then
            tmpBin = tmpPC Mod 2
            chkPortC(i).Value = tmpBin
            tmpPC = tmpPC \ 2
        End If
    Next i
End If

Text1.Text = ""

End Sub

Private Sub Hex2Bin(hexData As Long)
Dim tmpBin As String
Dim tmpCalc As Integer

For i = 0 To 4
    lblBit(i).Caption = ""
Next i

tmpBin = ""

For i = 0 To 7 '00 0 0 0 000
    tmpBin = (hexData Mod 2) & tmpBin
    If Len(tmpBin) = 3 Then lblBit(0).Caption = tmpBin
    If Len(tmpBin) = 4 Then lblBit(1).Caption = Left(tmpBin, 1)
    If Len(tmpBin) = 5 Then lblBit(2).Caption = Left(tmpBin, 1)
    If Len(tmpBin) = 6 Then lblBit(3).Caption = Left(tmpBin, 1)
    If Len(tmpBin) = 8 Then lblBit(4).Caption = Left(tmpBin, 2)
    hexData = hexData \ 2
Next i

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To 4
    lblI2CProp(i).BackColor = &H8000000F
Next i

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To 4
    lblI2CProp(i).BackColor = &H8000000F
Next i

End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

For i = 0 To 4
    lblI2CProp(i).BackColor = &H8000000F
Next i

End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To 4
    lblI2CProp(i).BackColor = &H8000000F
Next i

End Sub

Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To 4
    lblI2CProp(i).BackColor = &H8000000F
Next i

End Sub


Private Sub lblBit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tmp As Long

For i = 0 To 4
    lblI2CProp(i).BackColor = &H8000000F
Next i

lblI2CProp(Index).BackColor = vbGreen


End Sub

Private Sub txtHex_Change()
Dim propvalue As Byte

If FormLoad = True Or txtHex.Text = "" Then Exit Sub

If Len(txtHex.Text) = 2 Then
    If IsNumeric("&H" & txtHex.Text) = True Then
        Hex2Bin "&H" & txtHex.Text
        propvalue = "&H" & (txtHex.Text)
        If DAPI_SetProperty(hDevInstance, USBI2CIO_PROPERTY_INDEX.I2C_CONFIG, propvalue) = False Then
            MsgBox "Error!", vbExclamation, "Set Property Error"
            Exit Sub
        End If
    Else
        MsgBox "Not a valid Hex value!", vbExclamation, "I2C Properties"
        txtHex.SetFocus
        txtHex.SelStart = 0
        txtHex.SelLength = 2
    End If
End If

End Sub

Private Sub txtHex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To 4
    lblI2CProp(i).BackColor = &H8000000F
Next i
End Sub


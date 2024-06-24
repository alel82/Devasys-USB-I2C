VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VsFlex7L.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCompareResult 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compare Result"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
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
   ScaleHeight     =   6645
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Height          =   5895
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5535
      Begin VB.ListBox lstResult 
         Height          =   4380
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label lblFileName 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   5520
         Width           =   3975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Table Data"
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   5160
         Width           =   3975
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   4800
         Width           =   3975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Total"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Right Data"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Left Data"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   5160
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSaveResult 
      Caption         =   "Save Result"
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
      Left            =   3600
      TabIndex        =   3
      ToolTipText     =   "Browse"
      Top             =   6000
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
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Browse"
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   5295
      Index           =   0
      Left            =   7440
      TabIndex        =   0
      Top             =   120
      Width           =   6650
      Begin VSFlex7LCtl.VSFlexGrid grdData 
         Height          =   4905
         Left            =   120
         TabIndex        =   1
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
         ShowComboButton =   -1  'True
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
End
Attribute VB_Name = "frmCompareResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload Me
frmCompareWith.Show vbModal
End Sub

Private Sub cmdSaveResult_Click()
Dim mFile As String

Me.MousePointer = vbHourglass
DoEvents

mFile = App.Path & "\" & "Result_" & lblFileName.Caption & ".txt"

Open mFile For Output As #1

For j = 0 To lstResult.ListCount - 1
    Print #1, lstResult.List(j)
Next j

Print #1, "@@"

Close #1

Me.MousePointer = vbNormal

End Sub

Private Sub Form_Load()
Dim ResultAdd As String
Dim lAdd As String
Dim hAdd As String

Me.MousePointer = vbHourglass
DoEvents

lstResult.Clear
InitGrdData
LoadData

''Load Table to array
k = 0
For i = 1 To grdData.Rows - 1
    For j = 1 To 16
        ReDim Preserve CompareTable2(k)
        CompareTable2(k) = Trim(grdData.TextMatrix(i, j))
        k = k + 1
    Next j
Next i
'''

If UBound(CompareTable1) <> UBound(CompareTable2) Then
    MsgBox "Bytes of data not match!", vbExclamation, "Compare"
    cmdBack_Click
End If

j = 0
k = 0
For i = 0 To UBound(CompareTable1)
    lAdd = Hex(j)
    If Len(lAdd) = 1 Then lAdd = "0" & lAdd
    
    hAdd = Hex(k)
    If Len(hAdd) = 1 Then hAdd = "0" & hAdd
    
    ResultAdd = hAdd & lAdd
    
    If CompareTable1(i) <> CompareTable2(i) Then
        lstResult.AddItem ResultAdd & " : " & CompareTable1(i) & " : " & CompareTable2(i)
    End If
    j = j + 1
    If j > 255 Then
        j = 0
        k = k + 1
    End If
Next i

lblTotal.Caption = lstResult.ListCount
lblFileName.Caption = CompareFileName

Me.MousePointer = vbNormal
End Sub
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

Private Sub LoadData()

Open CompareFilePath For Input As #1

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

End Sub

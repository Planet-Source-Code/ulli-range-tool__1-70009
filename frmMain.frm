VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\A..\..\..\..\..\PROGRA~1\MIAF9D~1\VB98\VBSOUR~1\RANGET~1\RangeTool.vbp"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "RangeTool Test"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin pRangeTool.RangeTool RangeTool1 
      Height          =   705
      Left            =   195
      TabIndex        =   20
      Top             =   105
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   1244
      Color           =   255
      BorderStyle     =   2
   End
   Begin VB.CommandButton btColor 
      Caption         =   "Color Select"
      Height          =   375
      Left            =   690
      TabIndex        =   17
      Top             =   3150
      Width           =   1020
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   75
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton btEnd 
      Caption         =   "Close"
      Height          =   375
      Left            =   1830
      TabIndex        =   16
      Top             =   3150
      Width           =   1020
   End
   Begin VB.CheckBox chkUnitsOnly 
      Caption         =   "Units only"
      Height          =   195
      Left            =   2070
      TabIndex        =   15
      Top             =   2700
      Width           =   1005
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      Height          =   195
      Left            =   2070
      TabIndex        =   14
      Top             =   1230
      Value           =   1  'Aktiviert
      Width           =   885
   End
   Begin VB.CheckBox chkLowerLocked 
      Caption         =   "Lower Locked"
      Height          =   195
      Left            =   2070
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CheckBox chkLowerVisible 
      Caption         =   "Lower Visible"
      Height          =   195
      Left            =   2070
      TabIndex        =   11
      Top             =   2115
      Value           =   1  'Aktiviert
      Width           =   1245
   End
   Begin VB.CheckBox chkUpperLocked 
      Caption         =   "Upper Locked"
      Height          =   195
      Left            =   2070
      TabIndex        =   10
      Top             =   1815
      Width           =   1335
   End
   Begin VB.CheckBox chkUpperVisible 
      Caption         =   "Upper Visible"
      Height          =   195
      Left            =   2070
      TabIndex        =   9
      Top             =   1530
      Value           =   1  'Aktiviert
      Width           =   1245
   End
   Begin VB.ComboBox lstStyle 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   660
      List            =   "frmMain.frx":0013
      Style           =   2  'Dropdown-Liste
      TabIndex        =   8
      Top             =   2580
      Width           =   1110
   End
   Begin VB.TextBox txtUpper 
      Height          =   285
      Left            =   660
      TabIndex        =   6
      Text            =   "75"
      Top             =   2220
      Width           =   1080
   End
   Begin VB.TextBox txtLower 
      Height          =   285
      Left            =   660
      TabIndex        =   4
      Text            =   "25"
      Top             =   1890
      Width           =   1080
   End
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   660
      TabIndex        =   2
      Text            =   "100"
      Top             =   1560
      Width           =   1080
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Left            =   660
      TabIndex        =   0
      Text            =   "0"
      Top             =   1230
      Width           =   1080
   End
   Begin VB.Label Lbl 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Resolution"
      Height          =   195
      Left            =   690
      TabIndex        =   19
      Top             =   885
      Width           =   750
   End
   Begin VB.Label lbRes 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2205
      TabIndex        =   18
      Top             =   885
      Width           =   45
   End
   Begin VB.Label lblBorder 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Border"
      Height          =   195
      Left            =   105
      TabIndex        =   13
      Top             =   2640
      Width           =   465
   End
   Begin VB.Label lblUpper 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Upper"
      Height          =   195
      Left            =   135
      TabIndex        =   7
      Top             =   2280
      Width           =   435
   End
   Begin VB.Label lblLower 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Lower"
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   1950
      Width           =   435
   End
   Begin VB.Label lblMax 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Max"
      Height          =   195
      Left            =   270
      TabIndex        =   3
      Top             =   1620
      Width           =   300
   End
   Begin VB.Label lblMin 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Min"
      Height          =   195
      Left            =   315
      TabIndex        =   1
      Top             =   1290
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub btColor_Click()

    With cDlg
        On Error Resume Next
            .Flags = cdlCCRGBInit
            .Color = RangeTool1.Color
            .ShowColor
            If Err = 0 Then
                RangeTool1.Color = .Color
            End If
        On Error GoTo 0
    End With 'CDLG

End Sub

Private Sub btEnd_Click()

    Unload Me

End Sub

Private Sub chkEnabled_Click()

    RangeTool1.Enabled = (chkEnabled = vbChecked)

End Sub

Private Sub chkLowerLocked_Click()

    RangeTool1.LowerLocked = (chkLowerLocked = vbChecked)

End Sub

Private Sub chkLowerVisible_Click()

    RangeTool1.LowerVisible = (chkLowerVisible = vbChecked)

End Sub

Private Sub chkUnitsOnly_Click()

    RangeTool1.UnitsOnly = (chkUnitsOnly = vbChecked)
    lbRes = RangeTool1.Resolution

End Sub

Private Sub chkUpperLocked_Click()

    RangeTool1.UpperLocked = (chkUpperLocked = vbChecked)

End Sub

Private Sub chkUpperVisible_Click()

    RangeTool1.UpperVisible = (chkUpperVisible = vbChecked)

End Sub

Private Sub Form_Initialize()

    InitCommonControls

End Sub

Private Sub Form_Load()

    With RangeTool1
        lstStyle.ListIndex = .BorderStyle
        txtMin = .MinValue
        txtMax = .MaxValue
        txtLower = .LowerValue
        txtUpper = .UpperValue
        chkLowerLocked = IIf(.LowerLocked, vbChecked, vbUnchecked)
        chkUpperLocked = IIf(.UpperLocked, vbChecked, vbUnchecked)
        chkLowerVisible = IIf(.LowerVisible, vbChecked, vbUnchecked)
        chkUpperVisible = IIf(.UpperVisible, vbChecked, vbUnchecked)
        chkEnabled = IIf(.Enabled, vbChecked, vbUnchecked)
        chkUnitsOnly = IIf(.UnitsOnly, vbChecked, vbUnchecked)
        lbRes = .Resolution
    End With 'RANGETOOL1

End Sub

Private Sub lstStyle_Click()

    RangeTool1.BorderStyle = lstStyle.ListIndex

End Sub

Private Sub RangeTool1_Change(LowerValue As Single, UpperValue As Single)

    txtMin = RangeTool1.MinValue
    txtMax = RangeTool1.MaxValue
    txtLower = LowerValue
    txtUpper = UpperValue

End Sub

Private Sub RangeTool1_Scroll(LowerValue As Single, UpperValue As Single)

    txtLower = LowerValue
    txtUpper = UpperValue

End Sub

Private Sub txtLower_Validate(Cancel As Boolean)

    RangeTool1.LowerValue = Val(txtLower)

End Sub

Private Sub txtMax_Validate(Cancel As Boolean)

    RangeTool1.MaxValue = Val(txtMax)
    lbRes = RangeTool1.Resolution

End Sub

Private Sub txtMin_Validate(Cancel As Boolean)

    RangeTool1.MinValue = Val(txtMin)
    lbRes = RangeTool1.Resolution

End Sub

Private Sub txtUpper_Validate(Cancel As Boolean)

    RangeTool1.UpperValue = Val(txtUpper)

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-29 16:05)  Decl: 3  Code: 134  Total: 137 Lines
':) CommentOnly: 2 (1,5%)  Commented: 2 (1,5%)  Empty: 53 (38,7%)  Max Logic Depth: 3

VERSION 5.00
Begin VB.UserControl RangeTool 
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3225
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   69
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   215
   ToolboxBitmap   =   "RangeTool.ctx":0000
   Begin VB.PictureBox picLower 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   336
      Left            =   105
      MouseIcon       =   "RangeTool.ctx":0532
      MousePointer    =   99  'Benutzerdefiniert
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   1
      Top             =   555
      Width           =   372
   End
   Begin VB.PictureBox picUpper 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   336
      Left            =   75
      MouseIcon       =   "RangeTool.ctx":0684
      MousePointer    =   99  'Benutzerdefiniert
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   0
      Top             =   105
      Width           =   372
   End
   Begin VB.Shape shpFill 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   180
      Left            =   720
      Top             =   405
      Width           =   570
   End
End
Attribute VB_Name = "RangeTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'[RangeTool]
Option Explicit

'Range Tool

'Found on the WWW (sorry, don't remember the author)
'but credit goes to him anyway

'Extensively modified by Ulli (umgedv@yahoo.com)

'Note:

'To adjust the width to match values and pixels: the left and right margins are 7 pixels each,
'so if for example the control has a width of 215 pixels or 3225 twips the active slider range
'is 201 pixels - numbered from 0 to 200 and then scaled up or down by the min and max value
'difference to yield the thumb value. Note that all values are Single data types, therefore
'due to rounding errors when scaling pixels to values you may get slight differences.
'The best policy is to size the control width to a multiple of (Max - Min) pixels and
'then add 15 pixels. Use the resolution property as a guide.

'types
Private Type Rect
    L       As Long
    T       As Long
    R       As Long
    B       As Long
End Type

'enums
Public Enum BorderStyle
    [RT Flat] = 0
    [RT Raised] = 1
    [RT Sunken] = 2
    [RT Fillet] = 3
    [RT Ridge] = 4
End Enum

'variables
Private i                       As Long
Private Position                As Single
Private TmpLower                As Single
Private TmpUpper                As Single
Private DrawRange               As Single
Private Range                   As Single
Private DrawRatio               As Single
Private MouseDownOffset         As Single
Private Grabbed                 As Boolean
Private Rect                    As Rect

'properties
Private myMin                   As Single
Private myMax                   As Single
Private myLowerValue            As Single
Private myUpperValue            As Single
Private myLowerLocked           As Boolean
Private myUpperLocked           As Boolean
Private myLowerVisible          As Boolean
Private myUpperVisible          As Boolean
Private myUnitsOnly             As Boolean 'causes the thumbs to skip non-integer values; this may in rare cases lead to
'                                           an oscillating thumb when the true thumb position is exactly 'in between' pixels.
'                                           This is due to the thumb jumping to the next posn and thereby generating
'                                           a mouse move event which directs the thumb back to it's previous position
'                                           (I hope you can make sense out of this...)

Private myColor                 As OLE_COLOR
Private myBorderStyle           As BorderStyle
Private myByUserOnly            As Boolean 'causes the control to fire events thru user interaction only but not when
'                                           values are altered programmatically

Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qRC As Rect, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Const BDR_RAISEDOUTER   As Long = 1
Private Const BDR_SUNKENOUTER   As Long = 2
Private Const BDR_RAISEDINNER   As Long = 4
Private Const BDR_SUNKENINNER   As Long = 8
Private Const BDR_FILLET        As Long = BDR_RAISEDOUTER Or BDR_SUNKENINNER
Private Const BDR_RIDGE         As Long = BDR_SUNKENOUTER Or BDR_RAISEDINNER
Private Const BDR_RAISED        As Long = BDR_RAISEDOUTER Or BDR_RAISEDINNER
Private Const BDR_SUNKEN        As Long = BDR_SUNKENOUTER Or BDR_SUNKENINNER
Private Const BF_RECT           As Long = 15

Private Const BS                As String = "BorderStyle"
Private Const BUO               As String = "ByUserOnly"
Private Const COL               As String = "Color"
Private Const ENA               As String = "Enabled"
Private Const LLK               As String = "LowerLocked"
Private Const ULK               As String = "UpperLocked"
Private Const LVAL              As String = "LowerValue"
Private Const UVAL              As String = "UpperValue"
Private Const LVIS              As String = "LowerVisible"
Private Const UVIS              As String = "UpperVisible"
Private Const VALMAX            As String = "MaxValue"
Private Const VALMIN            As String = "MinValue"
Private Const UNON              As String = "UnitsOnly"

'events
Public Event Change(LowerValue As Single, UpperValue As Single)
Attribute Change.VB_Description = "Fired when the thumbs age moved (see ByUserOnly property)"
Attribute Change.VB_MemberFlags = "200"
Public Event Scroll(LowerValue As Single, UpperValue As Single)
Attribute Scroll.VB_Description = "Fired when the thumbs are moved (see ByUserOnly property)"

Public Property Let BorderStyle(nuBorderStyle As BorderStyle)
Attribute BorderStyle.VB_Description = "What it says: Borderstyle"

    Select Case nuBorderStyle
      Case [RT Fillet], [RT Flat], [RT Raised], [RT Ridge], [RT Sunken]
        myBorderStyle = nuBorderStyle
        Refresh
        PropertyChanged BS
      Case Else
        Err.Raise 380, UserControl.Name
    End Select

End Property

Public Property Get BorderStyle() As BorderStyle

    BorderStyle = myBorderStyle

End Property

Public Property Let ByUserOnly(nuByUserOnly As Boolean)
Attribute ByUserOnly.VB_Description = "When True will fire Change Events only whe the user moves the thumbs, not when you change the value internally."

    myByUserOnly = (nuByUserOnly <> False)
    PropertyChanged BUO

End Property

Public Property Get ByUserOnly() As Boolean

    ByUserOnly = myByUserOnly

End Property

Public Property Let Color(nuColor As OLE_COLOR)
Attribute Color.VB_Description = "Bar Color"

    myColor = nuColor
    shpFill.BorderColor = nuColor
    shpFill.FillColor = nuColor
    PropertyChanged COL

End Property

Public Property Get Color() As OLE_COLOR

    Color = shpFill.FillColor

End Property

Public Property Get Enabled() As Boolean

    Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(nuEnabled As Boolean)

    UserControl.Enabled = (nuEnabled <> False)
    PropertyChanged ENA
    picUpper.Refresh
    picLower.Refresh
    shpFill.BorderColor = IIf(UserControl.Enabled, myColor, &H808080)
    shpFill.FillColor = shpFill.BorderColor
    Refresh

End Property

Public Property Get LowerLocked() As Boolean

    LowerLocked = myLowerLocked

End Property

Public Property Let LowerLocked(ByVal nuLocked As Boolean)

    myLowerLocked = (nuLocked <> False)
    PropertyChanged LLK
    picLower.MousePointer = IIf(myLowerLocked, 0, 99)
    picLower.Refresh

End Property

Public Property Get LowerValue() As Single

    LowerValue = myLowerValue

End Property

Public Property Let LowerValue(ByVal nuValue As Single)

    If nuValue <> myLowerValue Then
        myLowerValue = Minimum(myUpperValue, Maximum(myMin, Minimum(myMax, nuValue)))
        UpdatePosition
        If myByUserOnly = False Then
            RaiseEvent Change(myLowerValue, myUpperValue)
        End If
        PropertyChanged LVAL
    End If

End Property

Public Property Get LowerVisible() As Boolean

    LowerVisible = myLowerVisible

End Property

Public Property Let LowerVisible(ByVal nuVisible As Boolean)

    myLowerVisible = (nuVisible <> False)
    picLower.Visible = myLowerVisible
    PropertyChanged LVIS

End Property

Private Function Maximum(Val1 As Single, Val2 As Single) As Single

    If Val1 > Val2 Then
        Maximum = Val1
      Else 'NOT VAL1...
        Maximum = Val2
    End If

End Function

Public Property Get MaxValue() As Single

    MaxValue = myMax

End Property

Public Property Let MaxValue(ByVal nuMax As Single)

    If nuMax <> myMax Then
        TmpLower = myLowerValue
        TmpUpper = myUpperValue
        myMax = Maximum(nuMax, myMin)
        UpdatePosition
        PropertyChanged VALMAX
        If myLowerValue <> TmpLower Or myUpperValue <> TmpUpper Then
            If myByUserOnly = False Then
                RaiseEvent Change(myLowerValue, myUpperValue)
            End If
        End If
    End If

End Property

Private Function Minimum(Val1 As Single, Val2 As Single) As Single

    If Val1 < Val2 Then
        Minimum = Val1
      Else 'NOT VAL1...
        Minimum = Val2
    End If

End Function

Public Property Get MinValue() As Single

    MinValue = myMin

End Property

Public Property Let MinValue(ByVal nuMin As Single)

    If nuMin <> myMin Then
        TmpLower = myLowerValue
        TmpUpper = myUpperValue
        myMin = Minimum(myMax, nuMin)
        UpdatePosition
        PropertyChanged VALMIN
        If myLowerValue <> TmpLower Or myUpperValue <> TmpUpper Then
            If myByUserOnly = False Then
                RaiseEvent Change(myLowerValue, myUpperValue)
            End If
        End If
    End If

End Property

Private Sub piclower_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not myLowerLocked Then
        Grabbed = True
        MouseDownOffset = X
        picLower.Refresh
    End If

End Sub

Private Sub picLower_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Grabbed Then
        Position = Maximum(0, Minimum(picLower.Left + X - MouseDownOffset, DrawRange))
        If myUpperLocked Then
            Position = Minimum(Position, picUpper.Left)
          Else 'MYUPPERLOCKED = FALSE/0
            myUpperValue = Maximum(myUpperValue, myMin + Position * DrawRatio)
        End If
        myLowerValue = myMin + Position * DrawRatio
        UpdatePosition
        RaiseEvent Scroll(myLowerValue, myUpperValue)
        PropertyChanged LVAL
    End If

End Sub

Private Sub picLower_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Grabbed = False
    UpdatePosition
    picLower.Refresh
    RaiseEvent Change(myLowerValue, myUpperValue)
    PropertyChanged LVAL

End Sub

Private Sub picLower_Paint()

  'paints the lower triangle

    For i = 0 To 14
        picLower.Line (7 - i / 2, i)-(7 + i / 2, i), vbButtonFace
    Next i
    If myLowerLocked Or Extender.Enabled = False Then
        'Ridged
        picLower.Line (1, 13)-(13, 13), vbButtonShadow 'bottom
        picLower.Line (8, 1)-(2, 13), vb3DHighlight 'left
        picLower.Line (7, 1)-(0, 14), vbButtonShadow 'left
        picLower.Line (7, 1)-(13, 13), vbButtonShadow 'right
        picLower.Line (8, 1)-(14, 14), vb3DHighlight 'right
        picLower.Line (1, 14)-(14, 14), vb3DHighlight 'bottom
      ElseIf Grabbed Then 'NOT MYLOWERLOCKED...
        'Indented
        picLower.Line (1, 13)-(8, 0), vbButtonShadow 'left
        picLower.Line (7, 1)-(14, 14), vb3DHighlight 'right
        picLower.Line (0, 14)-(7, 0), vb3DDKShadow 'left
        picLower.Line (0, 14)-(14, 14), vb3DHighlight 'bottom
      Else 'GRABBED = FALSE/0
        'Raised
        picLower.Line (0, 14)-(14, 14), vb3DDKShadow 'bottom
        picLower.Line (1, 13)-(13, 13), vbButtonShadow 'bottom
        picLower.Line (7, 0)-(13, 13), vbButtonShadow 'right
        picLower.Line (7, 1)-(14, 14), vb3DDKShadow 'right
        picLower.Line (0, 13)-(7, 0), vb3DHighlight 'left
    End If

End Sub

Private Sub picUpper_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not myUpperLocked Then
        Grabbed = True
        MouseDownOffset = X
        picUpper.Refresh
    End If

End Sub

Private Sub picUpper_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Grabbed Then
        Position = Maximum(0, Minimum(picUpper.Left + X - MouseDownOffset, DrawRange))
        If myLowerLocked Then
            Position = Maximum(Position, picLower.Left)
          Else 'MYLOWERLOCKED = FALSE/0
            myLowerValue = Minimum(myLowerValue, myMin + Position * DrawRatio)
        End If
        myUpperValue = myMin + Position * DrawRatio
        UpdatePosition
        RaiseEvent Scroll(myLowerValue, myUpperValue)
        PropertyChanged UVAL
    End If

End Sub

Private Sub picUpper_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Grabbed = False
    UpdatePosition
    picUpper.Refresh
    RaiseEvent Change(myLowerValue, myUpperValue)
    PropertyChanged UVAL

End Sub

Private Sub picUpper_Paint()

  'paints the upper triangle

    For i = 0 To 14
        picUpper.Line (i / 2, i)-(14 - i / 2, i), vbButtonFace
    Next i

    If myUpperLocked Or Extender.Enabled = False Then
        'Ridged
        picUpper.Line (1, 1)-(13, 1), vb3DHighlight
        picUpper.Line (1, 0)-(7, 12), vb3DHighlight
        picUpper.Line (14, 1)-(7, 14), vb3DHighlight
        picUpper.Line (13, 0)-(7, 13), vbButtonShadow
        picUpper.Line (0, 0)-(14, 0), vbButtonShadow
        picUpper.Line (0, 0)-(7, 14), vbButtonShadow
      ElseIf Grabbed Then 'NOT MYUPPERLOCKED...
        'Indented
        picUpper.Line (14, 0)-(7, 14), vb3DHighlight
        picUpper.Line (1, 1)-(13, 1), vbButtonShadow
        picUpper.Line (1, 1)-(7, 13), vbButtonShadow
        picUpper.Line (0, 0)-(14, 0), vb3DDKShadow
        picUpper.Line (0, 0)-(7, 14), vb3DDKShadow
      Else 'GRABBED = FALSE/0
        'Raised
        picUpper.Line (14, 0)-(7, 14), vb3DDKShadow
        picUpper.Line (13, 0)-(7, 13), vbButtonShadow
        picUpper.Line (0, 0)-(14, 0), vb3DHighlight
        picUpper.Line (0, 0)-(7, 14), vb3DHighlight
    End If

End Sub

Public Property Get Resolution() As Single
Attribute Resolution.VB_Description = "Informative only."

    If myUnitsOnly Then
        Resolution = 1
      Else 'MYUNITSONLY = FALSE/0
        Resolution = (myMax - myMin) / (Width / 15 - 15)
    End If

End Property

Public Property Let Resolution(ByVal nuResolution As Single)

    nuResolution = nuResolution 'dummy - do someting with param so this prop is not empty and param is used

End Property

Public Property Get UnitsOnly() As Boolean
Attribute UnitsOnly.VB_Description = "The thumbs will only stop permit fractional values."

    UnitsOnly = myUnitsOnly

End Property

Public Property Let UnitsOnly(ByVal nuUnitsOnly As Boolean)

    myUnitsOnly = (nuUnitsOnly <> False)
    PropertyChanged UNON
    UpdatePosition
    If myByUserOnly = False Then
        RaiseEvent Change(myLowerValue, myUpperValue)
    End If

End Property

Private Sub UpdatePosition()

  'Draw the thumbs in the correct places

    myLowerValue = Maximum(myMin, Minimum(myMax, myLowerValue))
    myUpperValue = Maximum(myMin, Minimum(myMax, myUpperValue))
    If myUnitsOnly Then
        myLowerValue = Int(myLowerValue + 0.5 * Sgn(myLowerValue))
        myUpperValue = Int(myUpperValue + 0.5 * Sgn(myUpperValue))
    End If
    DrawRange = ScaleWidth - 15
    Range = myMax - myMin
    DrawRatio = Range / DrawRange
    If myMin = myMax Then
        picUpper.Move DrawRange, 0
        picLower.Move 0, ScaleHeight - picLower.Height
      Else 'NOT MYMIN...
        On Error Resume Next
            picUpper.Move ((myUpperValue - myMin) / Range) * DrawRange, 0
            picLower.Move ((myLowerValue - myMin) / Range) * DrawRange, ScaleHeight - picLower.Height
        On Error GoTo 0
    End If
    shpFill.Move picLower.Left + 7, picUpper.Height + 4, picUpper.Left - picLower.Left, ScaleHeight - 2 * picLower.Height - 8

End Sub

Public Property Get UpperLocked() As Boolean

    UpperLocked = myUpperLocked

End Property

Public Property Let UpperLocked(ByVal nuLocked As Boolean)

    myUpperLocked = (nuLocked <> False)
    PropertyChanged ULK
    picUpper.MousePointer = IIf(nuLocked, 0, 99)
    picUpper.Refresh

End Property

Public Property Get UpperValue() As Single

    UpperValue = myUpperValue

End Property

Public Property Let UpperValue(ByVal nuValue As Single)

    If nuValue <> myUpperValue Then
        myUpperValue = Maximum(Maximum(myMin, Minimum(myMax, nuValue)), myLowerValue)
        UpdatePosition
        If myByUserOnly = False Then
            RaiseEvent Change(myLowerValue, myUpperValue)
        End If
        PropertyChanged UVAL
    End If

End Property

Public Property Get UpperVisible() As Boolean

    UpperVisible = myUpperVisible

End Property

Public Property Let UpperVisible(ByVal nuVisible As Boolean)

    myUpperVisible = (nuVisible <> False)
    picUpper.Visible = myUpperVisible
    PropertyChanged UVIS

End Property

Private Sub UserControl_AmbientChanged(PropertyName As String)

  'adjust color to container's color

    BackColor = Ambient.BackColor
    picUpper.BackColor = BackColor
    picLower.BackColor = BackColor

End Sub

Private Sub UserControl_Initialize()

    picUpper.Width = 15
    picUpper.Height = 15
    picLower.Width = 15
    picLower.Height = 15

End Sub

Private Sub UserControl_InitProperties()

    myMin = 0
    myMax = 100
    myLowerValue = 25
    myUpperValue = 75
    myBorderStyle = [RT Sunken]
    myLowerVisible = True
    picLower.Visible = True
    myUpperVisible = True
    picUpper.Visible = True
    myColor = vbRed
    shpFill.BorderColor = myColor
    shpFill.FillColor = myColor
    UserControl_AmbientChanged "" 'adjust color to container color
    UpdatePosition

End Sub

Private Sub UserControl_Paint()

    With Rect
        .L = 3
        .T = picUpper.Height
        .R = ScaleWidth - 4
        .B = ScaleHeight - picLower.Height
    End With 'RECT
    If Extender.Enabled Then
        Select Case myBorderStyle
          Case [RT Fillet]
            DrawEdge hDC, Rect, BDR_FILLET, BF_RECT
          Case [RT Ridge]
            DrawEdge hDC, Rect, BDR_RIDGE, BF_RECT
          Case [RT Raised]
            DrawEdge hDC, Rect, BDR_RAISED, BF_RECT
          Case [RT Sunken]
            DrawEdge hDC, Rect, BDR_SUNKEN, BF_RECT
        End Select
      Else 'EXTENDER.ENABLED = FALSE/0
        If myBorderStyle <> [RT Flat] Then
            DrawEdge hDC, Rect, BDR_RIDGE, BF_RECT
        End If
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        myByUserOnly = .ReadProperty(BUO, False)
        myMin = .ReadProperty(VALMIN, 0)
        myMax = .ReadProperty(VALMAX, 100)
        myLowerValue = .ReadProperty(LVAL, 25)
        myUpperValue = .ReadProperty(UVAL, 75)
        myColor = .ReadProperty(COL, vbHighlight)
        shpFill.BorderColor = myColor
        shpFill.FillColor = myColor
        LowerVisible = .ReadProperty(LVIS, True)
        UpperVisible = .ReadProperty(UVIS, True)
        myUnitsOnly = .ReadProperty(UNON, False)
        myLowerLocked = .ReadProperty(LLK, False)
        myUpperLocked = .ReadProperty(ULK, False)
        BorderStyle = .ReadProperty(BS, [RT Ridge])
    End With 'PROPBAG
    UserControl_AmbientChanged ""
    picUpper.MousePointer = IIf(myUpperLocked, 0, 99)
    picLower.MousePointer = IIf(myLowerLocked, 0, 99)
    UpdatePosition

End Sub

Private Sub UserControl_Resize()

    If ScaleHeight < 40 Then
        Height = 40 * Screen.TwipsPerPixelY
    End If
    If ScaleWidth < 40 Then
        Width = 40 * Screen.TwipsPerPixelX
    End If
    UpdatePosition

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty BUO, myByUserOnly, False
        .WriteProperty VALMIN, myMin, 0
        .WriteProperty VALMAX, myMax, 100
        .WriteProperty LVAL, myLowerValue, 25
        .WriteProperty UVAL, myUpperValue, 75
        .WriteProperty COL, myColor, vbHighlight
        .WriteProperty LVIS, myLowerVisible, True
        .WriteProperty UVIS, myUpperVisible, True
        .WriteProperty UNON, myUnitsOnly, False
        .WriteProperty LLK, myLowerLocked, False
        .WriteProperty ULK, myUpperLocked, False
        .WriteProperty BS, myBorderStyle, [RT Ridge]
    End With 'PROPBAG

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-29 16:05)  Decl: 97  Code: 549  Total: 646 Lines
':) CommentOnly: 35 (5,4%)  Commented: 33 (5,1%)  Empty: 157 (24,3%)  Max Logic Depth: 4

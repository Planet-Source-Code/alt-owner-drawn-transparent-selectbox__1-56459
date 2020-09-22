VERSION 5.00
Begin VB.UserControl SelectBox 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   HitBehavior     =   0  'None
   ScaleHeight     =   225
   ScaleWidth      =   1200
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   1230
   End
End
Attribute VB_Name = "SelectBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************************************************************************
' CustomSelectBox.ctl
' Simply a custom drawn checkbox to demonstrate how to draw on a transparent
' UserControl.
' Copyright © 2004, Alan Tucker, All Rights Reserved.
'**************************************************************************************************
Option Explicit
'**************************************************************************************************
' simply maintaining the proper case of my enum items...
' thanks to Evan Toder for the tip.
'**************************************************************************************************
#If False Then
     Dim None
     Dim FixedSingle
     Dim Round
     Dim Square
     Dim Diamond
     Dim LeftJustify
     Dim Center
     Dim RightJustify
#End If

'**************************************************************************************************
' Constants
'**************************************************************************************************
Private Const BOX_GAP = 20
Private Const DT_LEFT = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_CALCRECT = &H400
Private Const KEY_PRESSED As Integer = &H1000
Private Const PS_SOLID = 0
Private Const BS_SOLID = 0
Private Const HS_SOLID = 8

'**************************************************************************************************
' Enum/Struct Declarations
'**************************************************************************************************
Public Enum CTL_BORDERSTYLE
     [None]
     [FixedSingle]
End Enum ' CTL_BORDERSTYLE

Public Enum CTL_BOXSTYLE
     [Round]
     [Square]
     [Diamond]
End Enum ' CTL_BULLETSYLE

Public Enum LABEL_ALIGN
     [LeftJustify]
     [Center]
     [RightJustify]
End Enum ' LABEL_ALIGN

Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type ' LOGBRUSH

Private Type POINTAPI
    X As Long
    Y As Long
End Type ' POINTAPI

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type ' RECT

'**************************************************************************************************
' Win32 API Declarations
'**************************************************************************************************
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
     ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, _
     ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, _
     ByVal lpDrawTextParams As Any) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
     ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, _
     lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
     ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
     ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, _
     ByVal Y As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, _
     ByVal nCount As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
     ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
     ByVal hObject As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, _
     lpPoint As POINTAPI) As Long

'**************************************************************************************************
' Private Variable Declarations
'**************************************************************************************************
Private m_CaptionRect As RECT
Private m_Flag As Long
Private m_hBrush As Long
Private m_hPen As Long
Private m_LogBrush As LOGBRUSH
Private m_oldBrush As Long
Private m_oldPen As Long

'**************************************************************************************************
' Property Value Constant Declarations
'**************************************************************************************************
Const m_def_Alignment = False
Const m_def_BorderStyle = False
Const m_def_BoxBackgroundColor = vbWhite
Const m_def_BoxBorderColor = vbBlack
Const m_def_BoxBorderWidth = 2
Const m_def_BoxStyle = False
Const m_def_SelectMarkColor = vbRed
Const m_def_SelectValue = False
Const m_def_WordWrap = False

'**************************************************************************************************
' Property Variable Declarations
'**************************************************************************************************
Dim m_Alignment As LABEL_ALIGN
Dim m_BoxBackgroundColor As OLE_COLOR
Dim m_BoxBorderColor As OLE_COLOR
Dim m_BoxBorderWidth As Long
Dim m_BoxStyle As CTL_BOXSTYLE
Dim m_Caption As String
Dim m_bFocus As Boolean
Dim m_SelectMarkColor As OLE_COLOR
Dim m_SelectValue As Boolean
Dim m_WordWrap As Boolean

'**************************************************************************************************
' Control Event Declarations
'**************************************************************************************************
Public Event Click()
Public Event DblClick()
Public Event MouseDown(ByVal X As Single, ByVal Y As Single)
Public Event MouseMove(ByVal X As Single, ByVal Y As Single)
Public Event MouseUp(ByVal X As Single, ByVal Y As Single)

'**************************************************************************************************
' SelectBox Property Let/Get/Set Declarations
'**************************************************************************************************
Public Property Get Alignment() As LABEL_ALIGN
     Alignment = m_Alignment
End Property ' Get Alignment

Public Property Let Alignment(New_Alignment As LABEL_ALIGN)
     m_Alignment = New_Alignment
     PropertyChanged "Alignment"
     ' Redraw
     DrawCaption
End Property ' Let Alignment

Public Property Get BorderStyle() As CTL_BORDERSTYLE
     BorderStyle = UserControl.BorderStyle
End Property ' Get BorderStyle

Public Property Let BorderStyle(ByVal New_BorderStyle As CTL_BORDERSTYLE)
     UserControl.BorderStyle() = New_BorderStyle
     PropertyChanged "BorderStyle"
     ' Redraw
     DrawCaption
End Property ' Let BorderStyle

Public Property Get BoxBackgroundColor() As OLE_COLOR
     BoxBackgroundColor = m_BoxBackgroundColor
End Property ' Get BoxBackgroundColor

Public Property Let BoxBackgroundColor(New_BoxBackgroundColor As OLE_COLOR)
     m_BoxBackgroundColor = New_BoxBackgroundColor
     PropertyChanged "BoxBackgroundColor"
     DrawCaption
End Property ' Let BoxBackgroundColor

Public Property Get BoxBorderColor() As OLE_COLOR
     BoxBorderColor = m_BoxBorderColor
End Property ' Get BoxBorderColor

Public Property Let BoxBorderColor(New_BoxBorderColor As OLE_COLOR)
     m_BoxBorderColor = New_BoxBorderColor
     PropertyChanged "BoxBorderColor"
     DrawCaption
End Property ' Let BoxBorderColor

Public Property Get BoxBorderWidth() As Long
     BoxBorderWidth = m_BoxBorderWidth
End Property ' Get BoxBorderWidth

Public Property Let BoxBorderWidth(New_BoxBorderWidth As Long)
     m_BoxBorderWidth = New_BoxBorderWidth
     PropertyChanged "BoxBorderWidth"
     DrawCaption
End Property ' Let BoxBorderWidth

Public Property Get BoxStyle() As CTL_BOXSTYLE
     BoxStyle = m_BoxStyle
End Property ' Get BoxStyle

Public Property Let BoxStyle(New_BoxStyle As CTL_BOXSTYLE)
     m_BoxStyle = New_BoxStyle
     PropertyChanged "BoxStyle"
     DrawCaption
End Property ' Let BoxStyle

Public Property Get Caption() As String
     Caption = m_Caption
End Property ' Get Caption

Public Property Let Caption(New_Caption As String)
     m_Caption = New_Caption
     PropertyChanged "Caption"
     ' Redraw
     DrawCaption
End Property ' Let Caption

Public Property Get Enabled() As Boolean
     Enabled = UserControl.Enabled
End Property ' Get Enabled

Public Property Let Enabled(New_Enabled As Boolean)
     UserControl.Enabled() = New_Enabled
     PropertyChanged "Enabled"
End Property ' Let Enabled

Public Property Get Font() As Font
     Set Font = UserControl.Font
End Property ' Get Font

Public Property Set Font(ByVal New_Font As Font)
     Set UserControl.Font = New_Font
     PropertyChanged "Font"
     ' Redraw
     DrawCaption
End Property ' Let Font

Public Property Get ForeColor() As OLE_COLOR
     ForeColor = UserControl.ForeColor
End Property ' Get ForeColor

Public Property Let ForeColor(New_ForeColor As OLE_COLOR)
     UserControl.ForeColor() = New_ForeColor
     PropertyChanged "ForeColor"
     ' Redraw
     DrawCaption
End Property ' Let ForeColor

Public Property Get SelectMarkColor() As OLE_COLOR
     SelectMarkColor = m_SelectMarkColor
End Property ' Get SelectMarkColor

Public Property Let SelectMarkColor(New_SelectMarkColor As OLE_COLOR)
     m_SelectMarkColor = New_SelectMarkColor
     PropertyChanged "SelectMarkColor"
     DrawCaption
End Property ' Let SelectMarkColor

Public Property Get SelectValue() As Boolean
     SelectValue = m_SelectValue
End Property ' Get SelectValue

Public Property Let SelectValue(New_SelectValue As Boolean)
     m_SelectValue = New_SelectValue
     PropertyChanged "SelectValue"
     DrawCaption
End Property ' Let SelectValue

Public Property Get WordWrap() As Boolean
     WordWrap = m_WordWrap
End Property ' Get WordWrap

Public Property Let WordWrap(New_WordWrap As Boolean)
     m_WordWrap = New_WordWrap
     Select Case m_WordWrap
          Case 0
               m_Flag = DT_SINGLELINE
          Case Else
               m_Flag = DT_WORDBREAK
     End Select
     PropertyChanged "WordWrap"
     ' Redraw
     DrawCaption
End Property ' Let WordWrap

'**************************************************************************************************
' UserControl Intrinsic Methods
'**************************************************************************************************
Private Sub UserControl_GotFocus()
     m_bFocus = True
     DrawCaption
End Sub ' UserControl_GotFocus

Private Sub UserControl_InitProperties()
     BoxBorderColor = m_def_BoxBorderColor
     BoxBorderWidth = m_def_BoxBorderWidth
     BoxBackgroundColor = m_def_BoxBackgroundColor
     Caption = UserControl.Extender.Name
     Set UserControl.Font = Ambient.Font
     SelectMarkColor = m_def_SelectMarkColor
     WordWrap = m_def_WordWrap
End Sub ' UserControl_InitProperties

Private Sub UserControl_LostFocus()
     m_bFocus = False
     DrawCaption
End Sub ' UserControl_LostFocus

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     On Error Resume Next
     With PropBag
          Alignment = .ReadProperty("Alignment", m_def_Alignment)
          UserControl.BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
          BoxBorderColor = .ReadProperty("BoxBorderColor", m_def_BoxBorderColor)
          BoxBorderWidth = .ReadProperty("BoxBorderWidth", m_def_BoxBorderWidth)
          BoxBackgroundColor = .ReadProperty("BoxBackgroundColor", Ambient.ForeColor)
          BoxStyle = .ReadProperty("BoxStyle", 0)
          Caption = .ReadProperty("Caption", Extender.Name)
          UserControl.Enabled = .ReadProperty("Enabled", True)
          Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
          UserControl.ForeColor = .ReadProperty("ForeColor", Ambient.ForeColor)
          SelectMarkColor = .ReadProperty("SelectMarkColor", m_def_SelectMarkColor)
          SelectValue = .ReadProperty("SelectValue", m_def_SelectValue)
          WordWrap = .ReadProperty("WordWrap", m_def_WordWrap)
     End With
End Sub ' UserControl_ReadProperties

Private Sub UserControl_Resize()
     DrawCaption
End Sub ' UserControl_Resize

Private Sub UserControl_Show()
     If Ambient.UserMode Then
          Timer.Enabled = True
     End If
End Sub ' UserControl_Show

Private Sub UserControl_Terminate()
     Dim lRtn As Long
     Timer.Enabled = False
End Sub ' UserControl_Terminate

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     With PropBag
          Call .WriteProperty("Alignment", m_Alignment, m_def_Alignment)
          Call .WriteProperty("BorderStyle", UserControl.BorderStyle, m_def_BorderStyle)
          Call .WriteProperty("BoxBorderColor", m_BoxBorderColor, m_def_BoxBorderColor)
          Call .WriteProperty("BoxBorderWidth", m_BoxBorderWidth, m_def_BoxBorderWidth)
          Call .WriteProperty("BoxBackgroundColor", m_BoxBackgroundColor, Ambient.ForeColor)
          Call .WriteProperty("BoxStyle", m_BoxStyle, 0)
          Call .WriteProperty("Caption", m_Caption, Extender.Name)
          Call .WriteProperty("Enabled", UserControl.Enabled, True)
          Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
          Call .WriteProperty("ForeColor", UserControl.ForeColor, Ambient.ForeColor)
          Call .WriteProperty("SelectMarkColor", m_SelectMarkColor, m_def_SelectMarkColor)
          Call .WriteProperty("SelectValue", m_SelectValue, m_def_SelectValue)
          Call .WriteProperty("WordWrap", m_WordWrap, m_def_WordWrap)
     End With
End Sub ' UserControl_WriteProperties

'**************************************************************************************************
' SelectBox Private Methods
'**************************************************************************************************
Private Sub DrawCaption()
     Dim lRtn As Long
     Dim lHt As Long
     Dim lWt As Long
     Dim rc As RECT
     Dim m_FocusRect As RECT
     Cls
     ' move the caption to left by width BOX_GAP
     m_CaptionRect.Left = BOX_GAP
     ' terminate rectangle on the right by using scalewidth
     m_CaptionRect.Right = UserControl.ScaleWidth
     ' terminate rectangle on the bottom by using scaleheight
     m_CaptionRect.Bottom = UserControl.ScaleHeight
     ' call api
     lRtn = DrawTextEx(UserControl.hdc, m_Caption, Len(m_Caption), m_CaptionRect, _
          m_Flag Or m_Alignment, ByVal 0&)
     ' if control has the focus
     If m_bFocus Then
          ' calculate the rectangle size of the text
          lRtn = DrawTextEx(UserControl.hdc, m_Caption, Len(m_Caption), m_FocusRect, _
               DT_CALCRECT, ByVal 0&)
          ' add to the rectangle to give some space around the text
          m_FocusRect.Left = m_FocusRect.Left + (BOX_GAP - 2)
          m_FocusRect.Right = m_FocusRect.Right + (BOX_GAP + 2)
          ' call the api
          DrawFocusRect hdc, m_FocusRect
     End If
     ' draw checkbox
     DrawBox rc
     ' Send text to usercontrol canvas
     UserControl.MaskPicture = UserControl.Image
End Sub ' DrawText

Friend Function DrawBox(rc As RECT)
     Dim lRtn As Long
     Dim lColor As Long
     Dim olColor As Long
     Dim lPt(20) As POINTAPI
     ' Convert outline color
     olColor = TranslateOleColor(m_BoxBorderColor)
     ' Create pen object
     m_hPen = CreatePen(PS_SOLID, m_BoxBorderWidth, olColor)
     'Copy pen onto the dc and store old pen
     m_oldPen = SelectObject(UserControl.hdc, m_hPen)
     ' translate color before initializing LOGBRUSH struct
     lColor = TranslateOleColor(m_BoxBackgroundColor)
     ' Initialize logbrush struct
     With m_LogBrush
          .lbColor = lColor
          .lbStyle = BS_SOLID
          .lbHatch = HS_SOLID
     End With
     ' Create a brush
     m_hBrush = CreateBrushIndirect(m_LogBrush)
     ' Copy brush onto the dc and store old brush
     m_oldBrush = SelectObject(UserControl.hdc, m_hBrush)
     ' Process Box Style
     Select Case m_BoxStyle
          Case 0 ' Disc
               rc.Left = 2
               rc.Top = 0
               rc.Right = 15
               rc.Bottom = 13
               lRtn = Ellipse(UserControl.hdc, rc.Left, rc.Top, rc.Right, rc.Bottom)
          Case 1 ' Square
               rc.Left = 3
               rc.Top = 1
               rc.Right = 15
               rc.Bottom = 13
               lRtn = Rectangle(UserControl.hdc, rc.Left, rc.Top, rc.Right, rc.Bottom)
          Case 2 ' Diamond
               lPt(0).X = 9: lPt(0).Y = 0
               lPt(1).X = 2: lPt(1).Y = 7
               lPt(2).X = 9: lPt(2).Y = 14
               lPt(3).X = 16: lPt(3).Y = 7
               lRtn = Polygon(UserControl.hdc, lPt(0), 4)
     End Select
     ' Delete brush and restore old
     If m_oldBrush Then
          lRtn = SelectObject(UserControl.hdc, m_oldBrush)
          DeleteObject (lRtn)
     End If
     ' Delete pen and restore old
     If m_oldPen Then
          lRtn = SelectObject(UserControl.hdc, m_oldPen)
          DeleteObject (lRtn)
     End If
     ' Draw selectmark
     If m_SelectValue Then DrawSelectMark
End Function ' DrawBox

Private Sub DrawSelectMark()
     Dim pt As POINTAPI
     Dim lRtn As Long
     Dim lPen As Long
     Dim lPenColor As Long
     Dim lOldPen As Long
     Dim rc As RECT
     With UserControl
          ' Convert selectmark color
          lPenColor = TranslateOleColor(m_SelectMarkColor)
          ' Create Pen
          lPen = CreatePen(PS_SOLID, 2, lPenColor)
          'Copy pen onto the dc and store old pen
          lOldPen = SelectObject(.hdc, lPen)
          ' draw selectmark based on boxstyle
          Select Case m_BoxStyle
               Case 0, 1
                    rc.Left = 5
                    rc.Top = 5
                    rc.Right = 9
                    rc.Bottom = 9
                    MoveToEx UserControl.hdc, rc.Left, rc.Top, pt
                    LineTo UserControl.hdc, rc.Right, rc.Bottom
                    rc.Left = 16
                    rc.Top = 0
                    rc.Right = 8
                    rc.Bottom = 8
                    MoveToEx UserControl.hdc, rc.Right, rc.Bottom, pt
                    LineTo UserControl.hdc, rc.Left, rc.Top
               Case 2
                    rc.Left = 6
                    rc.Top = 6
                    rc.Right = 10
                    rc.Bottom = 10
                    MoveToEx UserControl.hdc, rc.Left, rc.Top, pt
                    LineTo UserControl.hdc, rc.Right, rc.Bottom
                    rc.Left = 18
                    rc.Top = 0
                    rc.Right = 9
                    rc.Bottom = 9
                    MoveToEx UserControl.hdc, rc.Right, rc.Bottom, pt
                    LineTo UserControl.hdc, rc.Left, rc.Top
          End Select
          ' Delete pen and restore old
          If lOldPen Then
               lRtn = SelectObject(UserControl.hdc, lOldPen)
               DeleteObject (lRtn)
          End If
     End With
End Sub ' DrawSelectMark

Private Function TranslateOleColor(ByVal lColor As OLE_COLOR) As Long
     Const cHighBitMask = &H80000000
     Dim lRslt As Long
     If lColor And cHighBitMask Then
          ' convert color
          lRslt = lColor And Not cHighBitMask
          lRslt = GetSysColor(lRslt)
     Else
          ' otherwise, use original color
          lRslt = lColor
     End If
     ' Return function
     TranslateOleColor = lRslt
End Function ' TranslateOleColor

Private Sub Timer_Timer()
     Dim r As RECT
     Dim pt As POINTAPI
     Dim lRtn As Long
     Dim X As Long
     Dim Y As Long
     ' Get the position of our window rectangle.  I used GetWindowRect
     ' instead of GetClientRect because I wanted the values in Screen
     ' coordinates.  Using GetClientRect returns position in relation
     ' to the parent window requiring me to convert the position relative
     ' to the screen....
     lRtn = GetWindowRect(UserControl.hWnd, r)
     ' Get the cursor position on the screen
     GetCursorPos pt
     ' Determine if point coordinates within our rectangle
     lRtn = PtInRect(r, pt.X, pt.Y)
     ' If yes
     If lRtn <> False Then
          'Since we're in our rectangle, let's convert screen coordinates
          ' to client coordinates to raise some events
          lRtn = ScreenToClient(UserControl.hWnd, pt)
          ' if succeeds
          If lRtn Then
               X = pt.X * Screen.TwipsPerPixelX
               Y = pt.Y * Screen.TwipsPerPixelX
          End If
          ' Generate a mouse_move event
          RaiseEvent MouseMove(X, Y)
          ' poll for detection of left mouse click.  This function returns
          ' a negative value while the button is being depressed...
          lRtn = GetKeyState(vbKeyLButton)
          ' If a negative value then the button is currently depressed.
          If lRtn < False Then
               ' Mouse is down on our control so set focus
               UserControl.SetFocus
               ' Button is down so raise mousedown event
               RaiseEvent MouseDown(X, Y)
               While lRtn < 0
                    ' Poll again until get a 1 or 0.  1 means the key
                    ' is in  the toggled state.  0 if not toggled.
                    lRtn = GetKeyState(vbKeyLButton)
                    ' do this or die inside this loop.
                    DoEvents
               Wend
               ' Set select value
               If SelectValue = False Then
                    SelectValue = True
               Else
                    SelectValue = False
               End If
               ' Voila....raise a click.
               RaiseEvent Click
               ' After our click event, let's raise the mouseup event
               RaiseEvent MouseUp(X, Y)
          End If
     End If
     ' If we have focus
     If m_bFocus Then
          ' Test for space bar so we can select like a normal checkbox
          lRtn = GetKeyState(vbKeySpace)
          If lRtn < False Then
               ' raise a key down event
               ' spacebar is down
               While lRtn < 0
                    ' Poll again until get a 1 or 0.  1 means the key
                    ' is in  the toggled state.  0 if not toggled.
                    lRtn = GetKeyState(vbKeySpace)
                    ' do this or die inside this loop.
                    DoEvents
               Wend
               ' Set select value
               If SelectValue = False Then
                    SelectValue = True
               Else
                    SelectValue = False
               End If
          End If
     End If
End Sub ' Timer_Timer

VERSION 5.00
Begin VB.UserControl TextBoxButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3435
   ScaleHeight     =   1590
   ScaleWidth      =   3435
   ToolboxBitmap   =   "TextBoxButton.ctx":0000
   Begin VB.TextBox txtText 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "TextBoxButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Enums
Public Enum tbbAppearenceConsts
    [Flat] = 0
    [3D] = 1
End Enum
Public Enum tbbBorderStyleConsts
    [None] = 0
    [Fixed Single] = vbFixedSingle
End Enum
Public Enum tbbButtonValueConstants
    [Up]
    [Down]
End Enum
'Types:
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'API Constants - Draw Edge:
Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_MIDDLE = &H800
'API Declarations:
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
'Property Variables:
Dim m_Value As tbbButtonValueConstants
Dim m_ButtonWidth As Single
Dim m_Picture As Picture

Private rctButton As RECT
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Event Change() 'MappingInfo=txtText,txtText,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event ButtonClick()
Attribute ButtonClick.VB_Description = "Occurs when the button is left-clicked."
Attribute ButtonClick.VB_MemberFlags = "200"
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtText,txtText,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtText,txtText,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtText,txtText,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
'Default Property Values:
Const m_def_Value = Up
Const m_def_ButtonWidth = 300

'API Types:
Private Type POINTAPI 'Type for holding X & Y co-ordinates
    X As Long
    Y As Long
End Type
'API Declarations:
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long 'API for finding the hWnd of the window under the cursor
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long                      'Finds the cursor's co-ordinates


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,Alignment
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Misc"
    Alignment = txtText.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    txtText.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As tbbAppearenceConsts
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As tbbAppearenceConsts)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
    Call DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Call DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As tbbBorderStyleConsts
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As tbbBorderStyleConsts)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    Call DrawControl
End Property

Private Sub txtText_Change()
    RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    Call DrawControl
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = txtText.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtText.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = txtText.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtText.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = UserControl.hDC
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
Attribute Locked.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Locked = txtText.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtText.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
Attribute MaxLength.VB_ProcData.VB_Invoke_Property = ";Behavior"
    MaxLength = txtText.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtText.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
Attribute MultiLine.VB_ProcData.VB_Invoke_Property = ";Behavior"
    MultiLine = txtText.MultiLine
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    PropertyChanged "Picture"
    Call DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
    UserControl.Refresh
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton And IsButtonActive(VK_LBUTTON, UserControl.hWnd) And _
        X >= (rctButton.Left * TwipsX) And X <= (rctButton.Right * TwipsX) And _
        Y >= (rctButton.Top * TwipsY) And Y <= (rctButton.Bottom * TwipsY) Then
        m_Value = Down
        Call DrawControl
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbLeftButton) And IsButtonActive(VK_LBUTTON, UserControl.hWnd) And IsHot(X, Y) Then
        m_Value = Down
        Call DrawControl
    ElseIf (Button And vbLeftButton) Or (m_Value = Down And _
        Not IsButtonActive(VK_LBUTTON, UserControl.hWnd)) And Not IsHot(X, Y) Then
        m_Value = Up
        Call DrawControl
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbLeftButton) Or (m_Value = Down And _
        Not IsButtonActive(VK_LBUTTON, UserControl.hWnd)) And Not IsHot(X, Y) Then
        m_Value = Up
        Call DrawControl
        If Button = vbLeftButton And IsHot(X, Y) Then RaiseEvent ButtonClick
    End If
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    On Local Error Resume Next
    Call SetButtonRECT
    Call DrawControl
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
Attribute SelLength.VB_ProcData.VB_Invoke_Property = ";Text"
    SelLength = txtText.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtText.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,ScrollBars
Public Property Get ScrollBars() As Integer
Attribute ScrollBars.VB_Description = "Returns/sets a value indicating whether an object has vertical or horizontal scroll bars."
Attribute ScrollBars.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ScrollBars = txtText.ScrollBars
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
Attribute SelStart.VB_ProcData.VB_Invoke_Property = ";Text"
    SelStart = txtText.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtText.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
Attribute SelText.VB_ProcData.VB_Invoke_Property = ";Text"
    SelText = txtText.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtText.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtText,txtText,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
Attribute Text.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Text.VB_UserMemId = -517
    Text = txtText.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtText.Text() = New_Text
    PropertyChanged "Text"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set m_Picture = LoadPicture("")
    m_ButtonWidth = m_def_ButtonWidth
    m_Value = m_def_Value
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtText.Alignment = PropBag.ReadProperty("Alignment", vbLeftJustify)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Set txtText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtText.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtText.Locked = PropBag.ReadProperty("Locked", False)
    txtText.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
    txtText.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtText.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtText.SelText = PropBag.ReadProperty("SelText", "")
    txtText.Text = PropBag.ReadProperty("Text", "")
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_ButtonWidth = PropBag.ReadProperty("ButtonWidth", m_def_ButtonWidth)
    Call DrawControl
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Alignment", txtText.Alignment, vbLeftJustify)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Font", txtText.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", txtText.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Locked", txtText.Locked, False)
    Call PropBag.WriteProperty("MaxLength", txtText.MaxLength, 0)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("SelLength", txtText.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtText.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtText.SelText, "")
    Call PropBag.WriteProperty("Text", txtText.Text, "")
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("ButtonWidth", m_ButtonWidth, m_def_ButtonWidth)
End Sub

Private Sub DrawControl()
    On Local Error Resume Next
    Dim sngImageWidth As Single, sngImageHeight As Single
    If (m_Picture Is Nothing) = False Then
    sngImageWidth = m_Picture.Width \ 2
    sngImageHeight = m_Picture.Height \ 2
    End If
    With UserControl
        'Clear past drawings
        Call .Cls
        Call SetButtonRECT
        'Set the RECT of the Button
        Dim rctImage As RECT
        rctImage.Left = (((rctButton.Right * TwipsX) - (m_ButtonWidth \ 2))) - (sngImageWidth \ 2) - _
            IIf(m_Value = Up, TwipsX, 0)
        rctImage.Right = rctImage.Left + (sngImageWidth \ 2)
        rctImage.Top = (((rctButton.Bottom - rctButton.Top) * TwipsY) \ 2) - (sngImageHeight \ 2) - _
            IIf(m_Value = Up, TwipsY, 0)
        rctImage.Bottom = rctImage.Top + (sngImageHeight \ 2)
        
        'Draw the border with the middle bit
        Call DrawEdge(.hDC, rctButton, IIf(m_Value = Down, EDGE_SUNKEN, EDGE_RAISED), BF_RECT Or BF_MIDDLE)
        'If we have a picture draw it
        If (m_Picture Is Nothing) = False Then
            Call .PaintPicture(m_Picture, rctImage.Left, rctImage.Top, , , , , , , vbSrcCopy)
            Call .PaintPicture(m_Picture, rctImage.Left, rctImage.Top, , , , , , , vbSrcAnd)
        End If
        'Draw the border without the middle bit - covers up the picture if too big
        Call DrawEdge(.hDC, rctButton, IIf(m_Value = Down, EDGE_SUNKEN, EDGE_RAISED), BF_RECT)
        'Show changes
        If .AutoRedraw Then Call .Refresh
    End With
End Sub

Private Function IsHot(ByVal X As Single, ByVal Y As Single) As Boolean
    On Local Error Resume Next
    Dim CursorPosition As POINTAPI 'Variable for cursor's X & Y values

    'Get the Cursor position
    Call GetCursorPos(CursorPosition)
    With CursorPosition
        IsHot = WindowFromPoint(.X, .Y) = hWnd And _
            X >= (rctButton.Left * TwipsX) And X <= (rctButton.Right * TwipsX) And _
            Y >= (rctButton.Top * TwipsY) And Y <= (rctButton.Bottom * TwipsY) 'Return whether the object is hot
    End With
End Function

Private Sub SetButtonRECT()
    On Local Error Resume Next
    With UserControl
        Call SetRect(rctButton, (.ScaleWidth - m_ButtonWidth) \ TwipsX, _
            0, .ScaleWidth \ TwipsY, .ScaleHeight \ TwipsY)
        txtText.Move TwipsY(1), TwipsX(2), .ScaleWidth - m_ButtonWidth - TwipsX(2), _
            .ScaleHeight - TwipsY(4)
    End With
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,300
Public Property Get ButtonWidth() As Single
Attribute ButtonWidth.VB_Description = "Returns/sets the width of the button."
Attribute ButtonWidth.VB_ProcData.VB_Invoke_Property = ";Scale"
    ButtonWidth = m_ButtonWidth
End Property

Public Property Let ButtonWidth(ByVal New_ButtonWidth As Single)
    m_ButtonWidth = New_ButtonWidth
    PropertyChanged "ButtonWidth"
    Call DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,1,2,Up
Public Property Get Value() As tbbButtonValueConstants
Attribute Value.VB_Description = "Returns the value of the button, i.e. Up or Down."
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Value.VB_MemberFlags = "400"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As tbbButtonValueConstants)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_Value = New_Value
    PropertyChanged "Value"
    Call DrawControl
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub


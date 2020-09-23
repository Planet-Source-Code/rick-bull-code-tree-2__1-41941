VERSION 5.00
Begin VB.UserControl Hyperlink 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5010
   MouseIcon       =   "Hyperlink.ctx":0000
   MousePointer    =   99  'Custom
   PropertyPages   =   "Hyperlink.ctx":030A
   ScaleHeight     =   1065
   ScaleWidth      =   5010
   ToolboxBitmap   =   "Hyperlink.ctx":033F
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPopupCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPopupCopyAddress 
         Caption         =   "Copy &Address"
      End
   End
End
Attribute VB_Name = "Hyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'Enumerations
Public Enum lnkAppearenceConsts
    [Flat] = 0
    [3D] = 1
End Enum
Public Enum lnkAddressTypeConsts
    [Shell Execute]
    [Call By Name]
End Enum
Public Enum lnkBorderStyleConsts
    [None] = 0
    [Fixed Single] = vbFixedSingle
End Enum
Public Enum lnkMousePointerConsts
    [Default] = vbDefault
    [Arrow] = vbArrow
    [Cross] = vbCrosshair
    [I-Beam] = vbIbeam
    [Icon] = vbIconPointer
    [Size] = vbSizePointer
    [Size NE SW] = vbSizeNESW
    [Size NS] = vbSizeNS
    [Size NW SE] = vbSizeNWSE
    [Size WE] = vbSizeWE
    [Up Arrow] = vbUpArrow
    [Hourglass] = vbHourglass
    [No Drop] = vbNoDrop
    [Arror and Hourglass] = vbArrowHourglass
    [Arrow and Question] = vbArrowQuestion
    [Size All] = vbSizeAll
    [Custom] = vbCustom
End Enum
Public Enum lnkOLEDropModeConsts
    [None] = vbOLEDropNone
    [Manual] = vbOLEDropManual
End Enum
'Default Property Values:
Const m_def_AddressType = [Shell Execute]
Const m_def_WindowStyle = vbNormalFocus
Const m_def_CanGetFocus = False
Const m_def_WordWrap = False
Const m_def_Alignment = vbLeftJustify
Const m_def_AutoSize = False
Const m_def_HotColor = vbBlue
Const m_def_Address = "http://www.rickmusic.co.uk/"
Const m_def_Caption = "www.RickMusic.co.uk/"
'Property Variables:
Dim m_AddressType As lnkAddressTypeConsts
Dim m_HotFont As Font
Dim m_WindowStyle As VbAppWinStyle
Dim m_CanGetFocus As Boolean
Dim m_WordWrap As Boolean
Dim m_Alignment As AlignmentConstants
Dim m_AutoSize As Boolean
Dim m_HotColor As OLE_COLOR
Dim m_Address As String
Dim m_Caption As String
'Event Declarations:
Event MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseOut.VB_Description = "Occurs when the cursor leaves the control."
Event MouseOver(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseOver.VB_Description = "Occurs when the cursor enters over the control."
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Attribute Hide.VB_Description = "Occurs when the control's Visible property changes to False."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Event OLECompleteDrag(Effect As Long) 'MappingInfo=UserControl,UserControl,-1,OLECompleteDrag
Attribute OLECompleteDrag.VB_Description = "Occurs at the OLE drag/drop source control after a manual or automatic drag/drop has been completed or canceled."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean) 'MappingInfo=UserControl,UserControl,-1,OLEGiveFeedback
Attribute OLEGiveFeedback.VB_Description = "Occurs at the source control of an OLE drag/drop operation when the mouse cursor needs to be changed."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Occurs when any part of a form or PictureBox control is moved, enlarged, or exposed."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show
Attribute Show.VB_Description = "Occurs when the control's Visible property changes to True."

'API Types:
Private Type POINTAPI 'Type for holding X & Y co-ordinates
    X As Long
    Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'API Declarations:
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long 'API for finding the hWnd of the window under the cursor
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long                      'Finds the cursor's co-ordinates
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long

'Variables:
Private bolHasCapture As Boolean
Private bolHasFocus As Boolean

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As lnkAppearenceConsts
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As lnkAppearenceConsts)
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
Public Property Get BorderStyle() As lnkBorderStyleConsts
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As lnkBorderStyleConsts)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
    Call DrawControl
End Property

Private Sub mnuPopupCopy_Click()
    On Local Error Resume Next
    Call Clipboard.SetText(m_Caption, vbCFText)
    Call Clipboard.SetText(m_Caption, vbCFRTF)
End Sub

Private Sub mnuPopupCopyAddress_Click()
    On Local Error Resume Next
    Call Clipboard.SetText(m_Address, vbCFText)
    Call Clipboard.SetText(m_Address, vbCFRTF)
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Call DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Call DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
    Call DrawControl
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

Private Sub UserControl_GotFocus()
    bolHasFocus = True
    If m_CanGetFocus Then Call DrawControl
End Sub

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub

Private Sub UserControl_Initialize()
    bolHasCapture = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii = vbKeyReturn Then Call Clicked
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_LostFocus()
    bolHasFocus = False
    Call DrawControl
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = vbRightButton Then Call PopupMenu(mnuPopup)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Misc"
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    On Local Error Resume Next
    Call DrawControl
    If Button = 0 And IsHot And Ambient.UserMode And bolHasCapture = False Then
        Call ReleaseCapture
        Call SetCapture(UserControl.hWnd)
        bolHasCapture = True
        RaiseEvent MouseOver(Button, Shift, X, Y)
    ElseIf Button = 0 And IsHot = False And Ambient.UserMode And bolHasCapture Then
        Call ReleaseCapture
        bolHasCapture = False
        RaiseEvent MouseOut(Button, Shift, X, Y)
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As lnkMousePointerConsts
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Misc"
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As lnkMousePointerConsts)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bolClicked As Boolean
    bolClicked = Button = vbLeftButton And IsHot
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If bolClicked Then Call Clicked
    On Local Error Resume Next
    Call DrawControl
    If IsHot And Ambient.UserMode Then
        Call ReleaseCapture
        Call SetCapture(UserControl.hWnd)
        bolHasCapture = True
        RaiseEvent MouseOver(Button, Shift, X, Y)
    ElseIf IsHot = False And Ambient.UserMode And bolHasCapture Then
        Call ReleaseCapture
        bolHasCapture = False
        RaiseEvent MouseOut(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    UserControl.OLEDrag
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As lnkOLEDropModeConsts
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As lnkOLEDropModeConsts)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
Attribute Refresh.VB_UserMemId = -550
    UserControl.Refresh
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    Call DrawControl
End Sub

Private Sub UserControl_Show()
    RaiseEvent Show
    Call DrawControl
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextHeight
Public Function TextHeight(ByVal Str As String) As Single
Attribute TextHeight.VB_Description = "Returns the height of a text string as it would be printed in the current font."
    TextHeight = UserControl.TextHeight(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextWidth
Public Function TextWidth(ByVal Str As String) As Single
Attribute TextWidth.VB_Description = "Returns the width of a text string as it would be printed in the current font."
    TextWidth = UserControl.TextWidth(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbBlue
Public Property Get HotColor() As OLE_COLOR
Attribute HotColor.VB_Description = "Returns/set the color that the text will be when the cursor is over the control."
Attribute HotColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    HotColor = m_HotColor
End Property

Public Property Let HotColor(ByVal New_HotColor As OLE_COLOR)
    m_HotColor = New_HotColor
    PropertyChanged "HotColor"
    If IsHot And Ambient.UserMode Then Call DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,http://www.rickmusic.co.uk/
Public Property Get Address() As String
Attribute Address.VB_Description = "Returns/sets the URI for the document that will be opened when the control is clicked."
Attribute Address.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Address = m_Address
End Property

Public Property Let Address(ByVal New_Address As String)
    m_Address = New_Address
    PropertyChanged "Address"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,www.RickMusic.co.uk/
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Call DrawControl
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_HotColor = m_def_HotColor
    m_Address = m_def_Address
    m_Caption = m_def_Caption
    m_WordWrap = m_def_WordWrap
    m_Alignment = m_def_Alignment
    m_AutoSize = m_def_AutoSize
    m_CanGetFocus = m_def_CanGetFocus
    m_WindowStyle = m_def_WindowStyle
    Set m_HotFont = Ambient.Font
    m_AddressType = m_def_AddressType
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", UserControl.MouseIcon)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", vbCustom)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    m_HotColor = PropBag.ReadProperty("HotColor", m_def_HotColor)
    m_Address = PropBag.ReadProperty("Address", m_def_Address)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_WordWrap = PropBag.ReadProperty("WordWrap", m_def_WordWrap)
    m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    m_CanGetFocus = PropBag.ReadProperty("CanGetFocus", m_def_CanGetFocus)
    m_WindowStyle = PropBag.ReadProperty("WindowStyle", m_def_WindowStyle)
    Set m_HotFont = PropBag.ReadProperty("HotFont", Ambient.Font)
    m_AddressType = PropBag.ReadProperty("AddressType", m_def_AddressType)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, UserControl.MouseIcon)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, vbCustom)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("HotColor", m_HotColor, m_def_HotColor)
    Call PropBag.WriteProperty("Address", m_Address, m_def_Address)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("WordWrap", m_WordWrap, m_def_WordWrap)
    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("CanGetFocus", m_CanGetFocus, m_def_CanGetFocus)
    Call PropBag.WriteProperty("WindowStyle", m_WindowStyle, m_def_WindowStyle)
    Call PropBag.WriteProperty("HotFont", m_HotFont, Ambient.Font)
    Call PropBag.WriteProperty("AddressType", m_AddressType, m_def_AddressType)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "Returns/sets a value that determines whether a control expands to fit the text in it's Caption."
Attribute WordWrap.VB_ProcData.VB_Invoke_Property = ";Misc"
    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    m_WordWrap = New_WordWrap
    PropertyChanged "WordWrap"
    Call DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text"
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Misc"
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    m_Alignment = New_Alignment
    PropertyChanged "Alignment"
    Call DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to fit it's entire contents."
Attribute AutoSize.VB_ProcData.VB_Invoke_Property = ";Position"
Attribute AutoSize.VB_UserMemId = -500
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    Call DrawControl
End Property

Private Function WrapText(Text As String, Optional ObjectWidth As Single = -1) As String
    On Local Error Resume Next
    'Set the width to the width of the parent if missing
    If ObjectWidth = -1 Then ObjectWidth = ScaleWidth
    
    'Split the text by spaces
    Dim SplitText() As String
    SplitText() = Split(Text, Space(1), , vbTextCompare)
    
    'Return Value = the first word
    WrapText = SplitText(LBound(SplitText))
    
    'Loop for all words minus the first
    Dim LoopCounter As Integer
    For LoopCounter = LBound(SplitText) + 1 To UBound(SplitText)
        'Add to the return value a new line if the text is bigger than the width, _
         and a space if it isn't - this avoids text being indented from the very left
        WrapText = WrapText + IIf(Me.TextWidth(WrapText + Space(1) + SplitText(LoopCounter)) > _
            ObjectWidth, vbNewLine, Space(1)) + SplitText(LoopCounter)
    Next LoopCounter
End Function

Public Function IsHot() As Boolean
Attribute IsHot.VB_Description = "Returns whether the cursor is over the control."
    On Local Error Resume Next
    Dim CursorPosition As POINTAPI 'Variable for cursor's X & Y values

    'Get the Cursor position
    Call GetCursorPos(CursorPosition)
    With CursorPosition
        IsHot = WindowFromPoint(.X, .Y) = hWnd 'Return whether the object is hot
    End With
End Function

Private Sub DrawControl()
    On Local Error Resume Next
    With UserControl
        'Remove the past drawings
        .Cls
        
        'Get the current for colour
        Dim lngForeColor As Long
        lngForeColor = .ForeColor
        Dim fntOldFont As Font
        Set fntOldFont = UserControl.Font
        
        'Set the forecolor to hot if necassary
        If UserControl.Enabled = False Then
            .ForeColor = vbGrayText
        ElseIf IsHot And .Ambient.UserMode Then
            .ForeColor = m_HotColor
            Set .Font = m_HotFont
        End If
    
        If m_WordWrap Then
            'Get the printable (wrapped if wanted) text
            Dim strPrintableCaption() As String
            strPrintableCaption = Split(WrapText(m_Caption, .ScaleWidth - TwipsX(2)), vbNewLine, , vbTextCompare)
        
            If m_AutoSize = False Then
                'Loop for all texts
                Dim intLoopCounter As Integer
                Dim sngCaptionHeight As Single
                For intLoopCounter = LBound(strPrintableCaption) To UBound(strPrintableCaption)
                    sngCaptionHeight = sngCaptionHeight + .TextHeight(strPrintableCaption(intLoopCounter))
                Next intLoopCounter
                .CurrentY = (.ScaleHeight \ 2) - (sngCaptionHeight \ 2)
            End If
            For intLoopCounter = LBound(strPrintableCaption) To UBound(strPrintableCaption)
                Select Case m_Alignment
                    'Set the X pos depending on the Alignment
                    Case vbLeftJustify
                        .CurrentX = TwipsX
                    Case vbCenter
                        .CurrentX = (.ScaleWidth / 2) - (.TextWidth(strPrintableCaption(intLoopCounter)) / 2)
                    Case vbRightJustify
                        .CurrentX = .ScaleWidth - .TextWidth(strPrintableCaption(intLoopCounter)) - TwipsX
                End Select
                
                'Print the caption
                Print strPrintableCaption(intLoopCounter)
            Next intLoopCounter
            
        Else
            Select Case m_Alignment
                'Set the X pos depending on the Alignment
                Case vbLeftJustify
                    .CurrentX = TwipsX
                Case vbCenter
                    .CurrentX = (.ScaleWidth \ 2) - (.TextWidth(m_Caption) \ 2)
                Case vbRightJustify
                    .CurrentX = .ScaleWidth - .TextWidth(m_Caption) - TwipsX
            End Select
            .CurrentY = (.ScaleHeight \ 2) - (.TextHeight(m_Caption) \ 2) - TwipsY
            'Print the caption
            Print m_Caption
        End If
    
        'Size the control to fit all text
        Call SizeControl
        'Reset the old fore color
        .ForeColor = lngForeColor
        Set UserControl.Font = fntOldFont
        
        If m_CanGetFocus And bolHasFocus Then
            Dim rctControlPos As RECT
            Call GetClientRect(UserControl.hWnd, rctControlPos)
            Call DrawFocusRect(UserControl.hDC, rctControlPos)
            If UserControl.AutoRedraw Then UserControl.Refresh
        End If
        
        'Show the changes
        If .AutoRedraw Then .Refresh
    End With
End Sub

Private Sub SizeControl()
    On Local Error Resume Next
    If m_AutoSize = False Then Exit Sub
    With UserControl
        If m_WordWrap Then
            Dim strPrintableText() As String
            strPrintableText() = Split(WrapText(m_Caption, .ScaleWidth - TwipsX(2)), vbNewLine, , vbTextCompare)
            Dim intLoopCounter As Integer
            Dim sngControlWidth As Single, sngControlHeight As Single
            For intLoopCounter = LBound(strPrintableText) To UBound(strPrintableText)
                If sngControlWidth < .TextWidth(strPrintableText(intLoopCounter)) + TwipsX(2) Then _
                    sngControlWidth = .TextWidth(strPrintableText(intLoopCounter)) + TwipsX(2)
                sngControlHeight = sngControlHeight + .TextHeight(m_Caption)
            Next intLoopCounter
            sngControlHeight = sngControlHeight + TwipsY(2)
        Else
            sngControlWidth = .TextWidth(m_Caption) + TwipsX(2)
            sngControlHeight = .TextHeight(m_Caption) + TwipsX(2)
        End If
        .Height = sngControlHeight
        .Width = sngControlWidth
    End With
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get CanGetFocus() As Boolean
Attribute CanGetFocus.VB_Description = "Returns or sets a value determining if a control itself can receive focus."
Attribute CanGetFocus.VB_ProcData.VB_Invoke_Property = ";Behavior"
    CanGetFocus = m_CanGetFocus
End Property

Public Property Let CanGetFocus(ByVal New_CanGetFocus As Boolean)
    m_CanGetFocus = New_CanGetFocus
    PropertyChanged "CanGetFocus"
    Call DrawControl
End Property

Private Sub Clicked()
    RaiseEvent Click
    On Local Error Resume Next
    If m_AddressType = [Shell Execute] And m_Address <> vbNullString Then
        Call ShellExecute(UserControl.hWnd, vbNullString, m_Address, _
            vbNullString, vbNullString, m_WindowStyle)
    ElseIf m_AddressType = [Call By Name] And m_Address <> vbNullString Then
        Call CallByName(UserControl.Parent, m_Address, VbMethod)
    End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,vbNormalFocus
Public Property Get WindowStyle() As VbAppWinStyle
Attribute WindowStyle.VB_Description = "Returns/sets the style of the window that gets opened when the control is clicked."
Attribute WindowStyle.VB_ProcData.VB_Invoke_Property = ";Behavior"
    WindowStyle = m_WindowStyle
End Property

Public Property Let WindowStyle(ByVal New_WindowStyle As VbAppWinStyle)
    m_WindowStyle = New_WindowStyle
    PropertyChanged "WindowStyle"
End Property

Public Sub About()
Attribute About.VB_Description = "Shows the about dialog."
Attribute About.VB_UserMemId = -552
    On Local Error Resume Next
    Call MsgBox("Hyperlink control written by Rick Bull on 16 April 2002", _
        vbInformation Or vbOKOnly, "About")
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    Call DrawControl
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get HotFont() As Font
Attribute HotFont.VB_Description = "Returns/sets the Font used when the cursor is over the control."
    Set HotFont = m_HotFont
End Property

Public Property Set HotFont(ByVal New_HotFont As Font)
    Set m_HotFont = New_HotFont
    PropertyChanged "HotFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,[Shell Execute]
Public Property Get AddressType() As lnkAddressTypeConsts
Attribute AddressType.VB_Description = "Returns/sets the type that the address is. If set to Shell Execute the address value will be executed when click. If the address type is set to Call By Name the method in the address type will be called."
    AddressType = m_AddressType
End Property

Public Property Let AddressType(ByVal New_AddressType As lnkAddressTypeConsts)
    m_AddressType = New_AddressType
    PropertyChanged "AddressType"
End Property


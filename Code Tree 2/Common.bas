Attribute VB_Name = "Common"
Option Explicit
'Enumerations:
'RTF Function Consts:
Public Enum TextMessages
    WM_CUT = &H300
    WM_COPY = &H301
    WM_PASTE = &H302
    WM_CLEAR = &H303
    WM_UNDO = &H304
End Enum
Public Enum VKButtons
    ' Virtual Keys, Standard Set
    VK_LBUTTON = &H1
    VK_RBUTTON = &H2
    VK_CANCEL = &H3
    VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON

    VK_BACK = &H8
    VK_TAB = &H9

    VK_CLEAR = &HC
    VK_RETURN = &HD

    VK_SHIFT = &H10
    VK_CONTROL = &H11
    VK_MENU = &H12
    VK_PAUSE = &H13
    VK_CAPITAL = &H14

    VK_ESCAPE = &H1B

    VK_SPACE = &H20
    VK_PRIOR = &H21
    VK_NEXT = &H22
    VK_END = &H23
    VK_HOME = &H24
    VK_LEFT = &H25
    VK_UP = &H26
    VK_RIGHT = &H27
    VK_DOWN = &H28
    VK_SELECT = &H29
    VK_PRINT = &H2A
    VK_EXECUTE = &H2B
    VK_SNAPSHOT = &H2C
    VK_INSERT = &H2D
    VK_DELETE = &H2E
    VK_HELP = &H2F

' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'

    VK_NUMPAD0 = &H60
    VK_NUMPAD1 = &H61
    VK_NUMPAD2 = &H62
    VK_NUMPAD3 = &H63
    VK_NUMPAD4 = &H64
    VK_NUMPAD5 = &H65
    VK_NUMPAD6 = &H66
    VK_NUMPAD7 = &H67
    VK_NUMPAD8 = &H68
    VK_NUMPAD9 = &H69
    VK_MULTIPLY = &H6A
    VK_ADD = &H6B
    VK_SEPARATOR = &H6C
    VK_SUBTRACT = &H6D
    VK_DECIMAL = &H6E
    VK_DIVIDE = &H6F
    VK_F1 = &H70
    VK_F2 = &H71
    VK_F3 = &H72
    VK_F4 = &H73
    VK_F5 = &H74
    VK_F6 = &H75
    VK_F7 = &H76
    VK_F8 = &H77
    VK_F9 = &H78
    VK_F10 = &H79
    VK_F11 = &H7A
    VK_F12 = &H7B
    VK_F13 = &H7C
    VK_F14 = &H7D
    VK_F15 = &H7E
    VK_F16 = &H7F
    VK_F17 = &H80
    VK_F18 = &H81
    VK_F19 = &H82
    VK_F20 = &H83
    VK_F21 = &H84
    VK_F22 = &H85
    VK_F23 = &H86
    VK_F24 = &H87

    VK_NUMLOCK = &H90
    VK_SCROLL = &H91

'
'   VK_L VK_R - left and right Alt, Ctrl and Shift virtual keys.
'   Used only as parameters to GetAsyncKeyState() and GetKeyState().
'   No other API or message will distinguish left and right keys in this way.
'  /
    VK_LSHIFT = &HA0
    VK_RSHIFT = &HA1
    VK_LCONTROL = &HA2
    VK_RCONTROL = &HA3
    VK_LMENU = &HA4
    VK_RMENU = &HA5

    VK_ATTN = &HF6
    VK_CRSEL = &HF7
    VK_EXSEL = &HF8
    VK_EREOF = &HF9
    VK_PLAY = &HFA
    VK_ZOOM = &HFB
    VK_NONAME = &HFC
    VK_PA1 = &HFD
    VK_OEM_CLEAR = &HFE
End Enum

'Types:
Private Type KeyboardBytes
     kbByte(0 To 255) As Byte
End Type
Public Type typFindDetails
    Start As Long
    SearchString As String
    LastPosition As Long
End Type

'API Declarations:
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer 'Returns -127 or -128 if the specified key is down, and 0 or 1 of not (alternates each time key is pressed)
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long 'API for making a form on top or not
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As KeyboardBytes) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As KeyboardBytes) As Long

'Constants
'OnTop consts:
'Button Messages (BM)
Private Const BM_SETSTYLE = &HF4
'Button Styles (BS)
Private Const BS_PUSHBUTTON = &H0&
Private Const BS_USERBUTTON = &H8&
Public Const WM_LBUTTONUP = &H202 'Left mouse button up
Private Const HWND_TOPMOST As Long = -1 'Constant for making a form stay on top
Private Const HWND_NOTOPMOST As Long = -2 'Constant for making a form not on top
Private Const SWP_NOMOVE As Long = &H2 'Flags for Always On Top
Private Const SWP_NOSIZE As Long = &H1


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Style Consts
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
'Window pos consts
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOZORDER = &H4
Private Type POINTAPI 'Type for holding X & Y co-ordinates
    X As Long
    Y As Long
End Type
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long 'API for finding the hWnd of the window under the cursor
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long                      'Finds the cursor's co-ordinates
Public Declare Sub InitCommonControls Lib "comctl32" ()

Private Const WM_USER = &H400
Private Const EM_SETTARGETDEVICE = (WM_USER + 72)
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Function BubbleSort(List() As String) As String()
    Dim bolSorted As Boolean
    Dim intLoopCounter As Integer
    Dim ReturnValue() As String
    ReturnValue = List
    
    bolSorted = False
    Do While bolSorted = False
        bolSorted = True
        For intLoopCounter = LBound(ReturnValue) To UBound(ReturnValue) - 1
            'string1 is greater than string2
            If StrComp(ReturnValue(intLoopCounter), ReturnValue(intLoopCounter + 1), vbBinaryCompare) = 1 Then
                Dim varTemp As Variant
                varTemp = ReturnValue(intLoopCounter)
                ReturnValue(intLoopCounter) = ReturnValue(intLoopCounter + 1)
                ReturnValue(intLoopCounter + 1) = varTemp
                bolSorted = False
            End If
        Next intLoopCounter
    Loop
    BubbleSort = ReturnValue
End Function

Public Sub Dent(TBox As Object, Optional ByVal Indent As Boolean = True, _
    Optional ByVal TabChar As String = vbTab)
    On Local Error Resume Next
    With TBox
        'Get the current selected text
        Dim lngSelStart As Long, lngSelLength As Long
        lngSelStart = .SelStart
        lngSelLength = .SelLength
        
        'Split the selected text by newlines
        Dim strLines() As String
        strLines() = Split(.SelText, vbNewLine, , vbTextCompare)
        
        Dim strOutput As String 'What gets put in the tbox
        Dim lngLoopCounter As Long 'Loop counter
        
        'If we are indenting
        If Indent Then
            'Loop for all lines
            For lngLoopCounter = LBound(strLines) To UBound(strLines)
                'Add a tab then the current line and a new line if nessaccary to the output string
                strOutput = strOutput & TabChar & strLines(lngLoopCounter) & IIf(lngLoopCounter < UBound(strLines), vbNewLine, vbNullString)
                'Length of the tab character must now be added to the selected length
                lngSelLength = lngSelLength + Len(TabChar)
            Next lngLoopCounter
        
        'Outdenting:
        Else
            'Loop for all lines
            For lngLoopCounter = LBound(strLines) To UBound(strLines)
                'If we have a tab
                If Left(strLines(lngLoopCounter), Len(vbTab)) = vbTab Then
                    'Remove it
                    strOutput = strOutput & Mid(strLines(lngLoopCounter), Len(vbTab) + 1)
                    'Length of the tab character must now be removed from the selected length
                    lngSelLength = lngSelLength - Len(TabChar)
                'No tab
                Else
                    'Just add the current line and don't decrement the selectedlength
                    strOutput = strOutput & strLines(lngLoopCounter)
                End If
                'Add a new line if nessaccary
                strOutput = strOutput & IIf(lngLoopCounter < UBound(strLines), vbNewLine, vbNullString)
            Next lngLoopCounter
        End If
        'Overwrite the new text and select the right amount
        .SelText = strOutput
        .SelStart = lngSelStart
        .SelLength = lngSelLength
    End With
End Sub

Public Function FixPath(ByVal Path As String) As String
    On Local Error Resume Next
    FixPath = Path & IIf(Right(Path, Len("\")) <> "\", "\", "")
End Function

Public Function Checked(Value As Boolean) As CheckBoxConstants
    On Local Error Resume Next
    Checked = IIf(Value = True, vbChecked, vbUnchecked)
End Function

Public Sub SetWrap(ByVal hWnd As Long, Optional ByVal Wrap As Boolean = False)
    Call SendMessageLong(hWnd, EM_SETTARGETDEVICE, 0, IIf(Wrap, 0, 1))
End Sub

Public Function TrueFalse(Checked As CheckBoxConstants) As Boolean
    On Local Error Resume Next
    TrueFalse = Checked = vbChecked
End Function

Public Function DoesDirectoryExist(ByVal Path As String) As Boolean
    On Local Error Resume Next
    'Remove past errors
    Err.Clear
    
    'Get the current Dir
    Dim CurrentDir As String
    CurrentDir = CurDir
    'Test the new path, and return whether it raised an error
    ChDir Path
    DoesDirectoryExist = (Err.Number <= 0)
    'Reset the original DIR
    ChDir CurrentDir
End Function

Public Function DoesFileExist(ByVal Filename As String) As Boolean
    On Local Error GoTo ErrorHandler
    Dim FileNumber As Integer
    DoesFileExist = False
    FileNumber = FreeFile
    'Open the file - if it exists no error will occur and continue to next statement
    Open Filename For Input As #FileNumber
    'Close it
    Close #FileNumber
    'Return true if the length of the file is > 0
    DoesFileExist = Len(Dir$(Filename)) > 0
    Exit Function
    
ErrorHandler:
End Function

Public Sub FormatButtons(Form As Object)
    On Local Error Resume Next
    'Loop for all controls in form
    Dim lngLoopCounter As Long
    For lngLoopCounter = 0 To Form.Controls.Count - 1
        'If Command Button set style to PushButton
        If TypeOf Form.Controls(lngLoopCounter) Is CommandButton Then _
            Call SendMessage(Form.Controls(lngLoopCounter).hWnd, _
            BM_SETSTYLE, BS_PUSHBUTTON, 0&)
    Next lngLoopCounter
End Sub

'Returns the parent of the specified dir (e.g. C:\ from C:\Windows)
Public Function GetParentDir(ByVal Dir As String) As String
    GetParentDir = Dir
    If Right(GetParentDir, 1) = "\" Then GetParentDir = Left(GetParentDir, Len(GetParentDir) - 1)
    Dim lngFound As Long
    lngFound = InStrRev(GetParentDir, "\")
    If lngFound > 0 Then GetParentDir = Left(GetParentDir, lngFound)
End Function

'Returns True if the specified key is down and False if not
Public Function IsButtonActive(ByVal VirtualKey As VKButtons, _
    ByVal hWnd As Long) As Boolean
    On Local Error Resume Next
    'Return if the button is down (less than -1) (and the window is active otherwise we might abort when not wanted)
    IsButtonActive = (GetKeyState(VirtualKey) < -1) 'And (GetActiveWindow = hWnd)
End Function

Public Function IsWindowHot(ByVal hWnd As Long) As Boolean
    On Local Error Resume Next
    Dim CursorPosition As POINTAPI 'Variable for cursor's X & Y values

    'Get the Cursor position
    Call GetCursorPos(CursorPosition)
    IsWindowHot = WindowFromPoint(CursorPosition.X, CursorPosition.Y) = hWnd 'Return     whether the object is hot
End Function

Public Function IsRECTHot(Area As RECT) As Boolean
    On Local Error Resume Next
    Dim CursorPosition As POINTAPI 'Variable for cursor's X & Y values

    'Get the Cursor position
    Call GetCursorPos(CursorPosition)
    IsRECTHot = CursorPosition.X >= Area.Left And _
        CursorPosition.X <= Area.Right And _
        CursorPosition.Y >= Area.Top And _
        CursorPosition.Y <= Area.Bottom
End Function

Public Function NewLine(Optional ByVal Amount As Integer = 1) As String
    On Local Error Resume Next
    'Loop for amount of new lines wanted
    Dim LoopCounter As Integer
    For LoopCounter = 1 To Amount
        'Add a new line to the return value
        NewLine = NewLine & vbNewLine
    Next LoopCounter
End Function

Public Sub OnTop(ByVal hWnd As Long, Optional ByVal OnTop As Boolean = True)
    On Local Error Resume Next
    'If the form is wanted to be on top
    If OnTop = True Then
        'Set it on top
        Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    'If the form isn't
    Else
        'Stop it always being on top
        Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If
End Sub

Public Function SetBorder(ByVal hWnd As Long, ByVal Visible As Boolean) As Boolean
    Dim lngStyle As Long
    'Get the current style
    lngStyle = GetWindowLong(hWnd, GWL_STYLE)
    'If we want the caption visible
    If Visible Then
        lngStyle = lngStyle Or WS_CAPTION
    'If we don't
    Else
        lngStyle = lngStyle And Not WS_CAPTION
    End If
    'Set the new style
    Call SetWindowLong(hWnd, GWL_STYLE, 0 Or lngStyle)
    'Show the changes
    Call SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or _
        SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE)
    'Return True if successful, false if not
    SetBorder = (lngStyle = GetWindowLong(hWnd, GWL_STYLE))
End Function

Public Sub SetKeyState(ByVal VirtualKey As VKButtons, _
    Optional ByVal Value As Integer = -1)
    Dim kbbState As KeyboardBytes
    'Get the current keyboard state
    Call GetKeyboardState(kbbState)
    'Invert the state if missing
    If Value = -1 Then Value = IIf(kbbState.kbByte(VirtualKey) = 0, 1, 0)
    'Change the desired key to the new value
    kbbState.kbByte(VirtualKey) = Value
    Call SetKeyboardState(kbbState)
End Sub

Public Function TwipsX(Optional ByVal _
    Amount As Integer = 1) As Single
    On Local Error Resume Next
    'Return the amount of twips in the specified number of pixels
    TwipsX = Amount * Screen.TwipsPerPixelX
End Function

Public Function TwipsY(Optional ByVal _
    Amount As Integer = 1) As Single
    On Local Error Resume Next
    'Return the amount of twips in the specified number of pixels
    TwipsY = Amount * Screen.TwipsPerPixelY
End Function

Public Sub Wait(Length As Long, Optional Sleep As Boolean = False)
    On Local Error GoTo ErrorHandler 'Exit on error otherwise we may be here for ever!
    Dim lngStartTime As Double 'Start time of the sub
    
    'Get the current time
    lngStartTime = GetTickCount
    
    'Loop while time taken is less than the length
    Do While GetTickCount - lngStartTime < Length
        'Do not freeze screen if wanted
        If Sleep = False Then DoEvents
    'On to next loop
    Loop
    Exit Sub
ErrorHandler:
End Sub

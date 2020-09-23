VERSION 5.00
Begin VB.Form frmColourPicker 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2070
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   2055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrActiveWindow 
      Interval        =   100
      Left            =   360
      Top             =   0
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   63
      Left            =   1795
      Top             =   1800
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   62
      Left            =   1545
      Top             =   1800
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   61
      Left            =   1290
      Top             =   1800
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   60
      Left            =   1035
      Top             =   1800
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   59
      Left            =   780
      Top             =   1800
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   58
      Left            =   525
      Top             =   1800
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   57
      Left            =   270
      Top             =   1800
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   56
      Left            =   15
      Top             =   1800
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   55
      Left            =   1795
      Top             =   1545
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   54
      Left            =   1545
      Top             =   1545
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   53
      Left            =   1290
      Top             =   1545
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   52
      Left            =   1035
      Top             =   1545
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   51
      Left            =   780
      Top             =   1545
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   50
      Left            =   525
      Top             =   1545
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   49
      Left            =   270
      Top             =   1545
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   48
      Left            =   15
      Top             =   1545
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00400040&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   47
      Left            =   1795
      Top             =   1290
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   46
      Left            =   1545
      Top             =   1290
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   45
      Left            =   1290
      Top             =   1290
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   44
      Left            =   1035
      Top             =   1290
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00004040&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   43
      Left            =   780
      Top             =   1290
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00404080&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   42
      Left            =   525
      Top             =   1290
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   41
      Left            =   270
      Top             =   1290
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   40
      Left            =   15
      Top             =   1290
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   39
      Left            =   1795
      Top             =   1030
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   38
      Left            =   1545
      Top             =   1030
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   37
      Left            =   1290
      Top             =   1030
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   36
      Left            =   1035
      Top             =   1030
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00008080&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   35
      Left            =   780
      Top             =   1030
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   34
      Left            =   525
      Top             =   1030
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   33
      Left            =   270
      Top             =   1030
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   32
      Left            =   15
      Top             =   1030
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00C000C0&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   31
      Left            =   1795
      Top             =   780
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   30
      Left            =   1545
      Top             =   780
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   29
      Left            =   1290
      Top             =   780
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   28
      Left            =   1035
      Top             =   780
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   27
      Left            =   780
      Top             =   780
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H000040C0&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   26
      Left            =   525
      Top             =   780
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   25
      Left            =   270
      Top             =   780
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   24
      Left            =   15
      Top             =   780
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FF00FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   23
      Left            =   1795
      Top             =   520
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   22
      Left            =   1545
      Top             =   520
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   21
      Left            =   1290
      Top             =   520
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   20
      Left            =   1035
      Top             =   520
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   19
      Left            =   780
      Top             =   520
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   18
      Left            =   525
      Top             =   520
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   17
      Left            =   270
      Top             =   520
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   16
      Left            =   15
      Top             =   520
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FF80FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   15
      Left            =   1795
      Top             =   270
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   14
      Left            =   1545
      Top             =   270
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   13
      Left            =   1290
      Top             =   270
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H0080FF80&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   12
      Left            =   1035
      Top             =   270
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   11
      Left            =   780
      Top             =   270
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   10
      Left            =   525
      Top             =   270
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   9
      Left            =   270
      Top             =   270
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   8
      Left            =   15
      Top             =   270
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   7
      Left            =   1795
      Top             =   15
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   6
      Left            =   1545
      Top             =   15
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   5
      Left            =   1290
      Top             =   15
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   4
      Left            =   1035
      Top             =   15
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   3
      Left            =   780
      Top             =   15
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   2
      Left            =   525
      Top             =   15
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   1
      Left            =   270
      Top             =   15
      Width           =   240
   End
   Begin VB.Shape shpColour 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   0
      Left            =   15
      Top             =   15
      Width           =   240
   End
   Begin VB.Shape shpSelected 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   -15
      Top             =   -15
      Width           =   300
   End
End
Attribute VB_Name = "frmColourPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strOpener As String
Private intSelected As Integer 'Which square is selected

Public Sub PositionForm()
    If strOpener = "frmMain" Then
        With frmMain.tbrFormatting
            'Get the position of this form
            Dim rctToolbarPos As RECT
            Call GetWindowRect(.hWnd, rctToolbarPos)
            'If we are off the bottom of the screen
            If (rctToolbarPos.Top * TwipsY) + .Buttons("Font Colour").Top + _
                .Buttons("Font Colour").Height + Me.Height > Screen.Height Then
                'Position us above the toolbar button
                Me.Top = (rctToolbarPos.Top * TwipsY) + .Buttons("Font Colour").Top - Me.Height
            'If we aren't off the screen
            Else
                'Position below the button
                Me.Top = (rctToolbarPos.Top * TwipsY) + .Buttons("Font Colour").Top + _
                    .Buttons("Font Colour").Height '- (Me.Height \ 2)
            End If
            
            'If we are off the right of the screen
            If (rctToolbarPos.Left * TwipsX) + .Buttons("Font Colour").Left + _
                .Buttons("Font Colour").Width + (Me.Width \ 2) > Screen.Width Then
                'Position us at the edge of the screen
                Me.Left = Screen.Width - Me.Width
            'If we aren't off the screen
            Else
                'Position at the start of the button
                Me.Left = (rctToolbarPos.Left * TwipsX) + .Buttons("Font Colour").Left
            End If
        End With
    End If
End Sub

Private Sub Form_Initialize()
    On Local Error Resume Next
    Call InitCommonControls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    'If escape unload
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    If strOpener = "frmMain" Then
        'Make sure the font colour button is pressed
        frmMain.tbrFormatting.Buttons("Font Colour").Value = tbrPressed
    End If
    
    'Load the last used custom colours from the INI file
    Dim intLoopCounter As Integer
    'Add a \ to the path if needed
    'Get the config file's path
    Dim strConfigFile As String
    strConfigFile = FixPath(App.Path) & "Config.ini"
    For intLoopCounter = shpColour.UBound - 15 To shpColour.UBound
        shpColour(intLoopCounter).FillColor = _
            GetINISetting(strConfigFile, "Colour Picker", _
            "Custom Colour" & shpColour.UBound - intLoopCounter, _
            shpColour(intLoopCounter).FillColor)
    Next intLoopCounter
    
    'Position this form
    Call OnTop(Me.hWnd)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    'If left button do the same as the Mouse_Move sub
    If Button = vbLeftButton Then Call Form_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    If Button = vbLeftButton Then
        'Get the colour the cursor is over
        intSelected = GetSelected(X, Y)
        'Move the selection rect to that colour
        Call MoveSelection(intSelected)
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    'Colour choosen
    If Button = vbLeftButton Then
        'Set the new colour
        If strOpener = "frmMain" Then
            Call frmMain.NewRTFColour(shpColour(intSelected).FillColor)
        End If
        'Unload
        Unload Me
    'Right click on a custom colour
    ElseIf Button = vbRightButton And GetSelected(X, Y) >= 48 Then
        'Get a colour from the show colour dialog
        Dim lngColour As Long
        lngColour = GetColour
        'If not cancel set the new colour
        If lngColour > -1 Then shpColour(GetSelected(X, Y)).FillColor = lngColour
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Local Error Resume Next
    'Save the custom colours
    Dim intLoopCounter As Integer
    'Get the config file's path
    Dim strConfigFile As String
    strConfigFile = FixPath(App.Path) & "Config.ini"
    For intLoopCounter = shpColour.UBound To shpColour.UBound - 15 Step -1
        Call SaveINISetting(strConfigFile, "Colour Picker", _
            "Custom Colour" & shpColour.UBound - intLoopCounter, _
            shpColour(intLoopCounter).FillColor)
    Next intLoopCounter
    
    If strOpener = "frmMain" Then
        'Font colour button = not pressed
        frmMain.tbrFormatting.Buttons("Font Colour").Value = tbrUnpressed
        'Return focus to the main form
        Dim rtfSelected As RichTextBox
        Select Case LCase(frmMain.tbsView.SelectedItem.Key)
            Case "code"
                Set rtfSelected = frmMain.rtfCode
            Case "notes"
                Set rtfSelected = frmMain.rtfNotes
            Case Else
                Exit Sub
        End Select
        rtfSelected.SetFocus
    End If
End Sub

Private Sub MoveSelection(Index As Integer)
    On Local Error Resume Next
    'Move the selection rect to the desired shape (minus 2px to account for border and 1px of white)
    shpSelected.Move shpColour(Index).Left - TwipsX(2), _
        shpColour(Index).Top - TwipsY(2)
End Sub

Public Sub SelectColour(ByVal Colour As Long)
    On Local Error Resume Next
    Dim intIndex As Integer, intLoopCounter As Integer
    intIndex = -1
    'Loop for all colour, while the correct one is not found
    intLoopCounter = -1
    Do While intLoopCounter <= shpColour.UBound And intIndex = -1
        'Increment the loop counter
        intLoopCounter = intLoopCounter + 1
        'If we've found the colour set it's index to the intIndex var - this will exit the loop
        If shpColour(intLoopCounter).FillColor = Colour Then intIndex = intLoopCounter
        'If we are past the last colour
        If intLoopCounter > shpColour.UBound Then
            'Hide the shape
            shpSelected.Move Me.Width
            'Exit sub so as not to position the higlight
            Exit Sub
        End If
    Loop
    'Position the highlight to the correct colour
    Call MoveSelection(intIndex)
End Sub

Private Function GetSelected(ByVal X As Single, Y As Single) As Integer
    On Local Error Resume Next
    Dim intXIndex As Integer, intYIndex As Integer
    'Get the index along the X-Axis of the colour
    intXIndex = (X \ (shpColour(0).Left + shpColour(0).Width))
    'Get the index along the Y-Axis of the colour
    intYIndex = (Y \ (shpColour(0).Top + shpColour(0).Height))
    'If X is below the lowest
    If intXIndex < 0 Then
        'Make = the lowest
        intXIndex = 0
    'If Y is above the highest
    ElseIf intXIndex > 7 Then
        'Make = the highest
        intXIndex = 7
    End If
    
    'If Y is below the lowest
    If intYIndex < 0 Then
        'Make = the lowest
        intYIndex = 0
    'If Y is above the highest
    ElseIf intYIndex > 7 Then
        'Make = the highest
        intYIndex = 7
    End If
    
    'Return the selected number
    GetSelected = intXIndex + (intYIndex * 8)
End Function

Private Sub tmrActiveWindow_Timer()
    On Local Error Resume Next
    'Unload this form if we are not the active window
    'this isn't the best way to do this but I don't know any other way!
    If GetActiveWindow <> Me.hWnd Then Unload Me
End Sub

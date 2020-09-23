VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3000
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4800
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
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
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private lngStartTime As Long 'When the form was loaded
Private bolShowSplash As Boolean 'Whether to show the splash

Private Sub Form_Initialize()
    On Local Error Resume Next
    Call InitCommonControls
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    With Me
        'If we are to show the splash
        bolShowSplash = GetINISetting(FixPath(App.Path) & "Config.ini", "General", "Show Splash", True)
        If bolShowSplash Then
            Dim strOutput As String
            'Print my credits!
            strOutput = "by Rick Bull"
            .CurrentX = TwipsX(7)
            .CurrentY = TwipsY(3)
            .ForeColor = vbBlack
            Print strOutput
            
            'Print the version
            strOutput = "Version " & App.Major & "." & App.Minor & "." & App.Revision
            .CurrentX = .ScaleWidth - .TextWidth(strOutput) - TwipsX(7)
            .CurrentY = .ScaleHeight - .TextHeight(strOutput) - TwipsY(3)
            .ForeColor = vbWhite
            Print strOutput
        
            'Show the changes
            If .AutoRedraw Then .Refresh
            
            'Show this form and put it on top
            .Show
            'Call OnTop(Me.hWnd)
            DoEvents
            
            'Get the start time
            'lngStartTime = GetTickCount
        End If
    End With
    'Load the main form
    Load frmMain
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Local Error Resume Next
    'If we are showing the splash
    'If bolShowSplash Then
    '    Const lngMinTime = 750 'The minimum time that the splash is shown for
    '    'Make sure the splash is shown for the minimum time
    '    Do While GetTickCount - lngStartTime < lngMinTime
    '        DoEvents
    '    Loop
    'End If
    'Show the main form
    frmMain.Show
    Call frmMain.SetFocus
End Sub

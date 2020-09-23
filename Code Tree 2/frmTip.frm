VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "Tip of the Day"
   ClientHeight    =   3480
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Tips at Startup"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   3180
      Width           =   1935
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Next Tip"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox picContainer 
      BackColor       =   &H00FFFFFF&
      Height          =   2955
      Left            =   120
      Picture         =   "frmTip.frx":0442
      ScaleHeight     =   2895
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label lblTipNumber 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip Number 0/0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2310
         TabIndex        =   6
         Top             =   2640
         Width           =   1305
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Did you know..."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Height          =   1995
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intCurrentTip As Integer 'The current index of the tip we are on
Private strTips() As String 'All the tips in an array

Private Sub cmdNextTip_Click()
    On Local Error Resume Next
    'Show the next tip
    Call SetTip(intCurrentTip + 1)
End Sub

Private Sub cmdOK_Click()
    On Local Error Resume Next
    Unload Me
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    'Get the value of 'Show Tips at Startup' checkbox
    chkLoadTipsAtStartup.Value = GetINISetting(FixPath(App.Path) & "Config.ini", "General", _
        "Show Tips at Startup", vbChecked)
    'Get the last tip we were on (-1 if not found)
    intCurrentTip = GetINISetting(FixPath(App.Path) & "Config.ini", "General", _
        "Last Tip", -1)
    'Load the tips
    Call GetTips
    'Set the tip to the next one
    Call SetTip(intCurrentTip + 1)
    Call FormatButtons(Me)
End Sub

Private Sub GetTips()
    On Local Error Resume Next
    Const strFileName As String = "Tips.txt" 'Filename of tips
    'Get the text from the tips file
    Dim strTemp As String
    strTemp = FixPath(App.Path) & strFileName
    If DoesFileExist(strTemp) Then strTemp = OpenText(strTemp)
    'If we have something
    If strTemp <> "" Then
        'Split it by new lines
        strTips = Split(strTemp, vbNewLine, , vbTextCompare)
        
    'No text
    Else
        'Tips length = 1
        ReDim strTips(0)
        'Default tip
        strTips(0) = "...that the tip of the day text file (" & strFileName & ") was not found in the application path." & _
            "Please create a file named " & strFileName & " with one tip per line."
    End If
    'Enable next tip command if there is more than one tip
    cmdNextTip.Enabled = UBound(strTips) > LBound(strTips)
End Sub

Private Sub SetTip(ByVal Index As Integer)
    On Local Error Resume Next
    'If the index is out of bounds, make it = 0
    If Index < LBound(strTips) Or Index > UBound(strTips) Then Index = LBound(strTips)
    'Set the tip
    lblTipText.Caption = strTips(Index)
    'Set the number
    lblTipNumber.Caption = "Tip Number " & Index + 1 & "/" & UBound(strTips) + 1
    'Set the new index to the variable
    intCurrentTip = Index
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Local Error Resume Next
    'Save whether we want tips at start up
    Call SaveINISetting(FixPath(App.Path) & "Config.ini", "General", _
        "Show Tips at Startup", chkLoadTipsAtStartup.Value)
    'Save current tip index
    Call SaveINISetting(FixPath(App.Path) & "Config.ini", "General", _
        "Last Tip", intCurrentTip)
End Sub

Private Sub lblTipNumber_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    'Left button
    If Button And vbLeftButton Then
        'Move tip back one, or put to end if we are at first tip
        Call SetTip(IIf(intCurrentTip > LBound(strTips), _
            intCurrentTip - 1, UBound(strTips)))
    'Right button
    ElseIf Button And vbRightButton Then
        'Move tip forward one, or put to start if we are at last tip
        Call SetTip(IIf(intCurrentTip < UBound(strTips), _
            intCurrentTip + 1, LBound(strTips)))
    End If
End Sub

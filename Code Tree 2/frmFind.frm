VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5175
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
   ScaleHeight     =   3900
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOptions 
      Caption         =   "Options:"
      Height          =   2055
      Left            =   240
      TabIndex        =   14
      Top             =   1680
      Width           =   3135
      Begin VB.CheckBox chkBeep 
         Caption         =   "&Beep on Mmatch"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox chkWholeWord 
         Caption         =   "&Whole Word Only"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkMatchCase 
         Caption         =   "&Match Case"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optFindIn 
         Caption         =   "Tree"
         Height          =   195
         Index           =   0
         Left            =   330
         TabIndex        =   2
         Top             =   525
         Width           =   735
      End
      Begin VB.OptionButton optFindIn 
         Caption         =   "Code"
         Height          =   195
         Index           =   1
         Left            =   1290
         TabIndex        =   3
         Top             =   525
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.OptionButton optFindIn 
         Caption         =   "Notes"
         Height          =   195
         Index           =   2
         Left            =   330
         TabIndex        =   4
         Top             =   765
         Width           =   735
      End
      Begin VB.OptionButton optFindIn 
         Caption         =   "Bookmarks"
         Height          =   195
         Index           =   3
         Left            =   1290
         TabIndex        =   5
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         Caption         =   "Find In:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Replace &All"
      Height          =   375
      Index           =   2
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Replace"
      Height          =   375
      Index           =   1
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2460
      Width           =   1335
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Index           =   0
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox txtReplace 
      Height          =   1125
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox txtFind 
      Height          =   1125
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      Caption         =   "Replace:"
      Height          =   195
      Index           =   2
      Left            =   2640
      TabIndex        =   13
      Top             =   120
      Width           =   630
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      Caption         =   "Find:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private lngLastPos As Long

Private Sub chkBeep_Click()
    bolBeepOnFind = chkBeep.Value And vbChecked
End Sub

Private Sub Form_Initialize()
    On Local Error Resume Next
    Call InitCommonControls
End Sub

Private Sub cmdFind_Click(Index As Integer)
    If Index > 0 Then MsgBox "Sorry but the replace functions are not yet implemented!", vbCritical Or vbOKOnly, "Not Coded": Exit Sub
    Call FindIn(Switch(optFindIn(Tree).Value, Tree, _
        optFindIn(Code), Code, _
        optFindIn(Notes), Notes, _
        optFindIn(Bookmarks), Bookmarks), txtFind.Text, txtReplace.Text, IIf(chkMatchCase.Value, vbBinaryCompare, vbTextCompare))
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    Dim strConfigFile As String
    strConfigFile = FixPath(App.Path) & "Config.ini"
    txtFind.Text = GetINISetting(strConfigFile, "Find Options", "Find Text", vbNullString)
    txtReplace.Text = GetINISetting(strConfigFile, "Find Options", "Replace Text", vbNullString)
    Select Case LCase(GetINISetting(strConfigFile, "Find Options", "Find In", "Code"))
        Case "tree"
            optFindIn(0).Value = True
        Case "code"
            optFindIn(1).Value = True
        Case "notes"
            optFindIn(2).Value = True
        Case "bookmarks"
            optFindIn(3).Value = True
    End Select
    chkMatchCase.Value = GetINISetting(strConfigFile, "Find Options", "Match Case", False)
    chkWholeWord.Value = GetINISetting(strConfigFile, "Find Options", "Whole Word", False)
    chkBeep.Value = GetINISetting(strConfigFile, "Find Options", "Beep on Match", False)
    
    'Make buttons 3D
    Call FormatButtons(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strConfigFile As String
    strConfigFile = FixPath(App.Path) & "Config.ini"
    Call SaveINISetting(strConfigFile, "Find Options", "Find Text", txtFind.Text)
    Call SaveINISetting(strConfigFile, "Find Options", "Replace Text", txtReplace.Text)
    Call SaveINISetting(strConfigFile, "Find Options", "Find In", _
        Switch(optFindIn(0).Value, "Tree", _
               optFindIn(1).Value, "Code", _
               optFindIn(2).Value, "Notes", _
               optFindIn(3).Value, "Bookmarks"))
    Call SaveINISetting(strConfigFile, "Find Options", "Match Case", chkMatchCase.Value)
    Call SaveINISetting(strConfigFile, "Find Options", "Whole Word", chkWholeWord.Value)
    Call SaveINISetting(strConfigFile, "Find Options", "Beep on Match", chkBeep.Value)
    frmMain.SetFocus
    Set frmFind = Nothing
End Sub

Private Sub optFindIn_Click(Index As Integer)
    On Local Error Resume Next
    cmdFind(1).Enabled = (Index = 1 Or Index = 2)
    cmdFind(2).Enabled = cmdFind(1).Enabled
    Select Case Index
        'Code
        Case 1
            frmMain.tbsView.Tabs("Code").Selected = True
        'Notes
        Case 2
            frmMain.tbsView.Tabs("Notes").Selected = True
        'Bookmarks
        Case 3
            frmMain.tbsView.Tabs("Bookmarks").Selected = True
    End Select
End Sub

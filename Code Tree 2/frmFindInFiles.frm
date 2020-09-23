VERSION 5.00
Begin VB.Form frmFindInFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find/Replace in Files"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "frmFindInFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin CodeTree.TextBoxButton tbbFolder 
      Height          =   330
      Left            =   60
      TabIndex        =   16
      Top             =   240
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   582
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CodeTree.Seperator Seperator1 
      Height          =   30
      Left            =   165
      TabIndex        =   15
      Top             =   4200
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3285
      Width           =   2655
   End
   Begin VB.Frame fraContainer 
      Caption         =   "Options:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   2775
      Begin VB.CheckBox chkBeep 
         Caption         =   "&Beep when done"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1005
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.CheckBox chkSubDirs 
         Caption         =   "Search &Sub Directories"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   765
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkConfirm 
         Caption         =   "Confirm &each change"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   525
         Width           =   1935
      End
      Begin VB.CheckBox chkCase 
         Caption         =   "&Case Sensitive"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   285
         Width           =   1455
      End
   End
   Begin VB.ComboBox cboPattern 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   45
      TabIndex        =   9
      Text            =   "*.*"
      Top             =   885
      Width           =   6060
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Replace"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4605
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2925
      Width           =   1335
   End
   Begin VB.TextBox txtReplace 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   1545
      Width           =   3015
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2925
      Width           =   1335
   End
   Begin VB.ListBox lstFiles 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   45
      TabIndex        =   3
      Top             =   4560
      Width           =   6075
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1545
      Width           =   3015
   End
   Begin VB.Label blblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Pattern:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   45
      TabIndex        =   10
      Top             =   645
      Width           =   600
   End
   Begin VB.Label blblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Replace with:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   3165
      TabIndex        =   7
      Top             =   1305
      Width           =   975
   End
   Begin VB.Label blblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Found/Replaced in Files:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   4320
      Width           =   1755
   End
   Begin VB.Label blblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Find what:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   45
      TabIndex        =   2
      Top             =   1305
      Width           =   765
   End
   Begin VB.Label blblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Folder:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   510
   End
End
Attribute VB_Name = "frmFindInFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const CB_FINDSTRINGEXACT = &H158
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private bolReplace As Boolean
Private bolError As Boolean

Private Sub cboPattern_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cboPattern_LostFocus
End Sub

Private Sub cboPattern_LostFocus()
    If SendMessage(cboPattern.hWnd, CB_FINDSTRINGEXACT, 0&, ByVal cboPattern.Text) <= -1 Then _
        cboPattern.AddItem cboPattern.Text
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click(Index As Integer)
    Me.Enabled = False
    lstFiles.Clear
    bolReplace = (Index = 1)
    bolError = False
    Call FindFiles(tbbFolder.Text, Me, "AddFile", , chkSubDirs.Value = vbChecked)
    Me.Enabled = True
    If chkBeep.Value = vbChecked Then Beep
    MsgBox "Searching has finished. The results of the search are displayed in the list box below." & _
        IIf(bolError, "Some errors occured whilst searching some files, perhaps due to a file type error (e.g. not a text file).", ""), _
        vbOKOnly Or vbInformation, "Finished"
End Sub

Public Sub AddFile(ByVal FileTitle As String, ByVal Filename As String)
    On Local Error Resume Next
    Dim strPattern() As String
    strPattern() = Split(cboPattern.Text, ";", , vbTextCompare)
    Dim intLoopCounter As Integer
    For intLoopCounter = LBound(strPattern) To UBound(strPattern())
        If FileTitle Like strPattern(intLoopCounter) Then
            Dim strFileText As String
            strFileText = OpenText(Filename, False)
            'If couldn't open file Then bolError = True
            If InStr(1, strFileText, txtFind.Text, IIf(chkCase.Value = vbChecked, vbBinaryCompare, vbTextCompare)) > 0 Then
                lstFiles.AddItem FileTitle & " - [" & Filename & "]"
                Dim bolRename As Boolean
                If chkConfirm.Value = vbUnchecked Then
                    bolRename = True
                Else
                    bolRename = (chkConfirm.Value = vbChecked And _
                        MsgBox("Replace text in the file """ & FileTitle & """", vbOKCancel Or vbQuestion, "Rename") = vbOK)
                End If
                If bolReplace And bolRename Then
                    strFileText = Replace(strFileText, txtFind.Text, txtReplace.Text, _
                        , , IIf(chkCase.Value = vbChecked, vbBinaryCompare, vbTextCompare))
                    Call SaveText(Filename, strFileText)
                End If
            End If
            Exit For
        End If
    Next intLoopCounter
End Sub

Private Sub Form_Initialize()
    On Local Error Resume Next
    Call InitCommonControls
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    Dim strInput As String
    Dim intLoopCounter As Integer
    'Get the list of patterns
    strInput = FixPath(App.Path) & "Find In Files Filter.txt"
    If DoesFileExist(strInput) Then
        strInput = OpenText(strInput)
        Dim strPatterns() As String
        strPatterns() = Split(strInput, vbNewLine, , vbTextCompare)
        For intLoopCounter = LBound(strPatterns) To UBound(strPatterns)
            'If it's not already there add it
            If SendMessage(cboPattern.hWnd, CB_FINDSTRINGEXACT, 0&, ByVal strPatterns(intLoopCounter)) <= -1 _
                And Trim(strPatterns(intLoopCounter) <> vbNullString) Then _
                cboPattern.AddItem strPatterns(intLoopCounter)
        Next intLoopCounter
    End If
    tbbFolder.Text = App.Path
    Set tbbFolder.Picture = LoadResPicture("OPENFOLDER", vbResBitmap)
    'Make buttons 3D
    Call FormatButtons(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intLoopCounter As Integer
    Dim strOutput As String
    'Save patterns
    For intLoopCounter = 0 To cboPattern.ListCount - 1
        strOutput = strOutput & cboPattern.List(intLoopCounter) & _
            IIf(intLoopCounter >= cboPattern.ListCount, vbNullString, vbNewLine)
    Next intLoopCounter
    Call SaveText(FixPath(App.Path) & "Find In Files Filter.txt", strOutput, OverWrite)
End Sub

Private Sub tbbFolder_ButtonClick()
    Dim strFolder As String
    strFolder = GetFolder(Me.hWnd)
    If strFolder <> vbNullString Then tbbFolder.Text = strFolder
End Sub

Private Sub txtFind_Change()
    cmdFind(0).Enabled = txtFind.Text <> vbNullString
    cmdFind(1).Enabled = cmdFind(0).Enabled
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
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
   ScaleHeight     =   5070
   ScaleWidth      =   5070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCatagories 
      Caption         =   "Tree:"
      Height          =   3855
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CheckBox chkShowRoot 
         Caption         =   "Show root node/directory"
         Height          =   195
         Left            =   1320
         TabIndex        =   39
         Top             =   360
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkCheckOnExpand 
         Caption         =   "Check for new/changed items on node expand"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Value           =   1  'Checked
         Width           =   4095
      End
      Begin VB.CheckBox chkVisible 
         Caption         =   "Visible"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox txtPattern 
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   4095
      End
      Begin MSComctlLib.Slider sldTreeTimerInterval 
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   25
         SmallChange     =   5
         Min             =   1
         Max             =   100
         SelStart        =   1
         TickStyle       =   1
         TickFrequency   =   5
         Value           =   1
      End
      Begin CodeTree.TextBoxButton tbbCodePath 
         Height          =   330
         Left            =   240
         TabIndex        =   27
         Top             =   2040
         Width           =   4095
         _ExtentX        =   7223
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
         ButtonWidth     =   315
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codes Path:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codes Pattern:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slower"
         Height          =   195
         Index           =   7
         Left            =   3840
         TabIndex        =   19
         Top             =   3480
         Width           =   480
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faster"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   18
         Top             =   3480
         Width           =   465
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tree Hide/Show Speed:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   17
         Top             =   2640
         Width           =   1695
      End
   End
   Begin VB.Frame fraCatagories 
      Caption         =   "Misc:"
      Height          =   3855
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CheckBox chkWordWrapNotes 
         Caption         =   "Word-wrap Notes window"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   1800
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkWordWrapCode 
         Caption         =   "Word-wrap Code window"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkStandardToolbar 
         Caption         =   "Standard"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox chkFormattingToolbar 
         Caption         =   "Format"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   3075
         Width           =   975
      End
      Begin VB.CheckBox chkAutoShowCode 
         Caption         =   "Auto Show Code"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkAutoShowNotes 
         Caption         =   "Auto Show Notes"
         Height          =   195
         Left            =   2040
         TabIndex        =   23
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkSaveBookmarks 
         Caption         =   "Save Bookmarks"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
      Begin CodeTree.TextBoxButton tbbBookmarksFile 
         Height          =   330
         Left            =   360
         TabIndex        =   28
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
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
      Begin CodeTree.TextBoxButton tbbStandardToolbarPath 
         Height          =   330
         Left            =   240
         TabIndex        =   35
         Top             =   2550
         Width           =   3975
         _ExtentX        =   7011
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
         ButtonWidth     =   315
      End
      Begin CodeTree.TextBoxButton tbbFormattingToolbarPath 
         Height          =   330
         Left            =   240
         TabIndex        =   36
         Top             =   3360
         Width           =   3975
         _ExtentX        =   7011
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
         ButtonWidth     =   315
      End
   End
   Begin VB.Frame fraCatagories 
      Caption         =   "General:"
      Height          =   3855
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   4575
      Begin ComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   4080
         TabIndex        =   31
         Top             =   2880
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   99
         BuddyControl    =   "txtMaxUndos"
         BuddyDispid     =   196610
         OrigLeft        =   4080
         OrigTop         =   2880
         OrigRight       =   4320
         OrigBottom      =   3135
         Increment       =   5
         Max             =   1000
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtMaxUndos 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   29
         Text            =   "10"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtNotesSeperator 
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   4095
      End
      Begin VB.CheckBox chkFixSeperator 
         Caption         =   "Replace <?KEYWORDS?> with the correct values"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   3855
      End
      Begin VB.CheckBox chkConfirmExit 
         Caption         =   "Confirm Exit"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox txtDefaultIndent 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3360
         TabIndex        =   6
         Text            =   "10"
         Top             =   2520
         Width           =   735
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   4080
         TabIndex        =   5
         Top             =   2520
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         Value           =   10
         BuddyControl    =   "txtDefaultIndent"
         BuddyDispid     =   196614
         OrigLeft        =   4560
         OrigTop         =   2520
         OrigRight       =   4800
         OrigBottom      =   2775
         Increment       =   10
         Max             =   1000
         Min             =   10
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin CodeTree.TextBoxButton tbbFormattingPath 
         Height          =   330
         Left            =   240
         TabIndex        =   26
         Top             =   2040
         Width           =   4095
         _ExtentX        =   7223
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
         ButtonWidth     =   315
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum undo for Code and Notes text boxes:"
         Height          =   390
         Index           =   8
         Left            =   240
         TabIndex        =   30
         Top             =   2880
         Width           =   2580
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes Seperator:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Indentation for Text Boxes:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   2580
         Width           =   2700
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formatting Path:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList imlCatagories 
      Left            =   4440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0000
            Key             =   "Tree"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":055C
            Key             =   "General"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":08B0
            Key             =   "Misc"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4605
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4605
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4605
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip tbsCatagories 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7646
      HotTracking     =   -1  'True
      ImageList       =   "imlCatagories"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tree"
            Key             =   "Tree"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Misc"
            Key             =   "Misc"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bolApplied As Boolean

Private Sub cmdApply_Click()
    Call SaveOptions
End Sub

Private Sub Form_Initialize()
    On Local Error Resume Next
    Call InitCommonControls
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call SaveOptions
    Unload Me
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    Set tbbBookmarksFile.Picture = LoadResPicture("SAVE", vbResBitmap)
    Set tbbCodePath.Picture = LoadResPicture("OPENFOLDER", vbResBitmap)
    Set tbbFormattingPath.Picture = LoadResPicture("OPENFOLDER", vbResBitmap)
    Set tbbFormattingToolbarPath.Picture = LoadResPicture("OPENFOLDER", vbResBitmap)
    Set tbbStandardToolbarPath.Picture = LoadResPicture("OPENFOLDER", vbResBitmap)
    Call LoadTabs(tbsCatagories)
    Call LoadOptions
    
    'Make buttons 3D
    Call FormatButtons(Me)
    bolApplied = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bolApplied Then
        Dim msgAnswer As VbMsgBoxResult
        msgAnswer = MsgBox("Some options will not show until you restart Code Tree." & vbNewLine & _
            "Would you like to restart now? You will loose any unsaved work!", vbYesNo, "Restart")
        If msgAnswer = vbYes Then
            Me.Hide
            frmMain.Hide
            Call Shell(FixPath(App.Path) & App.EXEName, vbNormalFocus)
            End
        End If
    End If
    Set frmOptions = Nothing
End Sub

Private Sub tbbBookmarksFile_ButtonClick()
    On Local Error Resume Next
    Dim strPath As String
    strPath = GetFileName("Configuration Files (*.ini)|*.ini|Text Files *.txt|*.txt|All Files|*.*", True)
    If strPath <> "" Then tbbBookmarksFile.Text = strPath
End Sub

Private Sub tbbCodePath_ButtonClick()
    On Local Error Resume Next
    Dim strPath As String
    strPath = GetFolder(Me.hWnd)
    If strPath <> "" Then tbbCodePath.Text = strPath
End Sub

Private Sub tbbFormattingPath_ButtonClick()
    On Local Error Resume Next
    Dim strPath As String
    strPath = GetFolder(Me.hWnd)
    If strPath <> "" Then tbbFormattingPath.Text = strPath
End Sub

Private Sub tbbFormattingToolbarPath_ButtonClick()
    On Local Error Resume Next
    Dim strPath As String
    strPath = GetFolder(Me.hWnd)
    If strPath <> "" Then tbbFormattingToolbarPath.Text = strPath
End Sub

Private Sub tbbStandardToolbarPath_ButtonClick()
    On Local Error Resume Next
    Dim strPath As String
    strPath = GetFolder(Me.hWnd)
    If strPath <> "" Then tbbStandardToolbarPath.Text = strPath
End Sub

Private Sub tbsCatagories_Click()
    Static intLastOn As Integer
    fraCatagories(intLastOn).Visible = False
    fraCatagories(tbsCatagories.SelectedItem.Index - 1).Visible = True
    intLastOn = tbsCatagories.SelectedItem.Index - 1
End Sub

Private Sub tbsCatagories_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call tbsCatagories_Click
End Sub

Private Sub LoadOptions()
    'Add a \ to the path if needed
    Dim strAppPath As String
    strAppPath = FixPath(App.Path)
    'Get the config file's path
    Dim strConfigFile As String
    strConfigFile = strAppPath & "Config.ini"
    
    '[General]
    txtPattern.Text = GetINISetting(strConfigFile, "General", _
        "Pattern", "*.rtf;*.txt") 'Code pattern
    tbbCodePath.Text = GetINISetting(strConfigFile, "General", _
        "Codes Path", strAppPath & "Codes\") 'Location of codes
    tbbFormattingPath.Text = GetINISetting(strConfigFile, "General", _
        "Formatting Templates Path", strAppPath & "Formatting\")
    chkAutoShowCode.Value = Checked(GetINISetting(strConfigFile, "General", _
        "Auto Show Code", True)) 'Show code on load
    chkAutoShowNotes.Value = Checked(GetINISetting(strConfigFile, "General", _
        "Auto Show Notes", False)) 'Show notes on load
    tbbBookmarksFile.Text = GetINISetting(strConfigFile, "General", _
        "Bookmarks File", strAppPath & "Bookmarks.ini") 'The old bookmarks file
    chkSaveBookmarks.Value = Checked(GetINISetting(strConfigFile, "General", _
        "Save Bookmarks", True)) 'The old bookmarks file
    If DoesFileExist(strAppPath & "Notes Seperator.txt") Then 'What seperates the notes as they are added
        txtNotesSeperator.Text = OpenText(strAppPath & "Notes Seperator.txt")
    Else
        'Default
        txtNotesSeperator.Text = "------------------" & NewLine & _
            "Taken from <?title?> in <?section?> on <?date?> at <?time?>" & NewLine(2)
    End If
    chkFixSeperator.Value = Checked(GetINISetting(strConfigFile, "General", _
        "Fix Seperator", True)) 'Show notes on load
    chkConfirmExit.Value = Checked(GetINISetting(strConfigFile, "General", _
        "Confirm Exit", True)) 'Confirm exit?
    chkShowRoot.Value = Checked(GetINISetting(strConfigFile, "General", _
        "Show Root", True)) 'Show root node/dir?
    txtDefaultIndent.Text = GetINISetting(strConfigFile, "General", _
        "Default Indent", 10) 'Indentation size
    sldTreeTimerInterval.Value = GetINISetting(strConfigFile, "General", _
        "Tree Speed", 25) 'Tree show/hide speed
    chkVisible.Value = Checked(GetINISetting(strConfigFile, "General", _
        "Tree Visible", True)) 'Whether tree is visible
    chkCheckOnExpand = Checked(GetINISetting(strConfigFile, "General", _
        "Check Codes On Node Expand", False)) 'Whether to check for changed items on node expand
    chkStandardToolbar.Value = Checked(GetINISetting(strConfigFile, "General", _
        "Standard Toolbar Visible", True)) 'Standard Toolbar Visible
    tbbStandardToolbarPath.Text = GetINISetting(strConfigFile, "General", _
        "Standard Toolbar Path", strAppPath & "Toolbars\Standard\")  'Standard Toolbar Path
    chkFormattingToolbar.Value = Checked(GetINISetting(strConfigFile, "General", _
        "Formatting Toolbar Visible", True)) 'Formatting Toolbar Visible
    tbbFormattingToolbarPath.Text = GetINISetting(strConfigFile, "General", _
        "Formatting Toolbar Path", strAppPath & "Toolbars\Formatting\")  'Formatting Toolbar Path
    'If GetINISetting(strAppPath  & "Config.ini", "General", _
    '    "Statusbar Visible", True) Then 'Status bar Visible
    '    sbrInfo.Visible = True
    '    mnuViewStatusBar.Checked = True
    'Else
    '    sbrInfo.Visible = False
    '    mnuViewStatusBar.Checked = False
    'End If
    'With tbrFormatting.Buttons("Change Case") 'Default Case
    '    .Tag = GetINISetting(strAppPath  & "Config.ini", "General", _
    '        "Change Case Default", "lower case")
    '    .ToolTipText = "Change Case - " & .Tag
    'End With
    'imlCodesLarge.MaskColor = GetColourFromString(GetINISetting(strAppPath  & "Config.ini", "General", _
    '    "Mask Colour", RGB(255, 0, 255)), RGB(255, 0, 255)) 'Mask Colour
    'imlCodesSmall.MaskColor = imlCodesLarge.MaskColor
    'imlToolbarFormatting.MaskColor = imlCodesLarge.MaskColor
    'imlToolbarStandard.MaskColor = imlCodesLarge.MaskColor
    'clsCodeUndo.MaxUndos = GetINISetting(strAppPath  & "Config.ini", "General", _
    '    "Max Undos", 99) 'Max undos
    txtMaxUndos.Text = GetINISetting(strConfigFile, "General", _
        "Max Undos", 99) 'Max undos
    chkWordWrapCode.Value = Checked(GetINISetting(strConfigFile, "General", _
        "Word Wrap Code", True)) 'Word wrap the code
    chkWordWrapNotes.Value = Checked(GetINISetting(strConfigFile, "General", _
        "Word Wrap Notes", True)) 'Word wrap the notes
End Sub

Private Sub SaveOptions()
    'Add a \ to the path if needed
    Dim strAppPath As String
    strAppPath = FixPath(App.Path)
    'Get the config file's path
    Dim strConfigFile As String
    strConfigFile = strAppPath & "Config.ini"
    
    '[General]
    Call SaveINISetting(strConfigFile, "General", _
        "Pattern", txtPattern.Text) 'Code pattern
    Call SaveINISetting(strConfigFile, "General", _
        "Codes Path", tbbCodePath.Text)  'Location of codes
    Call SaveINISetting(strConfigFile, "General", _
        "Formatting Templates Path", tbbFormattingPath.Text)
    Call SaveINISetting(strConfigFile, "General", _
        "Auto Show Code", TrueFalse(chkAutoShowCode.Value))  'Show code on load
    Call SaveINISetting(strConfigFile, "General", _
        "Auto Show Notes", TrueFalse(chkAutoShowNotes.Value))  'Show notes on load
    Call SaveINISetting(strConfigFile, "General", _
        "Bookmarks File", tbbBookmarksFile.Text)  'The old bookmarks file
    Call SaveINISetting(strConfigFile, "General", _
        "Save Bookmarks", TrueFalse(chkSaveBookmarks.Value))  'The old bookmarks file
    Call SaveText(strAppPath & "Notes Seperator.txt", txtNotesSeperator.Text)
    Call SaveINISetting(strConfigFile, "General", _
        "Fix Seperator", TrueFalse(chkFixSeperator.Value))  'Show notes on load
    Call SaveINISetting(strConfigFile, "General", _
        "Confirm Exit", TrueFalse(chkConfirmExit.Value))  'Confirm exit?
    Call SaveINISetting(strConfigFile, "General", _
        "Show Root", TrueFalse(chkShowRoot.Value))  'Show root node/dir?
    Call SaveINISetting(strConfigFile, "General", _
        "Default Indent", txtDefaultIndent.Text)  'Indentation size
    Call SaveINISetting(strConfigFile, "General", _
        "Tree Speed", sldTreeTimerInterval.Value)  'Tree show/hide speed
    Call SaveINISetting(strConfigFile, "General", _
        "Tree Visible", TrueFalse(chkVisible.Value)) 'Whether tree is visible
    Call SaveINISetting(strConfigFile, "General", _
        "Check Codes On Node Expand", TrueFalse(chkCheckOnExpand.Value)) 'Whether to check for changed items on node expand
    Call SaveINISetting(strConfigFile, "General", _
        "Standard Toolbar Visible", TrueFalse(chkStandardToolbar.Value))  'Standard Toolbar Visible
    Call SaveINISetting(strConfigFile, "General", _
        "Standard Toolbar Path", tbbStandardToolbarPath.Text)  'Standard Toolbar Path
    Call SaveINISetting(strConfigFile, "General", _
        "Formatting Toolbar Visible", TrueFalse(chkFormattingToolbar.Value))  'Formatting Toolbar Visible
    Call SaveINISetting(strConfigFile, "General", _
        "Formatting Toolbar Path", tbbFormattingToolbarPath.Text)  'Formatting Toolbar Path
    'If GetINISetting(strAppPath  & "Config.ini", "General", _
    '    "Statusbar Visible", True) Then 'Status bar Visible
    '    sbrInfo.Visible = True
    '    mnuViewStatusBar.Checked = True
    'Else
    '    sbrInfo.Visible = False
    '    mnuViewStatusBar.Checked = False
    'End If
    'With tbrFormatting.Buttons("Change Case") 'Default Case
    '    .Tag = GetINISetting(strAppPath  & "Config.ini", "General", _
    '        "Change Case Default", "lower case")
    '    .ToolTipText = "Change Case - " & .Tag
    'End With
    'imlCodesLarge.MaskColor = GetColourFromString(GetINISetting(strAppPath  & "Config.ini", "General", _
    '    "Mask Colour", RGB(255, 0, 255)), RGB(255, 0, 255)) 'Mask Colour
    'imlCodesSmall.MaskColor = imlCodesLarge.MaskColor
    'imlToolbarFormatting.MaskColor = imlCodesLarge.MaskColor
    'imlToolbarStandard.MaskColor = imlCodesLarge.MaskColor
    'clsCodeUndo.MaxUndos = GetINISetting(strAppPath  & "Config.ini", "General", _
    '    "Max Undos", 99) 'Max undos
    Call SaveINISetting(strConfigFile, "General", _
        "Max Undos", txtMaxUndos.Text)  'Max undos
    Call SaveINISetting(strConfigFile, "General", _
        "Word Wrap Code", TrueFalse(chkWordWrapCode.Value))  'Word wrap the code
    Call SaveINISetting(strConfigFile, "General", _
        "Word Wrap Notes", TrueFalse(chkWordWrapNotes.Value))  'Word wrap the notes
    Call frmMain.LoadOptions
    Call frmMain.Form_Resize
    bolApplied = True
End Sub



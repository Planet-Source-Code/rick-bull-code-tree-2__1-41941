VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFormattingEditor 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Formatting Editor"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   6030
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
   ScaleHeight     =   4200
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlDetails 
      Left            =   2160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormattingEditor.frx":0000
            Key             =   "String"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormattingEditor.frx":0354
            Key             =   "Keyword"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormattingEditor.frx":06A8
            Key             =   "Comment"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormattingEditor.frx":09FC
            Key             =   "Default"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwDetails 
      Height          =   4185
      Left            =   1635
      TabIndex        =   1
      Top             =   0
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   7382
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imlDetails"
      SmallIcons      =   "imlDetails"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Type"
         Text            =   "Type"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Bold"
         Text            =   "Bold"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Colour"
         Text            =   "Colour"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "End"
         Text            =   "End"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Font"
         Text            =   "Font"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Font Size"
         Text            =   "Font Size"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Index"
         Text            =   "Index"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "Italic"
         Text            =   "Italic"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "Start"
         Text            =   "Start"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "StrikeThru"
         Text            =   "StrikeThru"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Key             =   "Underline"
         Text            =   "Underline"
         Object.Width           =   882
      EndProperty
   End
   Begin VB.ListBox lstFiles 
      Height          =   4155
      IntegralHeight  =   0   'False
      ItemData        =   "frmFormattingEditor.frx":0D50
      Left            =   0
      List            =   "frmFormattingEditor.frx":0D52
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmFormattingEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
    On Local Error Resume Next
    Call InitCommonControls
End Sub

Private Sub LoadFormattingOptions()
    On Local Error GoTo ErrorHandler
    'Add a \ to the path if needed
    Dim strAppPath As String
    strAppPath = FixPath(App.Path)
    'Get the config file's path
    Dim strConfigFile As String
    strConfigFile = strAppPath & "Config.ini"
    'Get the first file in the formatting directory
    Dim strFileName() As String
    strFileName() = GetConfigFiles(GetINISetting(strConfigFile, _
        "General", "Formatting Templates Path", strAppPath & "Formatting\"))
    'Loop for all files
    Dim intLoopCounter As Integer
    For intLoopCounter = 0 To UBound(strFileName)
        lstFiles.AddItem strFileName(intLoopCounter)
    Next intLoopCounter
    Exit Sub
ErrorHandler:
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    Call LoadFormattingOptions
    'Make buttons 3D
    Call FormatButtons(Me)
End Sub

Private Sub Form_Resize()
    lstFiles.Height = Me.ScaleHeight
    lvwDetails.Move lstFiles.Left + lstFiles.Width, _
        0, Me.ScaleWidth - (lstFiles.Left + lstFiles.Width), _
        Me.ScaleHeight
End Sub

Private Sub lstFiles_Click()
    'Get the configuration
    Dim fmdConfig As FormattingDetails
    fmdConfig = GetConfig(frmMain.strFormattingPath & lstFiles.List(lstFiles.ListIndex), False)
    lvwDetails.ListItems.Clear
    Dim intLoopCounter As Integer
    Dim strTemp As String
    With fmdConfig.Comments
        For intLoopCounter = 1 To UBound(.StartString)
            lvwDetails.ListItems.Add , "Comments_" & intLoopCounter, _
                "Comment", "Comment", "Comment"
        
            lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , , .Bold(intLoopCounter)
            lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , , .Colour(intLoopCounter)
            lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , .EndString(intLoopCounter)
            lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , , .Font(intLoopCounter)
            lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , , .FontSize(intLoopCounter)
            lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , , intLoopCounter
            lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , , .Italic(intLoopCounter)
            lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , , .StartString(intLoopCounter)
            lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , , .StrikeThru(intLoopCounter)
            lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , , .Underline(intLoopCounter)
            'lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , "Start", .StartString(intLoopCounter)
            'lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , "End", .EndString(intLoopCounter)
            'lvwDetails.ListItems("Comments_" & intLoopCounter).ListSubItems.Add , "Index", intLoopCounter
        Next
    End With
    With fmdConfig.Default
        lvwDetails.ListItems.Add , "Default", _
            "Default", "Default", "Default"
        lvwDetails.ListItems("Default").ListSubItems.Add , , .Bold
        lvwDetails.ListItems("Default").ListSubItems.Add , , .Colour
        lvwDetails.ListItems("Default").ListSubItems.Add , , "N/A"
        lvwDetails.ListItems("Default").ListSubItems.Add , , .Font
        lvwDetails.ListItems("Default").ListSubItems.Add , , .FontSize
        lvwDetails.ListItems("Default").ListSubItems.Add , , "N/A"
        lvwDetails.ListItems("Default").ListSubItems.Add , , .Italic
        lvwDetails.ListItems("Default").ListSubItems.Add , , "N/A"
        lvwDetails.ListItems("Default").ListSubItems.Add , , .StrikeThru
        lvwDetails.ListItems("Default").ListSubItems.Add , , .Underline
    End With
    With fmdConfig.Keywords
        For intLoopCounter = 1 To UBound(.Keywords)
            strTemp = strTemp & .Keywords(intLoopCounter) & IIf(intLoopCounter < UBound(.Keywords), .Delimeter, "")
        Next intLoopCounter
        lvwDetails.ListItems.Add , "Keywords", "Keywords", "Keyword", "Keyword"
        
        lvwDetails.ListItems("Keywords").ListSubItems.Add , , .Bold
        lvwDetails.ListItems("Keywords").ListSubItems.Add , , .Colour
        lvwDetails.ListItems("Keywords").ListSubItems.Add , "N/A"
        lvwDetails.ListItems("Keywords").ListSubItems.Add , , .Font
        lvwDetails.ListItems("Keywords").ListSubItems.Add , , .FontSize
        lvwDetails.ListItems("Keywords").ListSubItems.Add , , "N/A"
        lvwDetails.ListItems("Keywords").ListSubItems.Add , , .Italic
        lvwDetails.ListItems("Keywords").ListSubItems.Add , , strTemp
        lvwDetails.ListItems("Keywords").ListSubItems.Add , , .StrikeThru
        lvwDetails.ListItems("Keywords").ListSubItems.Add , , .StrikeThru
        lvwDetails.ListItems("Keywords").ListSubItems.Add , , .Underline
    End With
    
    With fmdConfig.Strings
        For intLoopCounter = 1 To UBound(.StartString)
            lvwDetails.ListItems.Add , "Strings_" & intLoopCounter, _
                "Strings", "String", "String"
        
            lvwDetails.ListItems("Strings_" & intLoopCounter).ListSubItems.Add , , .Bold(intLoopCounter)
            lvwDetails.ListItems("Strings_" & intLoopCounter).ListSubItems.Add , , .Colour(intLoopCounter)
            lvwDetails.ListItems("Strings_" & intLoopCounter).ListSubItems.Add , .EndString(intLoopCounter)
            lvwDetails.ListItems("Strings_" & intLoopCounter).ListSubItems.Add , , .Font(intLoopCounter)
            lvwDetails.ListItems("Strings_" & intLoopCounter).ListSubItems.Add , , .FontSize(intLoopCounter)
            lvwDetails.ListItems("Strings_" & intLoopCounter).ListSubItems.Add , , intLoopCounter
            lvwDetails.ListItems("Strings_" & intLoopCounter).ListSubItems.Add , , .Italic(intLoopCounter)
            lvwDetails.ListItems("Strings_" & intLoopCounter).ListSubItems.Add , , .StartString(intLoopCounter)
            lvwDetails.ListItems("Strings_" & intLoopCounter).ListSubItems.Add , , .StrikeThru(intLoopCounter)
            lvwDetails.ListItems("Strings_" & intLoopCounter).ListSubItems.Add , , .Underline(intLoopCounter)
         Next
    End With
End Sub

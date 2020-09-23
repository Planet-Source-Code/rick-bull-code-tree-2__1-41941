VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Code Tree"
   ClientHeight    =   5460
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   6240
   Begin RichTextLib.RichTextBox rtfTemp 
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393217
      TextRTF         =   $"frmMain.frx":2CFA
   End
   Begin MSComctlLib.ImageList imlTabs 
      Left            =   4320
      Top             =   480
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
            Picture         =   "frmMain.frx":2DBC
            Key             =   "Bookmarks"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3110
            Key             =   "Code"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3464
            Key             =   "Notes"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlCodesLarge 
      Left            =   2880
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarFormatting 
      Left            =   2280
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483644
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbrFormatting 
      Height          =   330
      Left            =   2280
      TabIndex        =   8
      Top             =   840
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imlToolbarFormatting"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            Description     =   "Font"
            Style           =   4
            Object.Width           =   2000
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font Size"
            Description     =   "Font Size"
            Style           =   4
            Object.Width           =   700
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Description     =   "Bold"
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Description     =   "Italic"
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Description     =   "Underline"
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Strike-Thru"
            Description     =   "Strike-Thru"
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Description     =   "Left"
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Description     =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Description     =   "Right"
            Style           =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font Colour"
            Description     =   "Font Colour"
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bulletted List"
            Description     =   "Bulletted List"
            Style           =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Outdent"
            Description     =   "Outdent"
            Object.ToolTipText     =   "Outdent"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Indent"
            Description     =   "Indent"
            Object.ToolTipText     =   "Indent"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Character Map"
            Description     =   "Character Map"
            Style           =   1
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Change Case"
            Description     =   "Change Case"
            Object.ToolTipText     =   "Change Case - lower case"
            Object.Tag             =   "lower case"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "lower case"
                  Text            =   "lower case"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "UPPER CASE"
                  Text            =   "UPPER CASE"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "tOGGLE cASE"
                  Text            =   "tOGGLE cASE"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Proper Case"
                  Text            =   "Proper Case"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Sentance case"
                  Text            =   "Sentance case"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VaRy cAsE 1"
                  Text            =   "VaRy cAsE 1"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "vArY CaSe 2"
                  Text            =   "vArY CaSe 2"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      Begin VB.ComboBox cboFont 
         Height          =   315
         Left            =   15
         TabIndex        =   10
         Top             =   15
         Width           =   1930
      End
      Begin VB.ComboBox cboFontSize 
         Height          =   315
         Left            =   2000
         TabIndex        =   9
         Top             =   15
         Width           =   700
      End
   End
   Begin MSComDlg.CommonDialog cdlDialogs 
      Left            =   120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwBookmarks 
      Height          =   3255
      Left            =   2760
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5741
      Arrange         =   2
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imlCodesLarge"
      SmallIcons      =   "imlCodesSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Section"
         Object.Width           =   2540
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfNotes 
      Height          =   3255
      Left            =   2520
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5741
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      OLEDropMode     =   1
      TextRTF         =   $"frmMain.frx":37B8
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
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   5
      Top             =   5190
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Key             =   "Info"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            Key             =   "Position"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   1590
            MinWidth        =   2
            Key             =   "Spacer"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "NUM"
            Key             =   "Num"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "CAPS"
            Key             =   "Caps"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "INS"
            Key             =   "Ins"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "SCRL"
            Key             =   "Scrl"
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
   Begin MSComctlLib.ImageList imlToolbarStandard 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbrStandard 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarStandard"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "New"
            Object.ToolTipText     =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open"
            Object.ToolTipText     =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save"
            Object.ToolTipText     =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save As"
            Description     =   "Save As"
            Object.ToolTipText     =   "Save As"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save All"
            Description     =   "Save All"
            Object.ToolTipText     =   "Save All"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print"
            Object.ToolTipText     =   "Print"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Description     =   "Undo"
            Object.ToolTipText     =   "Undo"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Description     =   "Redo"
            Object.ToolTipText     =   "Redo"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Description     =   "Cut"
            Object.ToolTipText     =   "Cut"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Description     =   "Paste"
            Object.ToolTipText     =   "Paste"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Description     =   "Find"
            Object.ToolTipText     =   "Find"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy to Notes"
            Description     =   "Copy to Notes"
            Object.ToolTipText     =   "Copy to Notes"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Copy Selected"
                  Text            =   "Copy Selected"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Copy All"
                  Text            =   "Copy All"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add to Bookmarks"
            Description     =   "Add to Bookmarks"
            Object.ToolTipText     =   "Add to Bookmarks"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlCodesSmall 
      Left            =   120
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView tvwCodes 
      Height          =   4695
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   8281
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "imlCodesSmall"
      Appearance      =   1
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
   Begin VB.PictureBox picTreeSizer 
      BorderStyle     =   0  'None
      Height          =   4650
      Left            =   2040
      MouseIcon       =   "frmMain.frx":387A
      MousePointer    =   99  'Custom
      ScaleHeight     =   4650
      ScaleWidth      =   90
      TabIndex        =   0
      ToolTipText     =   "Drag to Resize or Double Click to Hide"
      Top             =   480
      Width           =   90
   End
   Begin RichTextLib.RichTextBox rtfCode 
      Height          =   3255
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5741
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      OLEDropMode     =   1
      TextRTF         =   $"frmMain.frx":39CC
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
   Begin MSComctlLib.TabStrip tbsView 
      Height          =   4695
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8281
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imlTabs"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Code"
            Key             =   "Code"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Notes"
            Key             =   "Notes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bookmarks"
            Key             =   "Bookmarks"
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileWindow 
         Caption         =   "&Window"
         Begin VB.Menu mnuFileWindowOnTop 
            Caption         =   "Always on &Top"
         End
         Begin VB.Menu mnuFileWindowFullScreen 
            Caption         =   "&Full Screen"
            Shortcut        =   {F11}
         End
         Begin VB.Menu mnuFileWindowTrayIcon 
            Caption         =   "&Tray Icon"
         End
         Begin VB.Menu mnuFileWindowSeperator1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileWindowNewWindow 
            Caption         =   "New &Window"
         End
      End
      Begin VB.Menu mnuFileSeperator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuEditSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditSeperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditSeperator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFormat 
         Caption         =   "&Format"
         Begin VB.Menu mnuEditFormatIndent 
            Caption         =   "&Indent"
         End
         Begin VB.Menu mnuEditFormatOutdent 
            Caption         =   "&Outdent"
         End
         Begin VB.Menu mnuEditFormatSeperator1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditFormatFormatAs 
            Caption         =   "&Format As"
            Begin VB.Menu mnuEditFormatFormatAsOption 
               Caption         =   "[NONE]"
               Enabled         =   0   'False
               Index           =   0
            End
            Begin VB.Menu mnuEditFormatFormatAsSeperator1 
               Caption         =   "-"
            End
            Begin VB.Menu mnuEditFormatFormatAsEditor 
               Caption         =   "Editor..."
            End
         End
         Begin VB.Menu mnuEditFormatSeperator2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEditFormatProperties 
            Caption         =   "Properties..."
         End
      End
      Begin VB.Menu mnuEditInsert 
         Caption         =   "&Insert..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbars 
         Caption         =   "&Toolbars"
         Begin VB.Menu mnuViewToolbarsStandard 
            Caption         =   "&Standard"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarsFormatting 
            Caption         =   "&Formatting"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarsSeperator1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolbarsShowHide 
            Caption         =   "&Show All"
            Index           =   0
         End
         Begin VB.Menu mnuViewToolbarsShowHide 
            Caption         =   "&Hide All"
            Index           =   1
         End
         Begin VB.Menu mnuViewToolbarsSeperator2 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuViewToolbarsCustomizeFormatting 
            Caption         =   "Customize..."
         End
         Begin VB.Menu mnuViewToolbarsCustomizeStandard 
            Caption         =   "Customize..."
         End
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuTree 
      Caption         =   "&Tree"
      Begin VB.Menu mnuTreeDeleteitem 
         Caption         =   "&Delete Selected Item..."
      End
      Begin VB.Menu mnuTreeSaveitem 
         Caption         =   "&Save Selected Item As..."
      End
      Begin VB.Menu mnuTreeRenameitem 
         Caption         =   "&Rename Selected Item..."
      End
      Begin VB.Menu mnuTreeNew 
         Caption         =   "&New Section..."
      End
      Begin VB.Menu mnuTreeSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeExpandAll 
         Caption         =   "&Expand All"
      End
      Begin VB.Menu mnuTreeCollapseAll 
         Caption         =   "&Collapse All"
      End
      Begin VB.Menu mnuTreeFindCodes 
         Caption         =   "&Find Codes"
      End
      Begin VB.Menu mnuTreeRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuTreeSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTreeCodeCount 
         Caption         =   "C&ode Count..."
      End
      Begin VB.Menu mnuTreeHideShow 
         Caption         =   "&Hide"
      End
   End
   Begin VB.Menu mnuBookmarks 
      Caption         =   "&Bookmarks"
      Begin VB.Menu mnuBookmarksAdd 
         Caption         =   "&Add Current Code/Section"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuBookmarksRemoveItem 
         Caption         =   "&Remove Current Item"
      End
      Begin VB.Menu mnuBookmarksClear 
         Caption         =   "&Clear All..."
      End
      Begin VB.Menu mnuBookmarksSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookmarksView 
         Caption         =   "View"
         Begin VB.Menu mnuBookmarksViewMode 
            Caption         =   "&Large Icon"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuBookmarksViewMode 
            Caption         =   "&Small Icon"
            Index           =   1
         End
         Begin VB.Menu mnuBookmarksViewMode 
            Caption         =   "&List"
            Index           =   2
         End
         Begin VB.Menu mnuBookmarksViewMode 
            Caption         =   "&Report"
            Index           =   3
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "T&ools"
      Begin VB.Menu mnuToolsMyComputer 
         Caption         =   "&My Computer"
      End
      Begin VB.Menu mnuToolsExplorer 
         Caption         =   "&Explorer"
      End
      Begin VB.Menu mnuToolsSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsFindInFiles 
         Caption         =   "&Find/Replace in Files..."
      End
      Begin VB.Menu mnuToolsSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpTip 
         Caption         =   "&Tip of the Day..."
      End
      Begin VB.Menu mnuHelpSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpLinks 
         Caption         =   "&Web Links"
         Begin VB.Menu mnuHelpLinksAdd 
            Caption         =   "&Add..."
         End
         Begin VB.Menu mnuHelpLinksManage 
            Caption         =   "&Manage..."
         End
         Begin VB.Menu mnuHelpLinksSeperator1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHelpLinksVBSites 
            Caption         =   "[None]"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuHelpChangeLog 
         Caption         =   "&Change Log..."
      End
      Begin VB.Menu mnuHelpThanks 
         Caption         =   "&Thanks/Credits..."
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Types
Private Type typCount
    Codes As Integer
    Sections As Integer
End Type

'API Declarations
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long 'Sends messages to windows
Private Declare Function GetFocus Lib "user32" () As Long
'Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long


'API Constants
Private Const EM_GETLINECOUNT = &HBA 'Gets the No of lines in a textbox
Private Const EM_LINEFROMCHAR = &HC9 'Gets the lines no in a textbox from the parsed position

'Variable Declarations
Private sngOffsetX As Single 'How much over the mouse was down on the sizer
Private strCodeFilename As String, strNotesFilename As String 'The filenames of the code and notes
Private strLastOn As String 'The last bolded item in the bookmarks

'Classes
Private clsCodeUndo As New clsUndo 'The undo info for the code rtf box
Private clsNotesUndo As New clsUndo 'The undo info for the notes rtf box

'Options variables:
Private bolAutoShowCode As Boolean 'Whether to automatically change to the code window when a code is loaded
Private bolAutoShowNotes As Boolean 'Whether to automatically change to the notes window when a code is loaded
Private bolConfirmExit As Boolean 'Whether to ask when closing
Private bolCheckCodesOnExpand As Boolean 'Whether to check for new codes when expand is clicked
Private bolWordWrapCode As Boolean 'Whether to word wrap the code window
Private bolWordWrapNotes As Boolean 'Whether to word wrap the notes window
Private bolShowRoot As Boolean 'Whether to show the root dir in the tree
Private strPattern As String 'The tree view filter
Private strCodePath As String 'Where the codes are kept
Public strFormattingPath As String 'Where the formatting options are kept
Private strBookmarksFile As String 'Where the bookmarks are kept
Private strStandardToolbarPath As String 'Where the images for the formatting toolbar are kept
Private strFormattingToolbarPath As String 'Where the images for the formatting toolbar are kept
Private bolSaveBookmarks As Boolean 'Whether to save the bookmarks
Private strNotesSeperator As String 'What seperates new notes
Private bolFixSeperator As Boolean 'Whether to replace keywords in the seperator
Private intDefaultIndent As Integer 'How much to indent text by
Private intTreeTimerInterval As Integer 'The update speed when hide/showing the tree
'Custom Constants
Private Const intMinWidth As Integer = 50 'The amount of px that tree/tabs can be


Private Sub cboFont_Change()
    On Local Error Resume Next
    Dim rtfSelected As RichTextBox
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = rtfCode
        Case "notes"
            Set rtfSelected = rtfNotes
        Case Else
            Exit Sub
    End Select
    rtfSelected.SelFontName = cboFont.Text
End Sub

Private Sub cboFont_Click()
    On Local Error Resume Next
    Dim rtfSelected As RichTextBox
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = rtfCode
        Case "notes"
            Set rtfSelected = rtfNotes
        Case Else
            Exit Sub
    End Select
    rtfSelected.SelFontName = cboFont.Text
    rtfSelected.SetFocus
End Sub

Private Sub cboFontSize_Change()
    On Local Error Resume Next
    Dim rtfSelected As RichTextBox
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = rtfCode
        Case "notes"
            Set rtfSelected = rtfNotes
        Case Else
            Exit Sub
    End Select
    rtfSelected.SelFontSize = cboFontSize.Text
End Sub

Private Sub cboFontSize_Click()
    On Local Error Resume Next
    Dim rtfSelected As RichTextBox
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = rtfCode
        Case "notes"
            Set rtfSelected = rtfNotes
        Case Else
            Exit Sub
    End Select
    rtfSelected.SelFontSize = cboFontSize.Text
    rtfSelected.SetFocus
End Sub

Private Sub Form_Initialize()
    On Local Error Resume Next
    Call InitCommonControls
End Sub

Public Sub AddCodes(ByVal FileTitle As String, ByVal Filename As String)
    On Local Error Resume Next
    Dim strTemp As String
    'Directory
    If GetAttr(Filename) And vbDirectory Then
        'If the small bmp is found and not already loaded _ load it
        strTemp = FixPath(Filename) & "Small.bmp"
        If DoesFileExist(strTemp) And DoesListImageExist(imlCodesSmall, Filename) = False Then _
            imlCodesSmall.ListImages.Add , FixPath(Filename), LoadPicture(strTemp)
        'If the large bmp is found and not already loaded _ load it
        strTemp = FixPath(Filename) & "Large.bmp"
        If DoesFileExist(strTemp) And DoesListImageExist(imlCodesLarge, Filename) = False Then _
            imlCodesLarge.ListImages.Add , FixPath(Filename), LoadPicture(strTemp)
        
        'If the node is not already loaded in the tree view
        If DoesNodeExist(tvwCodes, Filename) = False Then
            'Get the directory one up
            strTemp = Left(Filename, Len(Filename) - Len(FileTitle))
            'If the parent node exisits
            If DoesNodeExist(tvwCodes, strTemp) Then
                'Add this one as a child
                tvwCodes.Nodes.Add strTemp, tvwChild, FixPath(Filename), FileTitle, FixPath(Filename)
            'No parent
            Else
                'Load it as a top level node
                tvwCodes.Nodes.Add , , FixPath(Filename), FileTitle, FixPath(Filename)
            End If
        End If
    Else
        Dim strPatterns() As String
        strPatterns() = Split(strPattern, ";")
        Dim intLoopCounter As Integer
        'Loop for patterns
        For intLoopCounter = LBound(strPatterns) To UBound(strPatterns)
            'If this pattern is like the filename
            If FileTitle Like strPatterns(intLoopCounter) Then
                'Exit this loop and go to the next bit
                Exit For
            'If we have checked all patterns and this isn't the correct file type
            ElseIf intLoopCounter >= UBound(strPatterns) Then
                'Exit sub and don't do this next bit
                Exit Sub
            End If
        Next intLoopCounter

        'If the node is not already loaded in the tree view
        If DoesNodeExist(tvwCodes, Filename) = False Then
            'Get the directory one up
            strTemp = Left(Filename, Len(Filename) - Len(FileTitle))
            'If the parent node exisits
            If DoesNodeExist(tvwCodes, strTemp) Then
                'Add this one as a child
                tvwCodes.Nodes.Add strTemp, tvwChild, Filename, _
                    Left(FileTitle, Len(FileTitle) - Len(strPatterns(intLoopCounter)) + 1), strTemp
            'No parent
            Else
                'Load it as a top level node
                tvwCodes.Nodes.Add , , Filename, Left(FileTitle, _
                    Len(FileTitle) - Len(strPatterns(intLoopCounter)) + 1), FixPath(Filename)
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    'On Local Error Resume Next
    'Get the options
    Call LoadOptions
    
    'Make sure the size is at the front
    Call picTreeSizer.ZOrder(vbBringToFront)
    
    'Load the folders/files
    'Call AddNodes(tvwCodes, strCodePath, , , False)

    'If (bolShowRoot = True) Then _
        tvwCodes.Nodes.Add , , strCodePath & "\", "Root"
        'Call FindFiles(GetParentDir(strCodePath), Me, "AddCodes")
        
    Call FindFiles(strCodePath, Me, "AddCodes")
    'Load all the codes and then collapse all the nodes
    'Call mnuTreeExpandAll_Click
    'Call mnuTreeCollapseAll_Click
    'Load the bookmarks
    If bolSaveBookmarks Then Call LoadBookmarks(strBookmarksFile)
    
    'Load the toolbar images: Standard
    Call LoadBitmaps(tbrStandard, strStandardToolbarPath)
    Call LoadToolbar(tbrStandard)
    'Load the fonts
    Call LoadFonts(cboFont)
    cboFont.Text = rtfCode.Font.Name

    'Load the font sizes
    Dim intLoopCounter As Integer, intStep As Integer
    intStep = 2
    intLoopCounter = 6
    Do While intLoopCounter <= 96
        cboFontSize.AddItem intLoopCounter
        If intLoopCounter >= 20 And intLoopCounter < 40 Then
            intStep = 4
        ElseIf intLoopCounter >= 40 And intLoopCounter < 60 Then
            intStep = 8
        ElseIf intLoopCounter >= 60 And intLoopCounter < 92 Then
            intStep = 16
        ElseIf intLoopCounter >= 92 Then
            intStep = 32
        End If
        intLoopCounter = intLoopCounter + intStep
    Loop
    cboFontSize.Text = rtfCode.Font.Size

    'Formatting
    Call LoadBitmaps(tbrFormatting, strFormattingToolbarPath)
    Call LoadToolbar(tbrFormatting)
    
    'Set the save button's enabled value
    Call SetSaveButton
    
    'Load the tabs images
    Call LoadTabs(tbsView)
    
    'Set the char/line pos and stuff in the status bar
    Call SelectionChange
    
    'Load the formatting options
    Call LoadFormattingOptions
    
    'Load the links for the Help... menu
    Call LoadLinks
    
    'Make buttons 3D
    Call FormatButtons(Me)
    
    'Unload the splash screen
    Unload frmSplash
    
    If GetINISetting(FixPath(App.Path) & "Config.ini", "General", _
        "Show Tips at Startup", vbChecked) Then
        Load frmTip
        frmTip.Show vbModal, Me
    End If
    bolLoopRunning = False
End Sub

Private Sub Form_Paint()
    On Local Error Resume Next
    Static intLastWindowState  As Integer
    'If we have a different window state
    If Me.WindowState <> intLastWindowState Then
        'Resize the elements - if you maximize some things don't position properly
        Call Form_Resize
        'Set the current state to the variable
        intLastWindowState = Me.WindowState
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Local Error Resume Next
    'Make sure old windows are closed - I'm looking at you frmColourPicker!
    DoEvents
    'Check the user really wants to quit
    If bolConfirmExit Then _
        If MsgBox("Are you sure you want to exit Code Tree?", _
            vbYesNo Or vbQuestion, "Quit?") = vbNo Then Cancel = vbCancel
End Sub

Public Sub Form_Resize()
    On Local Error Resume Next
    'Get weird effects if we do this when minimized, so DON'T DO IT!
    If Me.WindowState <> vbMinimized Then
        'Adjust Tree Sizer if to far right
        If picTreeSizer.Left > ScaleWidth - TwipsX(intMinWidth) Then _
            picTreeSizer.Left = ScaleWidth - TwipsX(intMinWidth)
        'Set the position of things
        Call PositionElements
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Local Error Resume Next
    If bolLoopRunning Then
        Cancel = vbCancel
        Call MsgBox("Please wait until all processes have finished", vbExclamation Or vbOKOnly Or vbApplicationModal, "Please Wait")
    Else
        Call SaveOptions
        'Save the bookmarks if wanted
        If bolSaveBookmarks Then Call SaveBookmarks(strBookmarksFile)
    End If
End Sub

Private Sub lvwBookmarks_DblClick()
    On Local Error Resume Next
    'Load the code and do whatever should be done when a node is clicked
    With tvwCodes.Nodes(lvwBookmarks.SelectedItem.Key)
        .Selected = True
        .Expanded = True
    End With
    Call tvwCodes_NodeClick(tvwCodes.Nodes(lvwBookmarks.SelectedItem.Key))
End Sub

Private Sub lvwBookmarks_KeyDown(KeyCode As Integer, Shift As Integer)
    On Local Error Resume Next
    'Delete = delete the bookmark
    If KeyCode = vbKeyDelete Then
        Call mnuBookmarksRemoveItem_Click
    'Enter = Select the code
    ElseIf KeyCode = vbKeyReturn Then
        Call lvwBookmarks_DblClick
    End If
End Sub

Private Sub lvwBookmarks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Show the view menu if right button
    If Button = vbRightButton Then Call PopupMenu(mnuBookmarksView)
End Sub

Private Sub lvwBookmarks_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Loop for all files
    Dim intLoopCounter As Integer
    For intLoopCounter = 1 To Data.Files.Count
        'If the file is in the TreeView and not already a bookmark add it to the bookmarks
        If DoesNodeExist(tvwCodes, Data.Files(intLoopCounter)) And _
            DoesListItemExist(lvwBookmarks, Data.Files(intLoopCounter)) = False Then
            Call AddBookmark(Data.Files(intLoopCounter))
        'Same as above but in case we have a folder
        ElseIf DoesNodeExist(tvwCodes, Data.Files(intLoopCounter) & "\") _
            And DoesListItemExist(lvwBookmarks, Data.Files(intLoopCounter) & "\") = False _
            Then Call AddBookmark(Data.Files(intLoopCounter) & "\")
        End If
    Next intLoopCounter
End Sub

Private Sub mnuBookmarksAdd_Click()
    On Local Error Resume Next
    'Add the current bookmark
    Call AddBookmark(tvwCodes.SelectedItem.Key)
    Exit Sub
End Sub

Private Sub mnuBookmarksClear_Click()
    On Local Error Resume Next
    'Clear all bookmarks if user is sure
    If MsgBox("Are you sure you want to clear all of your bookmarks?", _
        vbQuestion Or vbOKCancel, "Clear All Bookmarks") = vbOK Then _
            lvwBookmarks.ListItems.Clear
End Sub

Private Sub mnuBookmarksRemoveItem_Click()
    On Local Error Resume Next
    'If there is a selection in the bookmarks
    If HasSelectedItem(lvwBookmarks) Then
        'Loop for all list itmes
        Dim intLoopCounter As Integer
        Do While intLoopCounter < lvwBookmarks.ListItems.Count
            intLoopCounter = intLoopCounter + 1
            'If this listitem is selected
            If lvwBookmarks.ListItems(intLoopCounter).Selected Then
                'Remove it
                Call lvwBookmarks.ListItems.Remove(intLoopCounter)
                'Start the loop from the one before the one we just removed, which is now the current one
                intLoopCounter = intLoopCounter - 1
            End If
            'Remove it
            'Call lvwBookmarks.ListItems.Remove(lvwBookmarks.SelectedItem.Index)
        Loop
    'If there isn't
    Else
        'Tell the user
        Call MsgBox("No item selected in bookmarks.", vbExclamation Or vbOKOnly, "Error")
    End If
End Sub

Private Sub mnuBookmarksViewMode_Click(Index As Integer)
    On Local Error Resume Next
    With lvwBookmarks
        'Remove the currnet check
        mnuBookmarksViewMode(.View).Checked = False
        'Set the new view for the bookmarks
        .View = Index
        'Add a check to the new view mode menu
        mnuBookmarksViewMode(.View).Checked = True
    End With
End Sub

Private Sub mnuEdit_Click()
    mnuEditPaste.Enabled = Clipboard.GetText <> vbNullString
End Sub

Private Sub mnuEditCopy_Click()
    On Local Error Resume Next
    Call EditRTFText(WM_COPY)
End Sub

Private Sub mnuEditCut_Click()
    On Local Error Resume Next
    Call EditRTFText(WM_CUT)
End Sub

Private Sub mnuEditDelete_Click()
    On Local Error Resume Next
    Call EditRTFText(WM_CLEAR)
End Sub

Private Sub mnuEditFind_Click()
    On Local Error Resume Next
    Load frmFind
    frmFind.Show vbModeless, Me
End Sub

Private Sub mnuEditFindNext_Click()
    Dim lngActivePane As Long
    lngActivePane = GetFocus
    Select Case lngActivePane
        Case rtfCode.hWnd
            Call FindIn(Code)
            rtfCode.SetFocus
        Case rtfNotes.hWnd
            Call FindIn(Notes)
            rtfNotes.SetFocus
        Case tvwCodes.hWnd
            Call FindIn(Tree)
            tvwCodes.SetFocus
        Case lvwBookmarks.hWnd
            Call FindIn(Bookmarks)
            lvwBookmarks.SetFocus
        Case Else
            Call MsgBox("Please select a pane to search in first.", vbCritical Or vbOKOnly, "Select Window")
    End Select
    'Call SetFocus(lngActivePane)
End Sub

Private Sub mnuEditFormatIndent_Click()
    On Local Error Resume Next
    'Choose which tab is active and indent the right rtf text by the default amount
    Select Case LCase(tbsView.SelectedItem.Key)
        'Code
        Case "code"
            rtfCode.SelIndent = rtfCode.SelIndent + TwipsX(intDefaultIndent)
        'Notes
        Case "notes"
            rtfNotes.SelIndent = rtfNotes.SelIndent + TwipsX(intDefaultIndent)
    End Select
End Sub

Private Sub mnuEditFormatOutdent_Click()
    On Local Error Resume Next
    'Choose which tab is active and indent the right rtf text by the default amount
    Select Case LCase(tbsView.SelectedItem.Key)
        'Code
        Case "code"
            rtfCode.SelIndent = rtfCode.SelIndent - TwipsX(intDefaultIndent)
        'Notes
        Case "notes"
            rtfNotes.SelIndent = rtfNotes.SelIndent - TwipsX(intDefaultIndent)
    End Select
End Sub

Private Sub mnuEditFormatProperties_Click()
    On Local Error Resume Next
    Dim rtfSelected As RichTextBox
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = rtfCode
        Case "notes"
            Set rtfSelected = rtfNotes
        Case Else
            Call MsgBox("You cannot view the properties for this. Please select either the Code or Notes window.", _
                vbExclamation Or vbOKOnly, "Error")
            Exit Sub
    End Select
    'Show the format dialog
    Load frmFormat
    With frmFormat
        'Set the indents in pixels
        .txtIndent.Text = rtfSelected.SelIndent \ TwipsX
        .txtRightIndent.Text = rtfSelected.SelRightIndent \ TwipsX
        .txtHangingIndent.Text = rtfSelected.SelHangingIndent \ TwipsX
        .txtBulletIndent.Text = rtfSelected.BulletIndent \ TwipsX
        'Show the form
        .Show vbModal, Me
    End With
End Sub

Private Sub mnuEditInsert_Click()
    On Local Error Resume Next
    Dim rtfSelected As RichTextBox
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = rtfCode
        Case "notes"
            Set rtfSelected = rtfNotes
        Case Else
            Call MsgBox("You cannot insert objects here. Please select either the Code or the Notes window.", _
                vbExclamation Or vbOKOnly, "Error")
            Exit Sub
    End Select
    
    Dim strFileName As String
    strFileName = GetFileName("All Supported Files|*.bmp;*.gif;*.jpeg;*.jpg;*.doc;*.xls" & _
        "|Images (*.bmp;*.gif;*.jpeg;*.jpg)|*.bmp;*.gif;*.jpeg;*.jpg" & _
        "|Word Documents (*.doc)|*.doc|Excel Spreadsheets (*.xls)|*.xls" & _
        "|All Files|*.*")
    If strFileName <> "" Then Call rtfSelected.OLEObjects.Add(, , strFileName)
End Sub

Private Sub mnuEditPaste_Click()
    On Local Error Resume Next
    Call EditRTFText(WM_PASTE)
End Sub

Private Sub mnuEditRedo_Click()
On Local Error Resume Next
    Dim udiInto As UndoInfo
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            With rtfCode
                'If we have Undos left
                If clsCodeUndo.CanRedo Then
                    'Start undoing
                    clsCodeUndo.Undoing = True
                    'Get the undo info
                    udiInto = clsCodeUndo.GetRedo
                    'Set it to the text box
                    .TextRTF = udiInto.Text
                    .SelStart = udiInto.SelStart
                    .SelLength = udiInto.SelLength
                    'End undoing
                    clsCodeUndo.Undoing = False
                'If we don't
                Else
                    'Beep to indicate no more left
                    Beep
                End If
            End With
        Case "notes"
            With rtfNotes
                'If we have Undos left
                If clsNotesUndo.CanRedo Then
                    'Start undoing
                    clsCodeUndo.Undoing = True
                    'Get the undo info
                    udiInto = clsNotesUndo.GetRedo
                    'Set it to the text box
                    .TextRTF = udiInto.Text
                    .SelStart = udiInto.SelStart
                    .SelLength = udiInto.SelLength
                    'End undoing
                    clsCodeUndo.Undoing = False
                'If we don't
                Else
                    'Beep to indicate no more left
                    Beep
                End If
            End With
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub mnuEditSelectAll_Click()
    On Local Error Resume Next
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            With rtfCode
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        Case "notes"
            With rtfNotes
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
    End Select
End Sub

Private Sub mnuEditUndo_Click()
    On Local Error Resume Next
    Dim udiInto As UndoInfo
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            With rtfCode
                'If we have Undos left
                If clsCodeUndo.CanUndo Then
                    'Start undo
                    clsCodeUndo.Undoing = True
                    'Get the undo info
                    udiInto = clsCodeUndo.GetUndo
                    'Set it to the text box
                    .TextRTF = udiInto.Text
                    .SelStart = udiInto.SelStart
                    .SelLength = udiInto.SelLength
                    'End undo
                    clsCodeUndo.Undoing = False
                'If we don't
                Else
                    'Beep to indicate no more left
                    Beep
                End If
            End With
        Case "notes"
            With rtfNotes
                'If we have Undos left
                If clsNotesUndo.CanUndo Then
                    'Start undo
                    clsCodeUndo.Undoing = True
                    'Get the undo info
                    udiInto = clsNotesUndo.GetUndo
                    'Set it to the text box
                    .TextRTF = udiInto.Text
                    .SelStart = udiInto.SelStart
                    .SelLength = udiInto.SelLength
                    'End undo
                    clsCodeUndo.Undoing = False
                'If we don't
                Else
                    'Beep to indicate no more left
                    Beep
                End If
            End With
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub mnuFileExit_Click()
    On Local Error Resume Next
    Unload Me
End Sub

Private Sub mnuFileWindowFullScreen_Click()
    On Local Error Resume Next
    'Static last window state so that if/when this sub is called again we will know the last state
    Static intLastState As Integer

    With Me
        'Invert the menu's check
        .mnuFileWindowFullScreen.Checked = Not .mnuFileWindowFullScreen.Checked
        'If we want full screen
        If .mnuFileWindowFullScreen.Checked Then
            'Get the current state (soon to be the last)
            intLastState = .WindowState
            'Set window state to minimized first otherwise taskbar will be visible
            .WindowState = vbNormal
            'Set the new border style (none for full screen)
            Call SetBorder(.hWnd, False)
            'Set the new window state to maximized
            .WindowState = vbMaximized
            'Make sure we cover the taskbar
            Call Me.ZOrder(vbBringToFront)
        'Normal mode
        Else
            'Set the new border style (none for full screen)
            Call SetBorder(.hWnd, True)
            'If we weren't minimized
            If intLastState <> vbMinimized Then
                'Restore the old state
                .WindowState = intLastState
            'If we were minimzed (or nothing which it should be)
            Else
                'Set to normal
                .WindowState = vbNormal
            End If
        End If
    End With
End Sub

Private Sub mnuFileWindowNewWindow_Click()
    On Local Error Resume Next
    'Save the options so that the new window will take the same ones as this one
    Call SaveOptions
    'Launch a new window of this app
    Call Shell(FixPath(App.Path) & App.EXEName, vbNormalFocus)
End Sub

Private Sub mnuFileWindowOnTop_Click()
    On Local Error Resume Next
    'Change the checked value of the menu
    mnuFileWindowOnTop.Checked = Not mnuFileWindowOnTop.Checked
    'Set the new ontop value
    Call OnTop(Me.hWnd, mnuFileWindowOnTop.Checked)
End Sub

Private Sub mnuFileWindowTrayIcon_Click()
    On Local Error Resume Next
    'Load the tray icon, everything is taken care of in it's Form_Load proc
    Load frmTray
End Sub

Private Sub mnuHelpAbout_Click()
    On Local Error Resume Next
    'Show the about dialog
    Load frmAbout
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuFileNew_Click()
    On Local Error Resume Next
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Call NewRTFText(strCodeFilename)
        Case "notes"
            Call NewRTFText(strNotesFilename)
    End Select
End Sub

Private Sub mnuFileOpen_Click()
    On Local Error Resume Next
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Call LoadRTFText(strCodeFilename)
        Case "notes"
            Call LoadRTFText(strNotesFilename)
    End Select
End Sub

Private Sub mnuFilePrint_Click()
    On Local Error Resume Next
    Call PrintRTFText
End Sub

Private Sub mnuFileSave_Click()
    On Local Error Resume Next
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Call SaveRTFText(strCodeFilename)
        Case "notes"
            Call SaveRTFText(strNotesFilename)
    End Select
End Sub

Private Sub mnuFileSaveAs_Click()
    On Local Error Resume Next
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Call SaveAsRTFText(strCodeFilename)
        Case "notes"
            Call SaveAsRTFText(strNotesFilename)
    End Select
End Sub

Private Sub mnuHelpChangeLog_Click()
    On Local Error Resume Next
    Call ShellExecute(Me.hWnd, "open", FixPath(App.Path) & "Change-Log.txt", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub mnuHelpHelp_Click()
    On Local Error Resume Next
    'Add a \ to the path if needed
    Dim strAppPath As String
    strAppPath = FixPath(App.Path)
    If DoesFileExist(strAppPath & "Help\index.htm") Then
        Call ShellExecute(Me.hWnd, vbNullString, strAppPath & "Help\index.htm", _
            vbNullString, vbNullString, vbNormalFocus)
    Else
        Call MsgBox("Help files do not exist!", vbExclamation Or vbOKOnly, "Not Found")
    End If
End Sub

Private Sub mnuHelpLinksAdd_Click()
    Dim strAddress As String
    strAddress = InputBox("Enter the address of the web site you would like to add:" & vbNewLine, _
        "Add Link", "http://www.rickbull.com/")
    If strAddress <> vbNullString Then
        Call SaveText(FixPath(App.Path) & "Site List.txt", vbNewLine & strAddress, Add)
        If mnuHelpLinksVBSites(0).Enabled = True Then Load mnuHelpLinksVBSites(mnuHelpLinksVBSites.UBound + 1)
        With mnuHelpLinksVBSites(mnuHelpLinksVBSites.UBound)
            .Caption = strAddress
            .Enabled = True
            .Visible = True
        End With
    End If
End Sub

Private Sub mnuHelpLinksManage_Click()
    On Local Error Resume Next
    Load frmManageLinks
    frmManageLinks.Show vbModal, Me
End Sub

Private Sub mnuHelpLinksVBSites_Click(Index As Integer)
    On Local Error Resume Next
    'Execute the address
    Call ShellExecute(Me.hWnd, vbNullString, mnuHelpLinksVBSites(Index).Caption, _
        vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub mnuHelpThanks_Click()
    On Local Error Resume Next
    Call ShellExecute(Me.hWnd, "open", FixPath(App.Path) & "Thanks.txt", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub mnuHelpTip_Click()
    On Local Error Resume Next
    Load frmTip
    frmTip.Show vbModal, Me
End Sub

Private Sub mnuToolsExplorer_Click()
    On Local Error Resume Next
    'Show explorer with folder pane (/e)
    Call Shell("EXPLORER.exe /e, " & App.Path, vbNormalFocus)
End Sub

Private Sub mnuEditFormatFormatAsEditor_Click()
    On Local Error Resume Next
    'Load the editor
    Load frmFormattingEditor
    frmFormattingEditor.Show vbModal, Me
End Sub

Private Sub mnuEditFormatFormatAsOption_Click(Index As Integer)
    On Local Error Resume Next
    Dim bolSuccess As Boolean 'Whether the code has been formatted
    Dim rtfSelected As RichTextBox
    
    'Choose the selected tab
    Select Case LCase(tbsView.SelectedItem.Key)
        'Get the correct RTFBox
        Case "code"
            Set rtfSelected = rtfCode
            
        Case "notes"
            Set rtfSelected = rtfNotes
            
        'Not a RTFBox
        Case Else
            'Tell user they can't format this
            Call MsgBox("You can not format this. Please select either the Code or the Notes window.", vbExclamation Or vbOKOnly, "Error")
            'Exit so as not to do below
            Exit Sub
    End Select
    
    Dim strLastCaption As String
    strLastCaption = sbrInfo.Panels("Info").Text
    sbrInfo.Panels("Info").Text = "Formatting Text - Press Esc to Cancel"
    With rtfSelected
        'If not selected
        If .SelLength = 0 Then
            'Select all
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
        'Copy the text to be copied to the Temp RTF Box
        rtfTemp.TextRTF = .SelRTF
        'Format the text, and get whether it was successful
        bolSuccess = FormatText(rtfTemp, strFormattingPath & mnuEditFormatFormatAsOption(Index).Caption)
        'If successfull
        If bolSuccess Then
            'Put the new text in the RTFBox
            .SelRTF = rtfTemp.TextRTF
        'If failed/aborted
        Else
            'Tell the user
            Call MsgBox("Formatting was aborted by user or an error has occured!", vbExclamation Or vbOKOnly, "Formatting Cacelled")
        End If
    End With
    sbrInfo.Panels("Info").Text = strLastCaption
End Sub

Private Sub mnuToolsFindInFiles_Click()
    Load frmFindInFiles
    frmFindInFiles.Show vbModal, Me
End Sub

Private Sub mnuToolsMyComputer_Click()
    On Local Error Resume Next
    'Show explorer withOUT folder pane (My Coputer)
    Call Shell("EXPLORER.exe " & App.Path, vbNormalFocus)
End Sub

Private Sub mnuToolsOptions_Click()
    On Local Error Resume Next
    'Save the options so we can get the newest ones in the options dialog
    Call SaveOptions
    'Show the about dialog
    Load frmOptions
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuTreeCodeCount_Click()
    On Local Error Resume Next
    Dim cntDetails As typCount
    cntDetails = GetCount
    'Show the code count
    Call MsgBox("Number of Codes Found: " & cntDetails.Codes & vbNewLine & _
        "Number of Sections Found: " & cntDetails.Sections, _
        vbInformation Or vbOKOnly, "Code Count")
End Sub

Private Sub mnuTreeCollapseAll_Click()
    On Local Error Resume Next
    'Get the current status bar text
    Dim LastCaption As String
    LastCaption = sbrInfo.Panels("Info").Text
    'Change it to tell the user what's happening
    sbrInfo.Panels("Info").Text = "Collapsing nodes - Esc to cancel"
    
    'Loop for all nodes
    Dim LoopCounter As Integer
    LoopCounter = tvwCodes.Nodes.Count
    Do While LoopCounter > 0
        'Allow updates
        DoEvents
        'If user is pressing escape quit loop
        If IsButtonActive(VK_ESCAPE, Me.hWnd) Then Exit Do
        'Collapse node
        tvwCodes.Nodes(LoopCounter).Expanded = False
        'Decrement counter to move back in the nodes
        LoopCounter = LoopCounter - 1
    Loop
    
    'Reset the original status caption
    sbrInfo.Panels("Info").Text = LastCaption
End Sub

Private Sub mnuTreeDeleteitem_Click()
    On Local Error Resume Next
    'Delete the item
    Call FileOperation(IIf(Right(tvwCodes.SelectedItem.Key, 1) = "\", _
        Left(tvwCodes.SelectedItem.Key, _
        Len(tvwCodes.SelectedItem.Key) - 1), _
        tvwCodes.SelectedItem.Key), , FO_DELETE)
    'Show the changes
    Call mnuTreeRefresh_Click
End Sub

Private Sub mnuTreeExpandAll_Click()
    On Local Error Resume Next
    'Get the current status bar text
    Dim LastCaption As String
    LastCaption = sbrInfo.Panels("Info").Text
    'Change it to tell the user what's happening
    sbrInfo.Panels("Info").Text = "Expanding nodes - Esc to cancel"
    
    'Loop for all nodes
    Dim LoopCounter As Integer
    LoopCounter = 1
    Do While LoopCounter < tvwCodes.Nodes.Count
        'Allow updates
        DoEvents
        'If user is pressing escape quit loop
        If IsButtonActive(VK_ESCAPE, Me.hWnd) Then Exit Do
        'Expand node
        tvwCodes.Nodes(LoopCounter).Expanded = True
        'Increment counter to move back in the nodes
        LoopCounter = LoopCounter + 1
    Loop
    
    'Reset the original status caption
    sbrInfo.Panels("Info").Text = LastCaption
End Sub

Private Sub mnuTreeFindCodes_Click()
    On Local Error Resume Next
    'Loop for all nodes
    Dim lngLoopCounter As Long
    Do While lngLoopCounter < tvwCodes.Nodes.Count
        'Increment the loop counter
        lngLoopCounter = lngLoopCounter + 1
        'If not a folder/section (i.e. a code) make it visivle
        If Right(tvwCodes.Nodes(lngLoopCounter).Key, Len("\")) <> "\" Then _
            tvwCodes.Nodes(lngLoopCounter).EnsureVisible
    Loop
End Sub

Private Sub mnuTreeHideShow_Click()
    On Local Error Resume Next
    'How much to move the tree in each loop & wait interval
    Const intAddAmount As Integer = 90
    Dim lngTime As Long 'Long for time stuff

    'We want to Show the tree
    If mnuTreeHideShow.Caption = "S&how" Then
        'Do while we still need to bring tree out
        Do While picTreeSizer.Left < tvwCodes.Width
            DoEvents
            'Get the start time
            lngTime = GetTickCount
            'Move the Sizer and Tree right
            picTreeSizer.Left = picTreeSizer.Left + intAddAmount
            tvwCodes.Left = picTreeSizer.Left - tvwCodes.Width
            'Update elements positions
            Call Form_Resize
            'Get the Current Time - Start time, and work out how much of the
            'interval we need to use, this keeps it running similar on diff. speed PCs
            lngTime = intTreeTimerInterval - (GetTickCount - lngTime)
            'Wait so as we get an animation effect
            If lngTime > 0 And lngTime <= intTreeTimerInterval Then Call Wait(lngTime)
        Loop
        'Set the caption to menu Hide
        mnuTreeHideShow.Caption = "&Hide"
        'Put the sizer's mouse pointer back to the sizer, so the tree can be sized again
        picTreeSizer.MousePointer = vbCustom
    
    'We want to Hide the tree
    Else
        'Put the sizer's mouse pointer to default,
        'so the tree cannot be sized until it is shown from the same menu
        picTreeSizer.MousePointer = vbDefault
        'Set the caption to menu Show
        mnuTreeHideShow.Caption = "S&how"
        'Do while we still need to move the tree in
        Do While picTreeSizer.Left + picTreeSizer.Width - intAddAmount > 0
            DoEvents
            'Get the start time
            lngTime = GetTickCount
            'Move the Sizer and Tree left
            picTreeSizer.Left = picTreeSizer.Left - intAddAmount
            tvwCodes.Left = picTreeSizer.Left - tvwCodes.Width
            'Update elements positions
            Call PositionElements
            'Get the Current Time - Start time, and work out how much of the
            'interval we need to use, this keeps it running similar on diff. speed PCs
            lngTime = intTreeTimerInterval - (GetTickCount - lngTime)
            'Wait so as we get an animation effect
            If lngTime > 0 And lngTime <= intTreeTimerInterval Then Call Wait(lngTime)
        Loop
    End If
End Sub

Private Sub mnuTreeNew_Click()
    Dim strDir As String, strParent
    'Get the new folder name
    strDir = InputBox("Please enter the name of the new section:", "Add Section")
    If strDir <> vbNullString Then
        'Append the current folder to the start of this one
        If Right(tvwCodes.SelectedItem.Key, 1) = "\" Then
            strParent = tvwCodes.SelectedItem.Key
        Else
            strParent = GetParentDir(tvwCodes.SelectedItem.Key)
        End If
        'Make the new directory, copy bitmaps and show the changes
        Call MkDir(strParent & strDir)
        Call FileOperation(strParent & "Large.bmp", strParent & strDir, FO_COPY, FOF_SILENT)
        Call FileOperation(strParent & "Small.bmp", strParent & strDir, FO_COPY, FOF_SILENT)
        Call mnuTreeRefresh_Click
    End If
End Sub

Private Sub mnuTreeRefresh_Click()
    On Local Error Resume Next
    'Refresh the lists
    Call ValidateCodes
    Call ValidateBookmarks
End Sub

Private Sub mnuTreeRenameitem_Click()
    On Local Error Resume Next
    'Find if we have a directory
    Dim bolDir As Boolean
    bolDir = Right(tvwCodes.SelectedItem.Key, 1) = "\"
    'Get the old file/folder name (removing trailing "\" if folder)
    Dim strOld As String
    strOld = IIf(bolDir = True, _
        Left(tvwCodes.SelectedItem.Key, Len(tvwCodes.SelectedItem.Key) - 1), _
        tvwCodes.SelectedItem.Key)
    'Get the new filename
    Dim strNew As String
    strNew = InputBox("Please enter the new file-name:", "Rename", tvwCodes.SelectedItem.Text)

    'If we have a new name
    If strNew <> vbNullString Then
        Dim intFound As Integer
        'If a file, add the extension
        If bolDir = False Then
            intFound = InStrRev(strOld, ".")
            If intFound > 0 Then strNew = strNew & Right(strOld, Len(strOld) - intFound + 1)
        End If
        'Add the folder name to the start of the new name
        intFound = InStrRev(strOld, "\")
        strNew = Left(strOld, intFound) & strNew
        'Rename the file and show changes
        Call FileOperation(strOld, strNew, FO_RENAME)
        Call mnuTreeRefresh_Click
    End If
End Sub

Private Sub mnuTreeSaveitem_Click()
    On Local Error Resume Next
    Dim strFileName As String
    'If it's a directory
    If Right(tvwCodes.SelectedItem.Key, 1) = "\" Then
        'Get the new FOLDER name, and copy it to the new locoation
        strFileName = GetFolder(Me.hWnd)
        If strFileName <> vbNullString Then _
        Call FileOperation(Left(tvwCodes.SelectedItem.Key, _
            Len(tvwCodes.SelectedItem.Key) - 1), _
                strFileName, FO_COPY)
    Else
        'Get the new FILE name, and copy it to the new locoation
        strFileName = GetFileName(, True)
        If strFileName <> vbNullString Then _
            Call FileOperation(tvwCodes.SelectedItem.Key, _
                strFileName, FO_COPY)
    End If
End Sub

Private Sub mnuView_Click()
    On Local Error Resume Next
    'Show/hide the customize menus
    mnuViewToolbarsSeperator2.Visible = False
    mnuViewToolbarsCustomizeStandard.Visible = False
    mnuViewToolbarsCustomizeFormatting.Visible = False
End Sub

Private Sub mnuViewStatusBar_Click()
    On Local Error Resume Next
    'Invert the check
    sbrInfo.Visible = Not sbrInfo.Visible
    'Show/hide the staus bar
    mnuViewStatusBar.Checked = sbrInfo.Visible
    'Show the changes
    Call PositionElements
End Sub

Private Sub mnuViewToolbarsCustomizeFormatting_Click()
    On Local Error Resume Next
    'Show the customize menu
    Call tbrFormatting.Customize
End Sub

Private Sub mnuViewToolbarsCustomizeStandard_Click()
    On Local Error Resume Next
    'Show the customize menu
    Call tbrStandard.Customize
End Sub

Private Sub mnuViewToolbarsFormatting_Click()
    On Local Error Resume Next
    'Invert the check
    mnuViewToolbarsFormatting.Checked = Not mnuViewToolbarsFormatting.Checked
    'Show/hide the menu
    tbrFormatting.Visible = mnuViewToolbarsFormatting.Checked And _
        (LCase(tbsView.SelectedItem.Key) = "code" Or LCase(tbsView.SelectedItem.Key) = "notes")
    'Show the changes
    Call PositionElements
End Sub

Private Sub mnuViewToolbarsShowHide_Click(Index As Integer)
    On Local Error Resume Next
    'Get the visibility of the menus (0 = True, 1 = False)
    Dim bolTrueFalse As Boolean
    bolTrueFalse = (Index = 0)
    'Set the menus items to the correct state
    mnuViewToolbarsStandard.Checked = bolTrueFalse
    mnuViewToolbarsFormatting.Checked = bolTrueFalse
    'Show/hide the menus
    tbrStandard.Visible = bolTrueFalse
    tbrFormatting.Visible = bolTrueFalse
    'Show the changes
    Call PositionElements
End Sub

Private Sub mnuViewToolbarsStandard_Click()
    On Local Error Resume Next
    'Invert the check
    tbrStandard.Visible = Not tbrStandard.Visible
    'Show/hide the menu
    mnuViewToolbarsStandard.Checked = tbrStandard.Visible
    'Show the changes
    Call PositionElements
End Sub

Private Sub picTreeSizer_DblClick()
    On Local Error Resume Next
    'Show the tree if it's shown
    If mnuTreeHideShow.Caption = "&Hide" Then Call mnuTreeHideShow_Click
End Sub

Private Sub picTreeSizer_KeyPress(KeyAscii As Integer)
    On Local Error Resume Next
    'If escape
    If KeyAscii = vbKeyEscape Then
        'Set left mouse button = up
        Call SetKeyState(VK_LBUTTON, 0)
        'Put the size back where it was
        picTreeSizer.Left = tvwCodes.Left + tvwCodes.Width
        'Reset the appropriate stuff
        Call picTreeSizer_MouseUp(vbLeftButton, 0, 0, 0)
    End If
End Sub

Private Sub picTreeSizer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    'If mouse button is the left one and the tree is shown
    If Button = vbLeftButton And mnuTreeHideShow.Caption = "&Hide" Then
        'Get the down at X point, so we can take this into account when moving the picture box
        sngOffsetX = X
        'Change the colour to show we are in "move mode"
        picTreeSizer.BackColor = vb3DShadow
    End If
End Sub

Private Sub picTreeSizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    'If mouse button is the left one and the tree is shown and we have enough room to move to the specified position
    If Button = vbLeftButton And mnuTreeHideShow.Caption = "&Hide" And _
        (picTreeSizer.Left + X - sngOffsetX) > TwipsX(intMinWidth) And _
        (picTreeSizer.Left + X - sngOffsetX) < ScaleWidth - TwipsX(intMinWidth) Then
        'Move the sizer to the specified position
        picTreeSizer.Left = picTreeSizer.Left + X - sngOffsetX
    End If
End Sub

Private Sub picTreeSizer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    'If mouse button is the left one and the tree is shown
    If Button = vbLeftButton And mnuTreeHideShow.Caption = "&Hide" Then
        'Reset the back colour
        picTreeSizer.BackColor = Me.BackColor
        'Move the code tree to the correct position
        tvwCodes.Move 0, tvwCodes.Top, picTreeSizer.Left
        'Position the form's elements accordingly
        Call PositionElements
    End If
End Sub

Private Sub rtfCode_Change()
    On Local Error Resume Next
    With rtfCode
        'Add this text to the undo buffer
        Call clsCodeUndo.AddToBuffer(.TextRTF, .SelStart, .SelLength)
    End With
End Sub

Private Sub rtfCode_KeyDown(KeyCode As Integer, Shift As Integer)
    'If tab
    If KeyCode = vbKeyTab Then
        'Don't change focus
        KeyCode = 0
        'Make selected text a tab
        If rtfCode.SelLength = 0 Then
            'Make selected text a tab
            rtfCode.SelText = vbTab
        Else
            'Indent the selected text
            Call rtfCode_KeyPress(vbKeyTab)
        End If
    End If
    Call RTFKeyDown(KeyCode, Shift)
End Sub

Private Sub rtfCode_KeyPress(KeyAscii As Integer)
    'Tab
    If KeyAscii = vbKeyTab And IsButtonActive(VK_LSHIFT, rtfCode.hWnd) = False Then
        'Stop the new character being added to the text
        KeyAscii = 0
        'Indent the text
        Call Dent(rtfCode)
    'Backspace
    ElseIf KeyAscii = vbKeyTab And IsButtonActive(VK_LSHIFT, rtfCode.hWnd) = True Then
        KeyAscii = 0
        'Outdent the text
        Call Dent(rtfCode, False)
    End If
End Sub

Private Sub rtfNotes_KeyDown(KeyCode As Integer, Shift As Integer)
    'If tab
    If KeyCode = vbKeyTab Then
        'Don't change focus
        KeyCode = 0
        If rtfNotes.SelLength = 0 Then
            'Make selected text a tab
            rtfNotes.SelText = vbTab
        Else
            'Indent the selected text
            Call rtfNotes_KeyPress(vbKeyTab)
        End If
    End If
End Sub

Private Sub rtfNotes_KeyPress(KeyAscii As Integer)
    'Tab
    If KeyAscii = vbKeyTab And IsButtonActive(VK_LSHIFT, rtfNotes.hWnd) = False Then
        'Stop the new character being added to the text
        KeyAscii = 0
        'Indent the text
        Call Dent(rtfNotes)
    'Backspace
    ElseIf KeyAscii = vbKeyTab And IsButtonActive(VK_LSHIFT, rtfNotes.hWnd) = True Then
        KeyAscii = 0
        'Outdent the text
        Call Dent(rtfNotes, False)
    End If
End Sub

Private Sub rtfCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call mnuEdit_Click
        Call PopupMenu(mnuEdit)
    End If
End Sub

Private Sub rtfCode_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    'Load it
    rtfCode.Filename = Data.Files(1)
    'Set it's filename to the var
    strCodeFilename = Data.Files(1)
    'Reset the undos
    Call clsCodeUndo.Reset
End Sub

Private Sub rtfCode_SelChange()
    On Local Error Resume Next
    'Set the char/line number and stuff
    Call SelectionChange
End Sub

Private Sub rtfNotes_Change()
    On Local Error Resume Next
    With rtfNotes
        'Add this text to the undo buffer
        Call clsNotesUndo.AddToBuffer(.TextRTF, .SelStart, .SelLength)
    End With
End Sub

Private Sub rtfNotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        'SHow the Edit popup menu
        Call mnuEdit_Click
        Call PopupMenu(mnuEdit)
    End If
End Sub

Private Sub rtfNotes_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    'Load it
    rtfNotes.Filename = Data.Files(1)
    'Set it's filename to the var
    strNotesFilename = Data.Files(1)
    'Reset the undos
    Call clsNotesUndo.Reset
End Sub

Private Sub rtfNotes_SelChange()
    On Local Error Resume Next
    'Set the char/line number and stuff
    Call SelectionChange
End Sub

Private Sub sbrInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    'Show the view menu if right button
    If Button = vbRightButton Then Call PopupMenu(mnuView)
End Sub

Private Sub sbrInfo_PanelDblClick(ByVal Panel As MSComctlLib.Panel)
    Select Case LCase(Panel.Key)
        Case "ins"
            Call SetKeyState(VK_INSERT)
        Case "caps"
            Call SetKeyState(VK_CAPITAL)
        Case "num"
            Call SetKeyState(VK_NUMLOCK)
        Case "scrl"
            Call SetKeyState(VK_SCROLL)
    End Select
    sbrInfo.Refresh
End Sub

Private Sub tbrFormatting_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Local Error Resume Next
    Dim rtfSelected As RichTextBox
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = rtfCode
        Case "notes"
            Set rtfSelected = rtfNotes
        Case Else
            Exit Sub
    End Select
    With rtfSelected
        Select Case LCase(Button.Key)
            Case "bold"
                .SelBold = BooleanPressed(Button.Value)
            Case "italic"
                .SelItalic = BooleanPressed(Button.Value)
            Case "underline"
                .SelUnderline = BooleanPressed(Button.Value)
            Case "strike-thru"
                .SelStrikeThru = BooleanPressed(Button.Value)
            Case "left"
                .SelAlignment = vbLeftJustify
            Case "center"
                .SelAlignment = vbCenter
            Case "right"
                .SelAlignment = vbRightJustify
            Case "font colour"
                If Button.Value = tbrPressed Then
                    With frmColourPicker
                        Load frmColourPicker
                        .strOpener = Me.Name
                        Call .PositionForm
                        If Not IsNull(rtfSelected.SelColor) Then
                            Call .SelectColour(rtfSelected.SelColor)
                        Else
                            Call .SelectColour(-1)
                        End If
                        .Show vbModeless, Me
                    End With
                Else
                    Unload frmColourPicker
                End If
            Case "bulletted list"
                .SelBullet = BooleanPressed(Button.Value)
            Case "indent"
                Call mnuEditFormatIndent_Click
            Case "outdent"
                Call mnuEditFormatOutdent_Click
            Case "character map"
                If Button.Value = tbrPressed Then
                    Load frmCharacterMap
                    frmCharacterMap.Show vbModeless, Me
                Else
                    Unload frmCharacterMap
                End If
            Case "change case"
                Call tbrFormatting_ButtonMenuClick(tbrFormatting.Buttons("Change Case").ButtonMenus(tbrFormatting.Buttons("Change Case").Tag))
        End Select
    End With
End Sub

Private Sub tbrFormatting_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case LCase(ButtonMenu.Key)
        Case "lower case", "upper case", "toggle case", "proper case", _
            "sentance case", "vary case 1", "vary case 2"
            Select Case LCase(tbsView.SelectedItem.Key)
                Case "code"
                    rtfCode.SelText = ChangeCase(rtfCode.SelText, ButtonMenu.Index)
                Case "notes"
                    rtfNotes.SelText = ChangeCase(rtfNotes.SelText, ButtonMenu.Index)
                Case Else
                     Call MsgBox("You cannot format this. Please select either the Code or the Notes window.", _
                        vbExclamation Or vbOKOnly, "Error")
                    Exit Sub
            End Select
        With tbrFormatting.Buttons("Change Case")
            .Tag = ButtonMenu.Key
            .ToolTipText = "Change Case - " & ButtonMenu.Key
        End With
    End Select
End Sub

Private Sub tbrFormatting_Change()
    Call PositionElements
End Sub

Private Sub tbrFormatting_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    If Button = vbRightButton Then
        'Show/hide the correct menus
        mnuViewToolbarsSeperator2.Visible = True
        mnuViewToolbarsCustomizeStandard.Visible = False
        mnuViewToolbarsCustomizeFormatting.Visible = True
        'Popup the menu
        Call PopupMenu(mnuViewToolbars, , , , mnuViewToolbarsCustomizeFormatting)
    End If
End Sub

Private Sub tbrStandard_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Local Error Resume Next
    Dim strTabKey As String 'The selected tab's key in lowercase
    strTabKey = LCase(tbsView.SelectedItem.Key)
    
    'Choose the button's key (in lower case for ease of comparrision/less errors)
    Select Case LCase(Button.Key)
        'New
        Case "new"
            Call mnuFileNew_Click
        'Open
        Case "open"
            Call mnuFileOpen_Click
        'Save
        Case "save"
            Call mnuFileSave_Click

        'SaveAs
        Case "save as"
            Call mnuFileSaveAs_Click
            
        Case "save all"
            'Call mnuFileSaveAllClick
            
        'Save
        Case "print"
            Call mnuFilePrint_Click
        
        '---------------
        'Undo
        Case "undo"
            Call mnuEditUndo_Click
        'Save
        Case "redo"
            Call mnuEditRedo_Click
        'Cut
        Case "cut"
            Call mnuEditCut_Click
        'Copy
        Case "copy"
            Call mnuEditCopy_Click
        'Paste
        Case "paste"
            Call mnuEditPaste_Click
        'Find
        Case "find"
            Call mnuEditFind_Click
        'Add to bookmarks
        Case "add to bookmarks"
            Call mnuBookmarksAdd_Click
        'Copy to notes
        Case "copy to notes"
            With rtfNotes
                'Set the selected start in the notes to the end
                .SelStart = Len(.Text)
                'Add the seperator string
                .SelText = GetSeperator
                'Add the selected code
                .SelRTF = rtfCode.SelRTF
            End With
            'Show the notes window if wanted
            If bolAutoShowNotes Then
                tbsView.Tabs("Notes").Selected = True
                Call tbsView_Click
            End If
    End Select
End Sub

Private Sub tbrStandard_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Local Error Resume Next
    'Choose the button dropdown menu's key (in lower case for ease of comparrision/less errors)
    Select Case LCase(ButtonMenu.Key)
        'COPY TO NOTES
        'Copy Selected
        Case "copy selected"
            With rtfNotes
                'Set the selected start in the notes to the end
                .SelStart = Len(.Text)
                'Add the seperator string
                .SelText = GetSeperator
                'Add the selected code
                .SelRTF = rtfCode.SelRTF
            End With
        'Copy All
        Case "copy all"
            With rtfNotes
                'Set the selected start in the notes to the end
                .SelStart = Len(.Text)
                'Add the seperator string
                .SelText = GetSeperator
                'Add all of the code
                .SelRTF = rtfCode.TextRTF
            End With
    End Select
    'Show the notes window if wanted
    If bolAutoShowNotes Then
        tbsView.Tabs("Notes").Selected = True
        Call tbsView_Click
    End If
End Sub

Private Sub tbrStandard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    If Button = vbRightButton Then
        'Show/hide the correct menus
        mnuViewToolbarsSeperator2.Visible = True
        mnuViewToolbarsCustomizeStandard.Visible = True
        mnuViewToolbarsCustomizeFormatting.Visible = False
        'Popup the menu
        Call PopupMenu(mnuViewToolbars, , , , mnuViewToolbarsCustomizeStandard)
    End If
End Sub

Private Sub tbsView_Click()
    On Local Error Resume Next
    Call tbsView_MouseUp(vbLeftButton, 0, 0, 0)
End Sub

Private Sub tbsView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    'Choose the tab's key (in lower case for ease of comparrision/less errors)
    Select Case LCase(tbsView.SelectedItem.Key)
        'Show the current section and hide the others:
        'Code
        Case "code"
            rtfCode.Visible = True
            rtfNotes.Visible = False
            lvwBookmarks.Visible = False
            tbrFormatting.Visible = mnuViewToolbarsFormatting.Checked
        'Notes
        Case "notes"
            rtfCode.Visible = False
            rtfNotes.Visible = True
            lvwBookmarks.Visible = False
            tbrFormatting.Visible = mnuViewToolbarsFormatting.Checked
        'Bookmarks
        Case "bookmarks"
            rtfCode.Visible = False
            rtfNotes.Visible = False
            lvwBookmarks.Visible = True
            tbrFormatting.Visible = False
    End Select
    'Position the elements
    Call PositionElements
    'Set the char/line number and stuff
    Call SelectionChange
    'Set the save button's enabled value
    Call SetSaveButton
End Sub

Private Sub tvwCodes_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Local Error Resume Next
    'RENAME FILE
End Sub

Public Sub tvwCodes_Expand(ByVal Node As MSComctlLib.Node)
    On Local Error Resume Next
    'Add the files/sections for this sections (if they already exist they won't be overwritten)
    'Call AddNodes(tvwCodes, Node.Key, strPattern)
    If bolCheckCodesOnExpand Then Call FindFiles(strCodePath, Me, "AddCodes")
End Sub

Private Sub LoadCode(ByVal Node As MSComctlLib.Node)
    On Local Error GoTo ErrorHandler
    Dim bolFileLoaded As Boolean 'Whether the code has been loaded - for ErrorHandler
    'Code hasn't been loaded
    bolFileLoaded = False

    With rtfCode
        'Load the file
        .Filename = Node.Key
        'File has now been loaded
        bolFileLoaded = True
        'Set the filename
        strCodeFilename = Node.Key
        'Select the start otherwise position may be at the end of the file
        .SelStart = 1
    End With
    'Reset the undos
    Call clsCodeUndo.Reset
    'Set details about the current file
    sbrInfo.Panels("Info").Text = _
        "Title: " & Node.Text & _
        "   Date: " & FileDateTime(Node.Key)
    
    'Show the bookmark in the code window if wanted
    If bolAutoShowCode Then
        tbsView.SelectedItem = tbsView.Tabs("Code")
        Call tbsView_MouseUp(vbLeftButton, 0, 0, 0)
    End If
    'Set the char/line number and stuff
    Call SelectionChange
    
    'Un/Bold the menu in the bookmarks
    Call BoldBookmark(Node.Key)
    'Exit so as not to cause an unjustified error msgbox
    Exit Sub
ErrorHandler:
    'Tell the user of the error only if the file has not been loaded
    If bolFileLoaded = False Then MsgBox "The file could not be loaded:" & vbNewLine _
        & Err.Number & " - " & Err.Description, vbCritical Or vbOKOnly, "Error"
End Sub

Private Sub PositionElements()
    On Local Error Resume Next

    'Move the Tree Sizer to _
        Left: Same _
        Top: Below the toolbar _
        Width: Same _
        Height: between toolbar and status bar (if they are visible)
    picTreeSizer.Move picTreeSizer.Left, IIf(tbrStandard.Visible, tbrStandard.Top + tbrStandard.Height, 0), _
        picTreeSizer.Width, ScaleHeight - IIf(tbrStandard.Visible, tbrStandard.Top + tbrStandard.Height, 0) - _
        IIf(sbrInfo.Visible, sbrInfo.Height, 0)
    'Move Tree to _
        Left: Same _
        Top: Sizer's Top _
        Width: Sizer's Left - Tree's Left _
        Height: Sizer's height
    tvwCodes.Move tvwCodes.Left, picTreeSizer.Top, _
        picTreeSizer.Left - tvwCodes.Left, picTreeSizer.Height
    With tbsView
        'Move tabs to _
            Left: Past Sizer , _
            Top: Sizer's Top, _
            Width: Between Sizer and Edge of Form, _
            Height: Sizer's width - a bit
        .Move picTreeSizer.Left + picTreeSizer.Width, picTreeSizer.Top, _
            Me.ScaleWidth - (picTreeSizer.Left + picTreeSizer.Width) - TwipsX(2), _
            picTreeSizer.Height - TwipsY(2)
        'Position the Formatting toolbar
        tbrFormatting.Move .ClientLeft, .ClientTop, .ClientWidth
        cboFont.Move tbrFormatting.Buttons("Font").Left + TwipsX, _
            tbrFormatting.Buttons("Font").Top + TwipsY
        cboFontSize.Move tbrFormatting.Buttons("Font Size").Left + TwipsX, _
            tbrFormatting.Buttons("Font Size").Top + TwipsY
        'Position the RTF boxes to the visible area of the tab strip + _
         the formatting toolbar if visible
        rtfCode.Move .ClientLeft, .ClientTop + IIf(tbrFormatting.Visible, tbrFormatting.Height, 0), _
            .ClientWidth, .ClientHeight - IIf(tbrFormatting.Visible, tbrFormatting.Height, 0)
        'Same as rtfCode - more efficent that recalculating values
        rtfNotes.Move rtfCode.Left, rtfCode.Top, rtfCode.Width, rtfCode.Height
        'Position the bookmarks to the visible area of the tab strip
        lvwBookmarks.Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
    End With
End Sub

Private Sub tvwCodes_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            Call mnuTreeDeleteitem_Click
    End Select
End Sub

Private Sub tvwCodes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Call PopupMenu(mnuTree)
End Sub

Public Sub tvwCodes_NodeClick(ByVal Node As MSComctlLib.Node)
    On Local Error Resume Next
    'If we have a file load it
    If Right(Node.Key, Len("\")) <> "\" Then Call LoadCode(Node)
End Sub

Private Sub SelectionChange()
    On Local Error Resume Next
    Dim rtfSelected As RichTextBox
    'Choose the tabstrip's selected item's key in lower case (for ease of comparission)
    Select Case LCase(tbsView.SelectedItem.Key)
        'Set the status bar's text according to which window is visible
        'Code
        Case "code"
            Set rtfSelected = rtfCode
                 
        'Notes
        Case "notes"
            Set rtfSelected = rtfNotes
            
        'Bookmarks
        Case "bookmarks"
            Exit Sub
    End Select
    With rtfSelected
        'Selected Char / Total Chars _
        Selected Line / Total Line
        sbrInfo.Panels("Position").Text = _
            "Char: " & .SelStart + 1 & " / " & Len(.Text) + 1 & _
            "   Line: " & SendMessage(.hWnd, EM_LINEFROMCHAR, .SelStart, 0&) + 1 & _
            " / " & SendMessage(rtfCode.hWnd, EM_GETLINECOUNT, 0&, 0&)
        
        'Standard toolbar & Menus
        tbrStandard.Buttons("Cut").Enabled = .SelLength > 0
        mnuEditCut.Enabled = tbrStandard.Buttons("Cut").Enabled
        tbrStandard.Buttons("Copy").Enabled = .SelLength > 0
        mnuEditCopy.Enabled = tbrStandard.Buttons("Copy").Enabled
        mnuEditDelete.Enabled = tbrStandard.Buttons("Copy").Enabled

        'Formatting Toolbar
        cboFont.Text = .SelFontName
        cboFontSize.Text = .SelFontSize
        tbrFormatting.Buttons("Bold").Value = Pressed(.SelBold)
        tbrFormatting.Buttons("Italic").Value = Pressed(.SelItalic)
        tbrFormatting.Buttons("Underline").Value = Pressed(.SelUnderline)
        tbrFormatting.Buttons("Strike-Thru").Value = Pressed(.SelStrikeThru)
        tbrFormatting.Buttons("Left").Value = Pressed(.SelAlignment = vbLeftJustify)
        tbrFormatting.Buttons("Center").Value = Pressed(.SelAlignment = vbCenter)
        tbrFormatting.Buttons("Right").Value = Pressed(.SelAlignment = vbRightJustify)
        tbrFormatting.Buttons("Bulletted List").Value = Pressed(.SelBullet)
    End With
End Sub

Private Function GetSection(Optional ByVal Index = Null, _
    Optional ByVal ReplacePathSeperators As Boolean = True, _
    Optional ByVal ReplaceString As String = " » ") As String
    On Local Error Resume Next
    
    'Set index to the selected item in the code tree if not specified
    If Index = Null Then Index = tvwCodes.SelectedItem.Index
    'Return the current key for now - also more efficient
    GetSection = tvwCodes.Nodes(Index).Key
    'If we are not in a folder/section
    If Right(GetSection, Len("\")) <> "\" Then
        'Find the end of the section
        Dim Found As Long
        Found = InStrRev(GetSection, "\", , vbTextCompare)
        'Return the section
        If Found > -1 Then GetSection = Left(GetSection, Found)
    End If
    'Remove the last \ if present
    If Right(GetSection, Len("\")) = "\" Then GetSection = Left(GetSection, Len(GetSection) - 1)
    'Remove the code path from the beginning of the section, as that is irrelevant data
    GetSection = Right(GetSection, Len(GetSection) - Len(strCodePath))
    'If the \ are to be replaced (default =  » ) replace them
    If ReplacePathSeperators Then GetSection = Replace(GetSection, "\", ReplaceString, , , vbTextCompare)
End Function

Private Sub BoldBookmark(ByVal Key)
    On Local Error Resume Next
    'If the last item to be bolded exists in the bookmarks remove it's boldness
    If DoesListItemExist(lvwBookmarks, strLastOn) Then lvwBookmarks.ListItems(strLastOn).Bold = False
    'If the new one exists in the bookmars
    If DoesListItemExist(lvwBookmarks, Key) Then
        'Make it bold
        lvwBookmarks.ListItems(Key).Bold = True
        'Set the last on to this new one
        strLastOn = Key
    End If
End Sub

Public Sub LoadOptions()
    On Local Error Resume Next
    'Add a \ to the path if needed
    Dim strAppPath As String
    strAppPath = FixPath(App.Path)
    'Get the config file's path
    Dim strConfigFile As String
    strConfigFile = strAppPath & "Config.ini"
    
    '[General]
    strPattern = GetINISetting(strConfigFile, "General", _
        "Pattern", "*.rtf;*.txt") 'Code pattern
    strCodePath = GetINISetting(strConfigFile, "General", _
        "Codes Path", strAppPath & "Codes\") 'Location of codes
    If DoesDirectoryExist(strCodePath) = False Then _
        strCodePath = strAppPath & "Codes\" 'If we don't find the path use the default one
    strFormattingPath = GetINISetting(strConfigFile, "General", _
        "Formatting Templates Path", strAppPath & "Formatting\")
    If DoesDirectoryExist(strFormattingPath) = False Then _
        strFormattingPath = strAppPath & "Formatting\" 'If we don't find the path use the default one
    bolAutoShowCode = GetINISetting(strConfigFile, "General", _
        "Auto Show Code", True) 'Show code on load
    bolAutoShowNotes = GetINISetting(strConfigFile, "General", _
        "Auto Show Notes", False) 'Show notes on load
    strBookmarksFile = GetINISetting(strConfigFile, "General", _
        "Bookmarks File", strAppPath & "Bookmarks.ini") 'The old bookmarks file
    bolSaveBookmarks = GetINISetting(strConfigFile, "General", _
        "Save Bookmarks", True) 'The old bookmarks file
    bolCheckCodesOnExpand = GetINISetting(strConfigFile, "General", _
        "Check Codes On Node Expand", False) 'Whether to check for changed codes on node click
    If DoesFileExist(strAppPath & "Notes Seperator.txt") Then 'What seperates the notes as they are added
        strNotesSeperator = OpenText(strAppPath & "Notes Seperator.txt")
    Else
        'Default
        strNotesSeperator = "------------------" & NewLine & _
            "Taken from <?title?> in <?section?> on <?date?> at <?time?>" & NewLine(2)
    End If
    bolFixSeperator = GetINISetting(strConfigFile, "General", _
        "Fix Seperator", True) 'Show notes on load
    bolConfirmExit = GetINISetting(strConfigFile, "General", _
        "Confirm Exit", True) 'Confirm exit?
    bolShowRoot = GetINISetting(strConfigFile, "General", _
        "Show Root", True) 'Show root node/dir?
    intDefaultIndent = GetINISetting(strConfigFile, "General", _
        "Default Indent", 10) 'Indentation size
    intTreeTimerInterval = GetINISetting(strConfigFile, "General", _
        "Tree Speed", 25) 'Tree show/hide speed
    picTreeSizer.Left = GetINISetting(strConfigFile, "General", _
        "Tree Width", picTreeSizer.Left) 'Tree width
    If GetINISetting(strConfigFile, "General", _
        "Tree Visible", True) = False Then 'Tree (not) Visible
        picTreeSizer.Left = -picTreeSizer.Width
        picTreeSizer.MousePointer = vbDefault
        tvwCodes.Left = picTreeSizer.Left - tvwCodes.Width
        mnuTreeHideShow.Caption = "S&how"
    End If
    strStandardToolbarPath = GetINISetting(strConfigFile, "General", _
        "Standard Toolbar Path", strAppPath & "Toolbars\Standard\") 'Standard Toolbar Path
    If DoesDirectoryExist(strStandardToolbarPath) = False Then _
        strStandardToolbarPath = strAppPath & "Toolbars\Standard" 'If we don't find the path use the default one
    strFormattingToolbarPath = GetINISetting(strConfigFile, "General", _
        "Formatting Toolbar Path", strAppPath & "Toolbars\Formatting\") 'Formatting Toolbar Path
    If DoesDirectoryExist(strFormattingToolbarPath) = False Then _
        strFormattingToolbarPath = strAppPath & "Toolbars\Formatting\" 'If we don't find the path use the default one
    If GetINISetting(strConfigFile, "General", _
        "Standard Toolbar Visible", True) Then    'Standard Toolbar Visible
        tbrStandard.Visible = True
        mnuViewToolbarsStandard.Checked = True
    Else
        tbrStandard.Visible = False
        mnuViewToolbarsStandard.Checked = False
    End If
    If GetINISetting(strConfigFile, "General", _
        "Formatting Toolbar Visible", True) Then 'Formatting Toolbar Visible
        tbrFormatting.Visible = True
        mnuViewToolbarsFormatting.Checked = True
    Else
        tbrFormatting.Visible = False
        mnuViewToolbarsFormatting.Checked = False
    End If
    If GetINISetting(strConfigFile, "General", _
        "Statusbar Visible", True) Then 'Status bar Visible
        sbrInfo.Visible = True
        mnuViewStatusBar.Checked = True
    Else
        sbrInfo.Visible = False
        mnuViewStatusBar.Checked = False
    End If
    With tbrFormatting.Buttons("Change Case") 'Default Case
        .Tag = GetINISetting(strConfigFile, "General", _
            "Change Case Default", "lower case")
        .ToolTipText = "Change Case - " & .Tag
    End With
    imlCodesLarge.MaskColor = GetColourFromString(GetINISetting(strConfigFile, "General", _
        "Mask Colour", RGB(255, 0, 255)), RGB(255, 0, 255)) 'Mask Colour
    imlCodesSmall.MaskColor = imlCodesLarge.MaskColor
    imlToolbarFormatting.MaskColor = imlCodesLarge.MaskColor
    imlToolbarStandard.MaskColor = imlCodesLarge.MaskColor
    clsCodeUndo.MaxUndos = GetINISetting(strConfigFile, "General", _
        "Max Undos", 99) 'Max undos
    clsNotesUndo.MaxUndos = clsCodeUndo.MaxUndos
    bolWordWrapCode = GetINISetting(strConfigFile, "General", _
        "Word Wrap Code", True) 'Word wrap code
    Call SetWrap(rtfCode.hWnd, bolWordWrapCode)
    bolWordWrapNotes = GetINISetting(strConfigFile, "General", _
        "Word Wrap Notes", True) 'Word wrap notes
    Call SetWrap(rtfNotes.hWnd, bolWordWrapNotes)
    
    '[Window]:
    Me.WindowState = GetINISetting(strConfigFile, "Window", _
        "Window State", vbNormal)
    If Me.WindowState = vbNormal Then
        Dim sngTemp As Single
        'Width
        sngTemp = GetINISetting(strConfigFile, "Window", _
            "Width", Me.Width)
        If sngTemp > -1 Then Me.Width = sngTemp
        'Left
        sngTemp = GetINISetting(strConfigFile, "Window", _
            "Left", (Screen.Width \ 2) - (Me.Width \ 2))
        If sngTemp > -1 And sngTemp < Screen.Width - Me.Width Then _
            Me.Left = sngTemp
        
        'Height
        sngTemp = GetINISetting(strConfigFile, "Window", _
            "Height", Me.Height)
        If sngTemp > -1 Then Me.Height = sngTemp
        'Top
        sngTemp = GetINISetting(strConfigFile, "Window", _
            "Top", (Screen.Height \ 2) - (Me.Height \ 2))
        If sngTemp > -1 And sngTemp < Screen.Height - Me.Height Then _
            Me.Top = sngTemp
    End If
End Sub

Private Sub SaveOptions()
    On Local Error Resume Next
    'Add a \ to the path if needed
    Dim strAppPath As String
    strAppPath = FixPath(App.Path)
    'Get the config file's path
    Dim strConfigFile As String
    strConfigFile = strAppPath & "Config.ini"
    
    '[General]:
    Call SaveINISetting(strConfigFile, "General", _
        "Pattern", strPattern) 'Code pattern
    Call SaveINISetting(strConfigFile, "General", _
        "Codes Path", strCodePath) 'Location of codes
    Call SaveINISetting(strConfigFile, "General", _
        "Formatting Templates Path", strFormattingPath) 'Path for formatting config files
    Call SaveINISetting(strConfigFile, "General", _
        "Auto Show Code", bolAutoShowCode) 'Show code on load
    Call SaveINISetting(strConfigFile, "General", _
        "Auto Show Notes", bolAutoShowNotes) 'Show notes on load
    Call SaveINISetting(strConfigFile, "General", _
        "Bookmarks File", strBookmarksFile) 'The old bookmarks file
    Call SaveINISetting(strConfigFile, "General", _
        "Save Bookmarks", bolSaveBookmarks) 'The old bookmarks file
    Call SaveText(strAppPath & "\Notes Seperator.txt", strNotesSeperator) 'What seperates the notes as they are added
    Call SaveINISetting(strConfigFile, "General", _
        "Fix Seperator", bolFixSeperator) 'Show notes on load
    Call SaveINISetting(strConfigFile, "General", _
        "Confirm Exit", bolConfirmExit) 'Confirm exit?
    Call SaveINISetting(strConfigFile, "General", _
        "Show Root", bolShowRoot) 'Show root node/dir?
     Call SaveINISetting(strConfigFile, "General", _
        "Check Codes On Node Expand", bolCheckCodesOnExpand)  'Check for changed codes on node click
    Call SaveINISetting(strConfigFile, "General", _
        "Default Indent", intDefaultIndent) 'Indentation size
    Call SaveINISetting(strConfigFile, "General", _
        "Tree Speed", intTreeTimerInterval) 'Tree show/hide speed
    Call SaveINISetting(strConfigFile, "General", _
        "Tree Width", picTreeSizer.Left)  'Tree width
    Call SaveINISetting(strConfigFile, "General", _
        "Tree Visible", mnuTreeHideShow.Caption = "&Hide") 'Tree Visible
    Call SaveINISetting(strConfigFile, "General", _
        "Standard Toolbar Path", strStandardToolbarPath)  'Standard Toolbar Path
    Call SaveINISetting(strConfigFile, "General", _
        "Standard Toolbar Visible", mnuViewToolbarsStandard.Checked)  'Standard Toolbar Visible
    Call SaveINISetting(strConfigFile, "General", _
        "Formatting Toolbar Path", strFormattingToolbarPath)  'Formatting Toolbar Path
    Call SaveINISetting(strConfigFile, "General", _
        "Formatting Toolbar Visible", mnuViewToolbarsFormatting.Checked)  'Formatting Toolbar Visible
    Call SaveINISetting(strConfigFile, "General", _
        "Statusbar Visible", sbrInfo.Visible)  'Status bar Visible
    Call SaveINISetting(strConfigFile, "General", _
            "Change Case Default", tbrFormatting.Buttons("Change Case").Tag) 'Default Case
    Call SaveINISetting(strConfigFile, "General", _
        "Mask Colour", imlCodesLarge.MaskColor)  'Mask Colour
    Call SaveINISetting(strConfigFile, "General", _
        "Max Undos", clsCodeUndo.MaxUndos)  'Max undos
    Call SaveINISetting(strConfigFile, "General", _
        "Word Wrap Code", bolWordWrapCode)  'Word wrap code
    Call SaveINISetting(strConfigFile, "General", _
        "Word Wrap Notes", bolWordWrapNotes)  'Word wrap notes
    
    '[Window]:
    If Me.WindowState <> vbMinimized Then Call SaveINISetting(strConfigFile, _
        "Window", "Window State", Me.WindowState) 'Window State
    'Only do if normal
    If Me.WindowState = vbNormal Then
        Call SaveINISetting(strConfigFile, "Window", "Width", Me.Width)  'Width
        Call SaveINISetting(strConfigFile, "Window", "Left", Me.Left)  'Left
        Call SaveINISetting(strConfigFile, "Window", "Height", Me.Height)  'Height
        Call SaveINISetting(strConfigFile, "Window", "Top", Me.Top)  'Top
    End If
End Sub

Private Sub NewRTFText(ByRef Filename As String)
    On Local Error GoTo ErrorHandler
    'File has not yet been saved
    Dim bolFileSaved As Boolean
    bolFileSaved = False
    
    Dim rtfSelected As RichTextBox
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = rtfCode
        Case "notes"
            Set rtfSelected = rtfNotes
        Case Else
            Exit Sub
    End Select
    'Ask to save changes
    Dim msgAnswer As VbMsgBoxResult
    msgAnswer = MsgBox("Save the changes before clearing?", vbYesNoCancel Or vbQuestion, "Sace changes?")
    'If yes
    If msgAnswer = vbYes Then
        'Get a filename if needed
        Dim strFileName As String
        If strNotesFilename = vbNullString Then
            strFileName = GetFileName(, True)
        'If not make = to the existing one
        Else
            strFileName = Filename
        End If
        'If we have a filename
        If strFileName <> vbNullString Then
            'Save the file and remove the filename
            Call rtfSelected.SaveFile(strFileName)
            bolFileSaved = True
            Filename = vbNullString
        
        'If we don't (i.e. cancel at Save dialog)
        Else
            'Exit without changing or saving
            Exit Sub
        End If
    
    'If Cancel
    ElseIf msgAnswer = vbCancel Then
        'Exit without saving or changing
        Exit Sub
    End If
    
    'If no
    If msgAnswer = vbNo Or strFileName <> vbNullString Then
        'Clear the text and the filename
        rtfSelected.Text = vbNullString
        Filename = vbNullString
        'Set the save button's enabled value
        Call SetSaveButton
        'Reset the undos
        Call clsCodeUndo.Reset
    End If
    Exit Sub
ErrorHandler:
    'Tell user file couldn't be saved
    If bolFileSaved = False Then MsgBox "The file could not be saved.", _
        vbCritical Or vbOKOnly, "Error"
End Sub

Private Sub LoadRTFText(ByRef Filename As String)
    On Local Error GoTo ErrorHandler
    'Whether the code has been loaded; initial = false
    Dim bolFileLoaded As Boolean
    bolFileLoaded = False
    
    'Get the active RTF box
    Dim rtfSelected As RichTextBox
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = rtfCode
        Case "notes"
            Set rtfSelected = rtfNotes
        Case Else
            Exit Sub
    End Select
    
    'Get a filename from the user
    Dim strFileName As String
    strFileName = GetFileName
    'If name is the same as current, check user wants to revert
    Dim msgAnswer As VbMsgBoxResult
    msgAnswer = vbOK
    If strFileName <> "" And strFileName = rtfSelected.Filename Then _
        msgAnswer = MsgBox("Reload the current file and loose all changes since last save?", _
                vbQuestion Or vbOKCancel, "Revert?") = vbOK
    'If the user choose one or they want to revert/is different filename
    If strFileName <> "" And msgAnswer = vbOK Then
        'Load it
        rtfSelected.Filename = strFileName
        'File has now been loaded
        bolFileLoaded = True
        'Set it's filename to the var
        Filename = strFileName
        'Reset the undos
        Call clsCodeUndo.Reset
        'If we are to show the notes, show them
        If rtfSelected.Name = "rtfCode" And bolAutoShowCode Then
            tbsView.Tabs("Code").Selected = True
            Call tbsView_MouseUp(vbLeftButton, 0, 0, 0)
        ElseIf rtfSelected.Name = "rtfNotes" And bolAutoShowNotes Then
            tbsView.Tabs("Notes").Selected = True
            Call tbsView_MouseUp(vbLeftButton, 0, 0, 0)
        End If
        'Set the char/line number
        Call SelectionChange
        'Set the save button's enabled value
        Call SetSaveButton
    End If
    Exit Sub
ErrorHandler:
    'Tell the user of error if file hasn't been loaded
    If bolFileLoaded = False Then MsgBox "An error has occured and the file could not be loaded.", _
        vbCritical Or vbOKOnly, "Error"
End Sub

Private Sub SaveRTFText(ByRef Filename As String)
    On Local Error Resume Next
    Dim rtfSelected As RichTextBox
    'Get the active RTF box
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = rtfCode
        Case "notes"
            Set rtfSelected = rtfNotes
        Case Else
            Exit Sub
    End Select
    
    'If we have no filename
    If strCodeFilename = vbNullString Then
        'Do SaveAs instead
        Call SaveAsRTFText(Filename)
        Exit Sub
    'If we have a filename
    Else
        'Save over last one
        Call rtfSelected.SaveFile(Filename)
    End If
End Sub

Private Sub SaveAsRTFText(ByRef Filename As String)
    On Local Error Resume Next
    Dim rtfSelected As RichTextBox
    'Get the active RTF box
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = rtfCode
        Case "notes"
            Set rtfSelected = rtfNotes
        Case Else
            Exit Sub
    End Select
    
    Dim strFileName As String
    'Get a new filename
    strFileName = GetFileName(, True)
    'If cancel was not selected
    If strFileName <> vbNullString Then
        'Load the new file and set its filename
        Call rtfSelected.SaveFile(strFileName)
        Filename = strFileName
    End If
    'Set the save button's enabled value
    Call SetSaveButton
End Sub

Private Sub PrintRTFText()
    On Local Error GoTo ErrorHandler
    'Get the active RTF box
    Dim rtfSelected As RichTextBox
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = rtfCode
        Case "notes"
            Set rtfSelected = rtfNotes
        Case Else
            Exit Sub
    End Select
    
    With cdlDialogs
        'Show the printer dialog
        .CancelError = True
        .ShowPrinter
        'Get the selected start & length
        Dim lngSelStart As Long, lngSelLength As Long
        lngSelStart = rtfSelected.SelStart
        lngSelLength = rtfSelected.SelLength
        'If only part of the text is selected
        If lngSelLength > 0 And lngSelLength < Len(rtfSelected.Text) Then
            'Ask the user if they only want to print the selected text
            Dim msgAnswer As VbMsgBoxResult
            msgAnswer = MsgBox("Would you like to print only the selected text?", vbYesNoCancel, "Print Selected Only")
            'If no
            If msgAnswer = vbNo Then
                'Select all text
                rtfSelected.SelStart = 0
                rtfSelected.SelLength = Len(rtfSelected.Text)
            'Cancel
            ElseIf msgAnswer = vbCancel Then
                'Exit without printing
                Exit Sub
            End If
        End If
        'Print the text
        Call rtfSelected.SelPrint(Printer.hDC)
    End With
    'If the selected text is different from what it was restore it
    If rtfSelected.SelStart <> lngSelStart Then rtfSelected.SelStart = lngSelStart
    If rtfSelected.SelLength <> lngSelLength Then rtfSelected.SelLength = lngSelLength
    
    'Exit so as not to cause an error
    Exit Sub
ErrorHandler:
    If Err.Number <> cdlCancel Then MsgBox "An error has occured.", vbCritical Or vbOKOnly, "Error"
End Sub

Private Sub EditRTFText(ByVal Message As TextMessages, _
    Optional ByVal Parameter1 As Long, Optional ByVal Parameter2 As Long)
    On Local Error Resume Next
    'Convert the tab key to lowercase for ease of comparision
    Dim strTabKey As String
    strTabKey = LCase(tbsView.SelectedItem.Key)
    'Get the correct hWnd
    Dim lngRTFhWnd As Long
    lngRTFhWnd = Switch(strTabKey = "code", rtfCode.hWnd, _
        strTabKey = "notes", rtfNotes.hWnd)
    'Send the message (e.g. Cut, Copy, Paste, etc) to the RTF Box
    Call SendMessage(lngRTFhWnd, Message, Parameter1, Parameter2)
End Sub

Private Sub SetSaveButton()
    On Local Error Resume Next
    'Choose the active tab, and set the save button according to what the filename if for that rtf box
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            tbrStandard.Buttons("Save").Enabled = strCodeFilename <> vbNullString
        Case "notes"
            tbrStandard.Buttons("Save").Enabled = strNotesFilename <> vbNullString
        Case Else
            tbrStandard.Buttons("Save").Enabled = False
    End Select
End Sub

Private Sub LoadFormattingOptions()
    On Local Error GoTo ErrorHandler
    'Get the first file in the formatting directory
    Dim strFileName() As String
    strFileName() = GetConfigFiles(strFormattingPath)
    strFileName = BubbleSort(strFileName)
    'Loop for all files
    Dim intLoopCounter As Integer
    For intLoopCounter = 0 To UBound(strFileName)
        'If we already have some options loaded load a new menu item
        If mnuEditFormatFormatAsOption(mnuEditFormatFormatAsOption.UBound).Enabled Then _
            Load mnuEditFormatFormatAsOption(mnuEditFormatFormatAsOption.UBound + 1)
        'Set the new menu's caption and show it
        With mnuEditFormatFormatAsOption(mnuEditFormatFormatAsOption.UBound)
            .Caption = strFileName(intLoopCounter)
            .Visible = True
            .Enabled = True
        End With
    Next intLoopCounter
    Exit Sub
ErrorHandler:
End Sub

Private Function GetSeperator() As String
    On Local Error Resume Next
    GetSeperator = strNotesSeperator
    'If we are to replace <?TAGS?> in the seperator
    If bolFixSeperator Then
        '<?section?> = The current section
        GetSeperator = Replace(GetSeperator, "<?section?>", GetSection, , , vbTextCompare)
        '<?title?> = The current file
        GetSeperator = Replace(GetSeperator, "<?title?>", tvwCodes.SelectedItem.Text, , , vbTextCompare)
        '<?date?> = The current date
        GetSeperator = Replace(GetSeperator, "<?date?>", Date, , , vbTextCompare)
        '<?time?> = The current time
        GetSeperator = Replace(GetSeperator, "<?time?>", Time, , , vbTextCompare)
    End If
End Function

Public Sub NewRTFColour(ByVal Colour As Long)
    Select Case LCase(tbsView.SelectedItem.Key)
        Case "code"
            rtfCode.SelColor = Colour
        Case "notes"
            rtfNotes.SelColor = Colour
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub ValidateCodes()
    On Local Error Resume Next
    'Add the nodes from the code path to add any new folders files
    '(old ones will not be overwritten)
    'Call AddNodes(tvwCodes, strCodePath, , , False)
    Call FindFiles(strCodePath, Me, "AddCodes")
    
    'Loop for all codes/sections in Code Tree
    Dim lngLoopCounter As Integer
    lngLoopCounter = 1
    Do While lngLoopCounter <= tvwCodes.Nodes.Count
        'If we are in a section
        If Right(tvwCodes.Nodes(lngLoopCounter).Key, Len("\")) = "\" Then
            'If the folder/section no longer exists
            If DoesDirectoryExist(tvwCodes.Nodes(lngLoopCounter).Key) = False Then
                'Remove it's node
                Call tvwCodes.Nodes.Remove(lngLoopCounter)
                'Set LoopCounter to 0 so the loop starts from scratch again, otherwise we may miss some nodes
                lngLoopCounter = 0
            'If if does exist
            Else
                'Call the Expand sub so that any new folders/files are found and added
                Call tvwCodes_Expand(tvwCodes.Nodes(lngLoopCounter))
            End If
        
        'If we have a code
        Else
            'If the file no longer exists
            If DoesFileExist(tvwCodes.Nodes(lngLoopCounter).Key) = False Then
                'Remove it's node
                Call tvwCodes.Nodes.Remove(lngLoopCounter)
                'Set LoopCounter to 0 so the loop starts from scratch again, otherwise we may miss some nodes
                lngLoopCounter = 0
            End If
        End If
        'Incremenet the counter
        lngLoopCounter = lngLoopCounter + 1
    Loop
End Sub

Private Sub ValidateBookmarks()
    On Local Error Resume Next
    'Loop for all Bookmarks
    Dim lngLoopCounter As Long
    lngLoopCounter = 1
    Do While lngLoopCounter <= lvwBookmarks.ListItems.Count
        'If we are in a section
        If Right(lvwBookmarks.ListItems(lngLoopCounter).Key, Len("\")) = "\" Then
            'If the folder/section no longer exists
            If DoesDirectoryExist(lvwBookmarks.ListItems(lngLoopCounter).Key) = False Then
                    'And DoesListItemExist(lvwBookmarks, lvwBookmarks.ListItems(LoopCounter).Key)
                'Remove it's list item
                Call lvwBookmarks.ListItems.Remove(lvwBookmarks.ListItems(lngLoopCounter).Key)
                'Set LoopCounter to 0 so the loop starts from scratch again, otherwise we may miss some list items
                lngLoopCounter = 0
            End If
        
        'If we have a code
        Else
            'If the file no longer exists
            If DoesFileExist(lvwBookmarks.ListItems(lngLoopCounter).Key) = False Then
                'If DoesListItemExist(lvwBookmarks, lvwBookmarks.ListItems(LoopCounter).Key) Then
                'Remove it's list item
                Call lvwBookmarks.ListItems.Remove(lvwBookmarks.ListItems(lngLoopCounter).Key)
                'Set LoopCounter to 0 so the loop starts from scratch again, otherwise we may miss some list items
                lngLoopCounter = 0
                'End If
            End If
        End If
        'Incremenet the counter
        lngLoopCounter = lngLoopCounter + 1
    Loop
End Sub

Private Sub LoadBookmarks(ByVal Filename As String)
    On Local Error GoTo ErrorHandler
    Dim strFileText As String 'The bookmark file's text
    'If the file exists
    If DoesFileExist(Filename) Then
        'Get the bookmarks
        strFileText = OpenText(Filename)
        
        'Split the bookmarks by new lines
        Dim strBookmarks() As String
        strBookmarks() = Split(strFileText, vbNewLine, , vbTextCompare)
        
        'Loop for all bookmarks
        Dim lngLoopCounter As Long
        For lngLoopCounter = 0 To UBound(strBookmarks)
            'Add this one to the bookmarks
            If Right(strBookmarks(lngLoopCounter), Len("\")) = "\" Then
                If DoesDirectoryExist(strBookmarks(lngLoopCounter)) Then _
                    Call AddBookmark(strBookmarks(lngLoopCounter), False)
            Else
                If DoesFileExist(strBookmarks(lngLoopCounter)) Then _
                    Call AddBookmark(strBookmarks(lngLoopCounter), False)
            End If
        Next lngLoopCounter
    End If
    Exit Sub
ErrorHandler:
    MsgBox "An error has occured and the bookmarks could not be loaded.", vbCritical Or vbOKOnly, "Error"
End Sub

Private Sub SaveBookmarks(ByVal Filename As String)
    On Local Error GoTo ErrorHandler
    Dim strOutput As String 'What gets written to the file
    
    'Loop for all bookmarks
    Dim lngLoopCounter As Long
    For lngLoopCounter = 1 To lvwBookmarks.ListItems.Count
        'Add the current bookmarks key (and a new line if not on the last one)
        strOutput = strOutput & lvwBookmarks.ListItems(lngLoopCounter).Key & _
            IIf(lngLoopCounter < lvwBookmarks.ListItems.Count, vbNewLine, vbNullString)
    Next lngLoopCounter
    'Save the text to the specified file
    Call SaveText(Filename, strOutput)
    Exit Sub

ErrorHandler:
    'Tell the user
    MsgBox "An error has occured and the bookmarks could not be loaded.", vbCritical Or vbOKOnly, "Error"
End Sub

Private Sub AddBookmark(Optional ByVal Key, _
    Optional ByVal PromptOnError As Boolean = True)
    On Local Error GoTo ErrorHandler
    If Key = Null Then Key = tvwCodes.Nodes(tvwCodes.SelectedItem).Key
    
    With tvwCodes
        'If there is a selection
        'If HasSelectedItem(lvwBookmarks) Then
        'If the bookmark is already present
        If DoesListItemExist(lvwBookmarks, Key) Then
            'Tell the user (if wanted)
            If PromptOnError Then MsgBox _
                "The current code/section already exists in your bookmarks.", _
                vbExclamation Or vbOKOnly, "Error"
            Exit Sub
        
        'If the bookmark's not already there
        Else
            Dim strIconKey As String
            'If the bookmark is a section rather than a code
            If Right(Key, Len("\")) = "\" Then
                'Section = Key
                strIconKey = Key
            'Code
            Else
                'Section = Parent's Key
                strIconKey = tvwCodes.Nodes(Key).Parent.Key
            End If
            'Add it to the bookmarks
            lvwBookmarks.ListItems.Add , Key, .Nodes(Key).Text, strIconKey, strIconKey
            'Add the section for the new bookmark
            Call lvwBookmarks.ListItems(Key).ListSubItems.Add(, , GetSection(Key))
            'Un/Bold the menu
            If rtfCode.Filename = Key Then Call BoldBookmark(Key)
        End If
        'There is no selection
        'Else
            'Tell the user
        '    MsgBox "There is no code/section currently selected. Please select one and then try again.", _
                vbExclamation Or vbOKOnly, "Error"
        'End If
    End With
    Exit Sub
    
ErrorHandler:
    If PromptOnError Then MsgBox "An error has occured and the bookmark could not be added.", _
        vbCritical Or vbOKOnly, "Error"
End Sub

Private Function GetCount() As typCount
    On Local Error Resume Next
    'Loop for all nodes
    Dim lngLoopCounter As Long
    Do While lngLoopCounter < tvwCodes.Nodes.Count
        'Increment the loop counter
        lngLoopCounter = lngLoopCounter + 1
        'If not a folder/section (i.e. a code) make it visivle
        If Right(tvwCodes.Nodes(lngLoopCounter).Key, Len("\")) <> "\" Then
            GetCount.Codes = GetCount.Codes + 1
        Else
            GetCount.Sections = GetCount.Sections + 1
        End If
    Loop
End Function

'Loads the weblinks to the Help --> Web Links Menus
Public Sub LoadLinks()
    'Get the site list
    Dim strTemp As String
    strTemp = OpenText(FixPath(App.Path) & "Site List.txt")
    
    'If found
    If strTemp <> "" Then
        Dim strVBSites() As String
        strVBSites() = Split(strTemp, vbNewLine, , vbTextCompare)
        Dim intLoopCounter As Integer
        For intLoopCounter = 0 To UBound(strVBSites)
            If strVBSites(intLoopCounter) <> vbNullString Then
                If intLoopCounter > 0 Then Load mnuHelpLinksVBSites(mnuHelpLinksVBSites.UBound + 1)
                With mnuHelpLinksVBSites(mnuHelpLinksVBSites.UBound)
                    .Caption = strVBSites(intLoopCounter)
                    .Enabled = True
                    .Visible = True
                End With
            End If
        Next intLoopCounter
    End If
End Sub

Private Sub RTFKeyDown(ByVal KeyCode As KeyCodeConstants, ByVal Shift As ShiftConstants)
    'If (KeyCode And vbKeyB) And shift = vbCtrlMask Then
    '    tbrFormatting.Buttons("Bold").Value = IIf(tbrFormatting.Buttons("Bold").Value = tbrPressed, tbrUnpressed, tbrPressed)
    '    Call tbrFormatting_ButtonClick(tbrFormatting.Buttons("Bold"))
    'ElseIf (KeyCode And vbKeyI) And shift = vbCtrlMask Then
    '    tbrFormatting.Buttons("Italic").Value = IIf(tbrFormatting.Buttons("Italic").Value = tbrPressed, tbrUnpressed, tbrPressed)
    '    Call tbrFormatting_ButtonClick(tbrFormatting.Buttons("Italic"))
    'ElseIf (KeyCode And vbKeyU) And shift = vbCtrlMask Then
    '    tbrFormatting.Buttons("Underline").Value = IIf(tbrFormatting.Buttons("Underline").Value = tbrPressed, tbrUnpressed, tbrPressed)
    '    Call tbrFormatting_ButtonClick(tbrFormatting.Buttons("Underline"))
    'End If
End Sub

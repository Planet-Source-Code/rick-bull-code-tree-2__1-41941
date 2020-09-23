VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmFormat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFormat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   2265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtBulletIndent 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Text            =   "0"
      Top             =   1560
      Width           =   510
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2400
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtHangingIndent 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "0"
      Top             =   1080
      Width           =   510
   End
   Begin VB.TextBox txtRightIndent 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Width           =   510
   End
   Begin ComCtl2.UpDown udIndent 
      Height          =   300
      Left            =   1950
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   327681
      BuddyControl    =   "txtIndent"
      BuddyDispid     =   196614
      OrigLeft        =   1560
      OrigTop         =   120
      OrigRight       =   1800
      OrigBottom      =   375
      Increment       =   10
      Max             =   1000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtIndent 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "0"
      Top             =   120
      Width           =   510
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   300
      Left            =   1950
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   327681
      BuddyControl    =   "txtRightIndent"
      BuddyDispid     =   196613
      OrigLeft        =   1560
      OrigTop         =   120
      OrigRight       =   1800
      OrigBottom      =   375
      Increment       =   10
      Max             =   1000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown udHangingIndent 
      Height          =   300
      Left            =   1950
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   327681
      BuddyControl    =   "txtHangingIndent"
      BuddyDispid     =   196612
      OrigLeft        =   1560
      OrigTop         =   120
      OrigRight       =   1800
      OrigBottom      =   375
      Increment       =   10
      Max             =   1000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown udBulletIndent 
      Height          =   300
      Left            =   1950
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1560
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   327681
      BuddyControl    =   "txtBulletIndent"
      BuddyDispid     =   196609
      OrigLeft        =   1560
      OrigTop         =   120
      OrigRight       =   1800
      OrigBottom      =   375
      Increment       =   10
      Max             =   1000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Bullet Indent:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Handing Indent:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1185
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Right Indent:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   660
      Width           =   960
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Indent:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   540
   End
End
Attribute VB_Name = "frmFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
    On Local Error Resume Next
    Call InitCommonControls
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rtfSelected As RichTextBox
    Select Case LCase(frmMain.tbsView.SelectedItem.Key)
        Case "code"
            Set rtfSelected = frmMain.rtfCode
            
        Case "notes"
            Set rtfSelected = frmMain.rtfNotes
        
        Case Else
            GoTo Finish
    End Select
    With rtfSelected
        .SelIndent = TwipsX(txtIndent.Text)
        .SelRightIndent = TwipsX(txtRightIndent.Text)
        .SelHangingIndent = TwipsX(txtHangingIndent.Text)
        .BulletIndent = TwipsX(txtBulletIndent.Text)
    End With
Finish:
    Unload Me
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    'Make buttons 3D
    Call FormatButtons(Me)
End Sub

Private Sub txtHangingIndent_LostFocus()
    txtHangingIndent.Text = Val(txtHangingIndent.Text)
End Sub

Private Sub txtIndent_LostFocus()
    txtIndent.Text = Val(txtIndent.Text)
End Sub

Private Sub txtRightIndent_LostFocus()
    txtRightIndent.Text = Val(txtRightIndent.Text)
End Sub

Private Sub txtBulletIndent_LostFocus()
    txtBulletIndent.Text = Val(txtBulletIndent.Text)
End Sub

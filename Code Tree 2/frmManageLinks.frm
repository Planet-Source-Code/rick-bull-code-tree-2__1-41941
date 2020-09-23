VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmManageLinks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manage Links"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3420
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3420
      Width           =   1335
   End
   Begin CodeTree.Seperator Seperator1 
      Height          =   30
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit.."
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdSeperator 
      Caption         =   "&Seperator"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Move &Down"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Move &Up"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add.."
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "D&elete"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwLinks 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   5530
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Address"
         Text            =   "Address"
         Object.Width           =   6174
      EndProperty
   End
End
Attribute VB_Name = "frmManageLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    On Local Error Resume Next
    Dim strAddress As String
    strAddress = InputBox("Enter the address of the web site you would like to add:" & vbNewLine, _
        "Add Link", "http://www.rickbull.com/")
    If strAddress <> vbNullString Then lvwLinks.ListItems.Add _
        lvwLinks.SelectedItem.Index, , strAddress
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    On Local Error Resume Next
    Dim intLoopCounter As Integer
    intLoopCounter = 1
    Do While intLoopCounter <= lvwLinks.ListItems.Count
        If lvwLinks.ListItems(intLoopCounter).Selected Then
            Call lvwLinks.ListItems.Remove(intLoopCounter)
            intLoopCounter = 0
        End If
        intLoopCounter = intLoopCounter + 1
    Loop
    'Call lvwLinks.ListItems.Remove(lvwLinks.SelectedItem.Key)
    Call CheckButtons
    lvwLinks.SetFocus
End Sub

Private Sub cmdDown_Click()
    On Local Error Resume Next
    If lvwLinks.SelectedItem.Index < lvwLinks.ListItems.Count Then
        Dim strTemp As String
        strTemp = lvwLinks.ListItems(lvwLinks.SelectedItem.Index + 1).Text
        lvwLinks.ListItems(lvwLinks.SelectedItem.Index + 1).Text = lvwLinks.ListItems(lvwLinks.SelectedItem.Index).Text
        lvwLinks.ListItems(lvwLinks.SelectedItem.Index).Text = strTemp
        lvwLinks.ListItems(lvwLinks.SelectedItem.Index).Selected = False
        lvwLinks.ListItems(lvwLinks.SelectedItem.Index + 1).Selected = True
        lvwLinks.SetFocus
    Else
        MsgBox "You cannot move this item dpwn - already at the bottom!", vbOKOnly Or vbExclamation, "Error"
    End If
End Sub

Private Sub cmdOK_Click()
    On Local Error Resume Next
    Dim strTemp As String
    Dim intLoopCounter As Integer
    'Add all the links together
    For intLoopCounter = 1 To lvwLinks.ListItems.Count
        strTemp = strTemp & lvwLinks.ListItems(intLoopCounter) & _
            IIf(intLoopCounter < lvwLinks.ListItems.Count, vbNewLine, vbNullString)
    Next intLoopCounter
    'Save them to the file
    Call SaveText(FixPath(App.Path) & "Site List.txt", strTemp)
    If frmMain.mnuHelpLinksVBSites.LBound < frmMain.mnuHelpLinksVBSites.UBound Then
        'Unload them all (minus the first)
        For intLoopCounter = frmMain.mnuHelpLinksVBSites.LBound + 1 To frmMain.mnuHelpLinksVBSites.UBound
            Unload frmMain.mnuHelpLinksVBSites(intLoopCounter)
        Next intLoopCounter
        'Make the first one disabled
        With frmMain.mnuHelpLinksVBSites(frmMain.mnuHelpLinksVBSites.LBound)
            .Caption = "[NONE]"
            .Enabled = False
        End With
    End If
    'Show the new ones
    Call frmMain.LoadLinks
    Unload Me
End Sub

Private Sub cmdUp_Click()
    On Local Error Resume Next
    If lvwLinks.SelectedItem.Index > 1 Then
        Dim strTemp As String
        strTemp = lvwLinks.ListItems(lvwLinks.SelectedItem.Index - 1).Text
        lvwLinks.ListItems(lvwLinks.SelectedItem.Index - 1).Text = lvwLinks.ListItems(lvwLinks.SelectedItem.Index).Text
        lvwLinks.ListItems(lvwLinks.SelectedItem.Index).Text = strTemp
        lvwLinks.ListItems(lvwLinks.SelectedItem.Index).Selected = False
        lvwLinks.ListItems(lvwLinks.SelectedItem.Index - 1).Selected = True
        lvwLinks.SetFocus
    Else
        MsgBox "You cannot move this item up - already at the top!", vbOKOnly Or vbExclamation, "Error"
    End If
End Sub

Private Sub cmdSeperator_Click()
    On Local Error Resume Next
    lvwLinks.ListItems.Add lvwLinks.SelectedItem.Index, , "-"
    lvwLinks.Refresh
End Sub

Private Sub cmdEdit_Click()
    On Local Error Resume Next
    Dim strTemp As String
    strTemp = InputBox("Please enter the address for the item:", _
        "Enter Address", lvwLinks.SelectedItem.Text)
    If strTemp <> vbNullString Then lvwLinks.SelectedItem.Text = strTemp
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    Dim strTemp As String, strSites() As String
    strTemp = OpenText(FixPath(App.Path) & "Site List.txt")
    strSites() = Split(strTemp, vbNewLine)
    Dim intLoopCounter As Integer
    For intLoopCounter = LBound(strSites) To UBound(strSites)
        lvwLinks.ListItems.Add , , _
            strSites(intLoopCounter) 'IIf(strSites(intLoopCounter) = "-", "[SEPERATOR]", _
            strSites(intLoopCounter))
    Next intLoopCounter
    Call FormatButtons(Me)
    Call CheckButtons
End Sub

Private Sub CheckButtons()
    On Local Error Resume Next
    cmdDelete.Enabled = lvwLinks.ListItems.Count > 0
    'cmdUp.Enabled = cmdDelete.Enabled And lvwLinks.ListItems.Count > 0
    'cmdDown.Enabled = cmdUp.Enabled
End Sub


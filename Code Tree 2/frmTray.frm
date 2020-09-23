VERSION 5.00
Begin VB.Form frmTray 
   ClientHeight    =   510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   510
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPopupRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuPopupSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuPopupExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Original Tray code info:
'Created by the KPD-Team 1999
'Pieter Philippaerts
'URL: http://users.turboline.be/btl10148/
'E-Mail: kpd_team@hotmail.com

'Type for the Tray Icon
Private Type NOTIFYICONDATA
    cbSize As Long 'Size of the variable
    hWnd As Long 'Owner hWnd
    uId As Long
    uFlags As Long 'Flags
    ucallbackMessage As Long 'Call back from which Window Message
    hIcon As Long 'Icon's Handle
    szTip As String * 64 'Tooltip
End Type

'API Constants: Notify Icon Messages (NIM)
Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
'Notify Icon Flags (NIF)
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'Window Messages (WM)
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONUP = &H205

'API Declarations
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Variable declarations
Private nidTray As NOTIFYICONDATA 'Holds the info about the Tray Icon

Private Sub Form_Initialize()
    On Local Error Resume Next
    Call InitCommonControls
End Sub

Private Sub Form_Load()
    On Local Error Resume Next
    With nidTray
        'Set the size of the variable
        .cbSize = Len(nidTray)
        'Link the trayicon to this form
        .hWnd = Me.hWnd
        .uId = 1&
        'Show the Icon, Tooltip, and return Window Messages
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        'Call back when Left Mouse Butotn is pressed
        .ucallbackMessage = WM_LBUTTONDOWN
        'Tray icon = the main form's icon
        .hIcon = frmMain.Icon.Handle
        'Tooltip
        .szTip = "Code Tree" & vbNullChar
    End With
    'Create the icon
    Call Shell_NotifyIcon(NIM_ADD, nidTray)
    'Hide this form and the main form
    Me.Hide
    frmMain.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Local Error Resume Next
    With nidTray
        'Set the length of the variable
        .cbSize = Len(nidTray)
        'hWnd = this form
        .hWnd = Me.hWnd
        .uId = 1&
    End With
    'Delete the icon
    Call Shell_NotifyIcon(NIM_DELETE, nidTray)
    'Show the main form
    With frmMain
        .Visible = True
        Call .SetFocus
    End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Local Error Resume Next
    Dim lngMsg As Long
    lngMsg = X / TwipsX
    'If the user double-clicked on the icon
    If lngMsg = WM_LBUTTONDBLCLK Then
        'Do the default option
        Call mnuPopupRestore_Click
    'User Right clicked
    ElseIf lngMsg = WM_RBUTTONUP Then
        'Show popup menu
        Call PopupMenu(mnuPopup, , , , mnuPopupRestore)
    End If
End Sub

Private Sub mnuPopupAbout_Click()
    On Local Error Resume Next
    'Show the about form
    Load frmAbout
    frmAbout.Show vbModal
End Sub

Private Sub mnuPopupExit_Click()
    On Local Error Resume Next
    'Restore the form
    Call mnuPopupRestore_Click
    'Unload the main form
    Unload frmMain
End Sub

Private Sub mnuPopupRestore_Click()
    On Local Error Resume Next
    'Show the main form
    frmMain.Visible = True
    'Unload this one
    Unload Me
End Sub

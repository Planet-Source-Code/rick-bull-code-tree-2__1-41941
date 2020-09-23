VERSION 5.00
Begin VB.Form frmError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Details »"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1120
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   680
      Width           =   1215
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore"
      Height          =   375
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin CodeTree.Hyperlink lnkEmail 
      Height          =   225
      Left            =   2160
      TabIndex        =   2
      Top             =   1500
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Address         =   "mailto:rickbull@rickmusic.co.uk?subject=Code%20Tree"
      Caption         =   "E-Mail the Author About this Error"
      Alignment       =   2
      AutoSize        =   -1  'True
      CanGetFocus     =   -1  'True
      BeginProperty HotFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDetails 
      AutoSize        =   -1  'True
      Height          =   1000
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   6165
      WordWrap        =   -1  'True
   End
   Begin VB.Line lnBorder 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   240
      X2              =   6600
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line lnBorder 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   255
      X2              =   6600
      Y1              =   1935
      Y2              =   1920
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmError.frx":058A
      Height          =   1215
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   4425
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgIcon 
      Height          =   720
      Left            =   120
      Picture         =   "frmError.frx":06C6
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const LANG_NEUTRAL = &H0


Private Sub cmdExit_Click()
    On Local Error Resume Next
    End
End Sub

Private Sub cmdIgnore_Click()
    On Local Error Resume Next
    Unload Me
End Sub

Private Sub Form_Initialize()
    On Local Error Resume Next
    Call InitCommonControls
End Sub

Private Sub cmdDetails_Click()
    With Me
    If Right(cmdDetails.Caption, Len("»")) = "»" Then
        cmdDetails.Caption = "&Details «"
        .ScaleLeft = 0
        .ScaleTop = 0
        .ScaleHeight = Height
        .ScaleWidth = Width
        Me.Height = ScaleY(lblDetails.Top + lblDetails.Height, _
            vbTwips, vbTwips)
    Else
        cmdDetails.Caption = "&Details »"
        .ScaleLeft = 0
        .ScaleTop = 0
        .ScaleHeight = Height
        .ScaleWidth = Width
        Me.Height = ScaleY(lnBorder(1).Y1, _
            vbTwips, vbTwips)
    End If
    End With
    'Make buttons 3D
    Call FormatButtons(Me)
End Sub

Public Sub ShowDialog(ByVal ErrorNumber As Long, _
    Optional ByVal Procedure As String = "")
    'Get the last DLL error
    Dim lngDLLError As Long
    lngDLLError = GetLastError
    'Format it
    Dim strMessageDesc As String
    strMessageDesc = Space(200)
    Dim lngReturnValue As Long
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lngDLLError, LANG_NEUTRAL, strMessageDesc, 200, ByVal 0&
    If lngReturnValue <> 0 Then strMessageDesc = Left$(strMessageDesc, lngReturnValue + Len(vbNewLine))
    strMessageDesc = Replace(strMessageDesc, vbNullChar, "")
        
    'Get the error details
    Dim strOutput As String
    strOutput = "Error " & ErrorNumber & ": " & Error$(ErrorNumber) & vbNewLine & _
        "Error Source: " & App.EXEName & vbNewLine & _
        "Procedure: " & Procedure & vbNewLine & _
        "Last DLL Error: " & lngDLLError & ": " & strMessageDesc
    'Set them to the label's caption
    lblDetails.Caption = strOutput

    'Replace spaces and new lines with their HTML equivilants
    strOutput = Replace(strOutput, Space(1), "%20", , , vbTextCompare)
    strOutput = Replace(strOutput, vbNewLine, "%0D%0A ", , , vbTextCompare)
    'Add it to the link address
    lnkEmail.Address = lnkEmail.Address & "&body=" & strOutput
    'Show us
    Me.Show vbModal
End Sub

Attribute VB_Name = "BrowseForFolders"
Option Explicit
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Public Enum BIF_Flags
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_BROWSEINCLUDEURLS = &H80
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_EDITBOX = &H10
    BIF_NEWDIALOGSTYLE = &H40
    BIF_RETURNFSANCESTORS = &H8
    BIF_RETURNONLYFSDIRS = &H1
    BIF_SHAREABLE = &H8000
    BIF_STATUSTEXT = &H4
    BIF_USENEWUI = &H40
    BIF_VALIDATE = &H20
End Enum
Private Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function GetFolder(ByVal hWnd As Long, _
    Optional ByVal Message As String = "Please select the folder:", _
    Optional ByVal Flags As BIF_Flags = BIF_RETURNONLYFSDIRS) As String
    On Error Resume Next
    Dim ReturnValue As Long
    Dim BrowseOptions As BrowseInfo

    'Options for the dialog
    With BrowseOptions
        'Owner window
        .hWndOwner = hWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat(Message, "")
        'Set the flags
        .ulFlags = Flags
    End With

    'Show the Browse for folder dialog
    ReturnValue = SHBrowseForFolder(BrowseOptions)
    'If Cancel was not choosen
    If ReturnValue Then
        'Convert the return value to a string
        GetFolder = String$(MAX_PATH, 0)
        'To the path into GetFolder's return value
        Call SHGetPathFromIDList(ReturnValue, GetFolder)
        'Free memory used by the dialog
        CoTaskMemFree ReturnValue
        'Remove vbNullChartext after folder
        ReturnValue = InStr(GetFolder, vbNullChar)
        'If there is more text remove it
        If ReturnValue Then GetFolder = Left$(GetFolder, ReturnValue - 1)
    End If
End Function


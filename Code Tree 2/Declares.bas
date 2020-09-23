Attribute VB_Name = "Declares"
Option Explicit
'Types:
Public Type UndoInfo 'Type for the undo info
    SelStart As Long 'Start of selection
    SelLength As Long 'Length of selection
    Text As String 'The text
End Type
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long

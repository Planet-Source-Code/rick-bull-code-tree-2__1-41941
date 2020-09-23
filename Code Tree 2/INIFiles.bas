Attribute VB_Name = "INIFiles"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetINISetting(ByVal Filename As String, ByVal Section As String, _
    ByVal Key As String, Optional ByVal Default As String = "", _
    Optional BufferSize As Integer = 1023) As String
    On Local Error Resume Next
    Dim lngReturnValue As Long
  
    'Create a buffer for the string
    GetINISetting = String(BufferSize, 0)
    'Get the value, and store the return value
    lngReturnValue = GetPrivateProfileString(Section, Key, Default, _
        GetINISetting, Len(GetINISetting), Filename)
    'Remove trailing junk from the string if present
    If lngReturnValue <> 0 Then GetINISetting = Left$(GetINISetting, lngReturnValue)
End Function

Public Sub SaveINISetting(ByVal Filename As String, ByVal Section As String, _
    ByVal Key As String, Optional ByVal Value As String = "")
    On Local Error Resume Next
    'Write the setting
    Call WritePrivateProfileString(Section, Key, Value, Filename)
End Sub


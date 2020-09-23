Attribute VB_Name = "FileFunctions"
Option Explicit
'Enumerations for the type of output wanted
Public Enum OutputModeConsts
    Add = 0 'Append to the current text file
    OverWrite = 1 'Overwrite any existing text in the file
End Enum
Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long '  only used if FOF_SIMPLEPROGRESS
End Type
Public Declare Function SHFileOperation Lib "shell32.dll" _
    Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Enum FileOperations
    FO_DELETE = &H3 'Delete the file
    FO_MOVE = &H1 'Move the file
    FO_RENAME = &H4 'Rename the file
    FO_COPY = &H2 'Copy the file
End Enum
Public Enum FileOperationFlags
    FOF_ALLOWUNDO = &H40 'Prompt user to confirm
    FOF_NOCONFIRMATION = &H10 ' Don't prompt the user.
    FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed directories
    FOF_RENAMEONCOLLISION = &H8 'If files are same name rename the new one
    FOF_SILENT = &H4 ' don't create progress/report (indication of what's going on)
    FOF_SIMPLEPROGRESS = &H100 ' means don't show names of files
End Enum

Public Function OpenText(ByVal Filename As String, _
    Optional ByVal ShowError As Boolean = True) As String
    On Local Error GoTo ErrorHandler
    Dim FileNumber As Integer
    Dim TempText As String
                                        
    'Find a free file number
    FileNumber = FreeFile
    'Open the file for input
    Open Filename For Input As #FileNumber
        'Return the file's contents
        OpenText = Input(LOF(FileNumber), FileNumber)
    'Close the file
    Close #FileNumber

    'Exit the function so as not cause an error
    Exit Function

ErrorHandler:
    'Tell the user the error and ask if another method of opening should be tried
    If ShowError Then MsgBox "Sorry the file " & Filename & " could not be opened." & vbNewLine & _
        "Details: " & Err.Number & " - " & Err.Description, vbCritical + vbOKOnly, "Error"
End Function

Public Sub SaveText(ByVal Filename As String, ByVal TextToSave As String, _
    Optional ByVal OutputMode As OutputModeConsts = OverWrite)
    On Local Error GoTo ErrorHandler
    Dim FileNumber As Integer
    
    FileNumber = FreeFile
    'If the FileName is not ""
    If Filename <> "" Then
        If OutputMode = OverWrite Then
            'Open the for output
            Open Filename For Output As #FileNumber
                'Write the text to the file
                Print #FileNumber, TextToSave;
            'Close the file
            Close #FileNumber
            
        ElseIf OutputMode = Add Then
            'Open the for output
            Open Filename For Append As #FileNumber
                'Write the text to the file
                Print #FileNumber, TextToSave;
            'Close the file
            Close #FileNumber
        End If
    End If
    'Exit the sub so as not to cause an error
    Exit Sub

ErrorHandler:
    'Tell the user the error
    MsgBox "Sorry the file " & Filename & " could not be saved." & vbNewLine & _
        "Details: " & Err.Number & " - " & Err.Description, _
        vbCritical + vbOKOnly, "Error"
End Sub

Public Sub FileOperation(ByVal FromLocation As String, _
    Optional ByVal ToLocation As String, _
    Optional ByVal FunctionName As FileOperations = FO_DELETE, _
    Optional ByVal Flags As FileOperationFlags = FOF_ALLOWUNDO)
    On Local Error Resume Next
    Dim Operation As SHFILEOPSTRUCT
    With Operation
        .wFunc = FunctionName
        .pFrom = FromLocation
        .pTo = ToLocation
        .fFlags = Flags
    End With
    Call SHFileOperation(Operation)
End Sub


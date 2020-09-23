Attribute VB_Name = "FindAllFiles"
Option Explicit
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Const MAX_PATH = 260
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type
'Private Const strLongFileName = ""\\?\" 'Allows use of filenames longer than MAX_PATH

Public Function RemoveNullChars(ByVal Text As String) As String
    On Local Error Resume Next
    'Default return is the same text
    RemoveNullChars = Text
    'Find the null char
    Dim lngFound As Long
    lngFound = InStr(1, RemoveNullChars, vbNullChar, vbTextCompare)
    'If there is one, return what's to the left of it
    If lngFound > 0 Then RemoveNullChars = Left$(RemoveNullChars, lngFound - 1)
End Function

Public Sub FindFiles(ByVal Path As String, _
    ByVal CallerObject As Object, _
    ByVal AddFileSub As String, _
    Optional ByVal Pattern As String = "*.*", _
    Optional ByVal SubDirs As Boolean = True)
    On Local Error Resume Next
    'If no \ at end of path add it
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    
    'Find the first file
    Dim wfdInfo As WIN32_FIND_DATA
    Dim lngReturn As Long, lngContinue As Long
    lngContinue = 1
    lngReturn = FindFirstFile(Path & "*.*", wfdInfo)
    
    'Remove the null chars
    Dim strFileName As String, strFullName As String
    'Loop while still files/folders
    Do While lngContinue
        'Get rid of the null chars
        strFileName = RemoveNullChars(wfdInfo.cFileName)
        'Make the filename full (i.e. add the path)
        strFullName = Path & strFileName
        'If a valid folder
        If strFileName <> vbNullString And strFileName <> "." And _
            strFileName <> ".." Then
            'Call the add file sub
            If strFileName Like Pattern Or GetAttr(strFullName) And vbDirectory Then Call CallByName(CallerObject, AddFileSub, VbMethod, strFileName, strFullName)
            'If it's a folder, do this sub again for that dir so we get all _
             files/folders (will keep going until all folders have been done)
            If SubDirs And (GetAttr(strFullName) And vbDirectory) Then Call FindFiles(strFullName, CallerObject, AddFileSub, Pattern, SubDirs)
        End If
        'Find the next file and get the return value
        lngContinue = FindNextFile(lngReturn, wfdInfo)
    Loop
    Call FindClose(lngReturn)
End Sub


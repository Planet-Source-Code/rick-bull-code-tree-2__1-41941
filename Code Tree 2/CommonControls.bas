Attribute VB_Name = "CommonControls"
Option Explicit

'Public Sub AddNodes(ByVal TreeView As Object, _
'    ByVal Path As String, _
'    Optional ByVal Pattern As String = "*.*", _
'    Optional ByVal IncludeFolders As Boolean = True, _
'    Optional ByVal IncludeFiles As Boolean = True, _
'    Optional ByVal IncludeSubDirs As Boolean = True, _
'    Optional ByVal RemoveExtension As Boolean = True)
'    On Local Error Resume Next
'    Dim FileName As String, FileTitle As String, _
'        CurrentPath As String, CurrentPattern As String, _
'        NewFiles() As String
'    Dim Attributes As VbFileAttribute
'    Dim AddToArray As Boolean
'    Dim SplitPatterns() As String
'
'    'Append Path with \ if needed
'    If Right(Path, Len("\")) <> "\" Then Path = Path & "\"
'    'Split the patterns
'    SplitPatterns() = Split(Pattern, ";", , vbTextCompare)
'
'    'Folders:
'    If IncludeFolders Then
'        'Set the start details for the folders
'        CurrentPath = Path
'        CurrentPattern = "*"
'        Attributes = vbDirectory
'        'Means whether to add sub items once all the items in the start path have been found
'        If IncludeSubDirs Then AddToArray = True
'        'Add the firstfolders
'        GoSub AddFiles
'
'        Dim LoopCounter As Long, LoopCounter2 As Long 'Loop counters
'
'        'If there are new files
'        If LBound(NewFiles) < UBound(NewFiles) Then
'            'Don't add these files to the new array - only want 1 dir recurssion
'            AddToArray = False
'            'Loop for all new files
'            For LoopCounter = LBound(NewFiles) + 1 To UBound(NewFiles)
'                'Add the files in the new folders
'                '- done so we get a little + sign next to the folders if there are folders in them
'                CurrentPath = NewFiles(LoopCounter)
'                CurrentPattern = "*"
'                Attributes = vbDirectory
'                GoSub AddFiles
'
'                'If files are wanted
'                If IncludeFiles Then
'                    'Add them
'                    CurrentPath = NewFiles(LoopCounter)
'                    Attributes = vbNormal
'                    'Loop for all patterns
'                    For LoopCounter2 = LBound(SplitPatterns) To UBound(SplitPatterns)
'                        CurrentPattern = SplitPatterns(LoopCounter2)
'                        GoSub AddFiles
'                    Next LoopCounter2
'                End If
'            Next LoopCounter
'        End If
'    End If
'
'    'Files:
'    If IncludeFiles Then
'        'Add the files for the first path
'        CurrentPath = Path
'        Attributes = vbNormal
'        'Loop for all patterns
'        For LoopCounter = LBound(SplitPatterns) To UBound(SplitPatterns())
'            CurrentPattern = SplitPatterns(LoopCounter)
'            GoSub AddFiles
'        Next LoopCounter
'    End If
'    Exit Sub
'
'AddFiles:
'    'If files are to be added to the array reset the var to nothing
'    If AddToArray Then ReDim NewFiles(0)
'    'Get the first file
'    FileTitle = Dir(CurrentPath & CurrentPattern, Attributes)
'    'While there are more folders
'    Do While FileTitle <> ""
'        'If this folder is not the current folder (.) or the folder up (..)
'        If FileTitle <> "." And FileTitle <> ".." Then
'            'Get the full filename
'            FileName = CurrentPath & FileTitle
'            'Make sure we have the right file type
'            If FileName Like CurrentPath & CurrentPattern Then
'                'If it's a folder and we want folders
'                If Attributes = vbDirectory And GetAttr(FileName) = Attributes Then
'                        'Add a \ to the end so that we can tell a folder from a file
'                        FileName = FileName & "\"
'                        'If the small image is not already loaded load it
'                        If DoesListImageExist(TreeView.ImageList, FileName) = False Then _
'                            TreeView.ImageList.ListImages.Add , FileName, _
'                                LoadPicture(FileName & "Small.bmp")
'                        'If the large image is not already loaded load it
'                        If DoesListImageExist(frmMain.imlCodesLarge, FileName) = False Then _
'                            frmMain.imlCodesLarge.ListImages.Add , FileName, _
'                                LoadPicture(FileName & "Large.bmp")
'                End If
'
'                'Add this folder if it is a folder and that's what we want or we want files instead
'                If ((Attributes = vbDirectory And GetAttr(FileName) = Attributes) _
'                    Or Attributes <> vbDirectory) Then
'                    'If the file has not already been added
'                    If DoesNodeExist(TreeView, FileName) = False Then
'                        Dim IconKey As String, NodeCaption As String
'                        'Key = current directory if it is a folder, but parent if a file - saves added BMPs for every file
'                        IconKey = IIf(GetAttr(FileName) = vbDirectory, FileName, CurrentPath)
'                        'Caption = filetitle
'                        NodeCaption = FileTitle
'                        'If the current file is a file and not a folder and the extension is to be removed
'                        If Right(NodeCaption, Len("\")) <> "\" And RemoveExtension Then
'                            'Find the start of the extension (e.g. .rtf)
'                            Dim Found As Integer
'                            Found = InStrRev(NodeCaption, ".", , vbTextCompare)
'                            'Remove it if found
'                            If Found > 0 Then NodeCaption = Left(NodeCaption, Found - 1)
'                        End If
'                        'If we have a node for the current path
'                        If DoesNodeExist(TreeView, CurrentPath) Then
'                            'Add the file as a child
'                            TreeView.Nodes.Add CurrentPath, tvwChild, _
'                                FileName, NodeCaption, IconKey
'                        'If not
'                        Else
'                            'Create a new parent node
'                            TreeView.Nodes.Add , , FileName, NodeCaption, IconKey
'                        End If
'                    End If
'                    'If we are to add the files to the array
'                    If AddToArray Then
'                        'Add it
'                        ReDim Preserve NewFiles(0 To UBound(NewFiles) + 1)
'                        NewFiles(UBound(NewFiles)) = FileName
'                    End If
'                End If
'            End If
'        End If
'        'Find the next file/folder
'        FileTitle = Dir
'    Loop
'    Return
'End Sub

Public Function DoesNodeExist(ByVal TreeView As Object, Key) As Boolean
    'Resume next so as we can test for errors
    On Local Error Resume Next
    DoesNodeExist = True
    Dim TempVar As Long
    'Clear previous errors
    Err.Clear
    'Get the index of the sepcified node
    TempVar = TreeView.Nodes(Key).Index
    'Return true if there was no error and false if there was
    DoesNodeExist = Err.Number <= 0 '(Err.Number <> 35601 And Err.Number <> 35603) Or Err.Number = 35602
End Function

Public Function DoesListImageExist(ByVal ImageList As Object, _
    ByVal Key) As Boolean
    'Resume next so as we can test for errors
    On Local Error Resume Next
    DoesListImageExist = True
    
    'Clear previous errors
    Err.Clear
    'Get the list image's picture
    Dim ListImage As Picture
    Set ListImage = ImageList.ListImages(Key).Picture
    'Return true if the picture was gotten OK false if not
    DoesListImageExist = Err.Number <= 0 '(Err.Number <> 35601 And Err.Number <> 35603) Or Err.Number = 35602
End Function

Public Function HasSelectedItem(ByVal ListView As Object) As Boolean
    'Resume next so as we can test for errors
    On Local Error Resume Next
   
    'Clear previous errors
    Err.Clear
    'Get the selected item's index
    Dim Index As Long
    Index = ListView.SelectedItem.Index
    'Return True if no error false if there was
    HasSelectedItem = Err.Number = 0
End Function

Public Function DoesListItemExist(ByVal lwListView As Object, Key) As Boolean
    'Resume next so as we can test for errors
    On Local Error Resume Next
    DoesListItemExist = True
    Dim TempVar As Long
    
    'Clear previous errors
    Err.Clear
    'Get the list item's index
    TempVar = lwListView.ListItems(Key).Index 'ListView.ListItems(Key).Index
    'Return True if found false if not
    DoesListItemExist = Err.Number <= 0 '(Err.Number <> 35601 And Err.Number <> 35603) Or Err.Number = 35602
End Function

Public Sub LoadTabs(ByVal TabStrip As Object)
    On Local Error Resume Next
    'Loop for all tabs
    Dim LoopCounter As Integer
    For LoopCounter = 1 To TabStrip.Tabs.Count
        'If the image for this tab exists (based on Key's being the same) load it
        If DoesListImageExist(TabStrip.ImageList, _
            TabStrip.Tabs(LoopCounter).Key) Then _
                TabStrip.Tabs(LoopCounter).Image = _
                    TabStrip.Tabs(LoopCounter).Key
    Next LoopCounter
End Sub

Public Sub LoadToolbar(ByVal Toolbar As Object)
    On Local Error Resume Next
    'Loop for all buttons
    Dim LoopCounter As Integer
    For LoopCounter = 1 To Toolbar.Buttons.Count
        'If the image for this button exists (based on Key's being the same) load it
        If DoesListImageExist(Toolbar.ImageList, _
            Toolbar.Buttons(LoopCounter).Key) Then _
                Toolbar.Buttons(LoopCounter).Image = _
                    Toolbar.Buttons(LoopCounter).Key
    Next LoopCounter
End Sub

Public Function GetFileName(Optional ByVal Filter As String = _
    "Rich Text Files (*.rtf)|*.rtf|Text Files (*.txt)|*.txt|All Files|*.*", _
    Optional ByVal Save As Boolean = False) As String
    On Local Error GoTo ErrorHandler
        
    With frmMain.cdlDialogs
        .CancelError = True
        .Filter = Filter
        If Save Then
            .ShowSave
        Else
            .ShowOpen
        End If
        GetFileName = .Filename
    End With
    Exit Function
ErrorHandler:
    If Err.Number <> cdlCancel Then MsgBox "An error has occured", vbOKOnly Or vbExclamation, "Error"
End Function

Public Function GetColour() As Long
    On Local Error GoTo ErrorHandler
    GetColour = -1
    With frmMain.cdlDialogs
        .CancelError = True
        .Flags = cdlCCFullOpen
        .ShowColor
        GetColour = .Color
    End With
    Exit Function
ErrorHandler:
    If Err.Number <> cdlCancel Then MsgBox "An error has occured", vbOKOnly Or vbExclamation, "Error"
End Function


Public Sub LoadBitmaps(ByVal Toolbar As Object, ByVal Path As String, _
    Optional ByVal Extension As String = ".bmp")
    If Right(Path, Len("\")) <> "\" Then Path = Path & "\"
    'Loop for all toolbar's buttons
    Dim intLoopCounter As Integer
    For intLoopCounter = 1 To Toolbar.Buttons.Count
        'If not a placeholder or seperator (i.e. no image needed)
        If Toolbar.Buttons(intLoopCounter).Style <> tbrPlaceholder And _
            Toolbar.Buttons(intLoopCounter).Style <> tbrSeparator Then
            'Get the filename for the button
            Dim strFileName As String
            strFileName = Path & Toolbar.Buttons(intLoopCounter).Key & Extension
            'If the file exists load it
            If DoesFileExist(strFileName) Then Toolbar.ImageList.ListImages.Add , _
                Toolbar.Buttons(intLoopCounter).Key, LoadPicture(strFileName)
        End If
    Next intLoopCounter
End Sub

Public Function Pressed(ByVal TrueFalse As Boolean) As ValueConstants
    On Local Error Resume Next
    Pressed = IIf(TrueFalse = False, tbrUnpressed, tbrPressed)
End Function

Public Function BooleanPressed(ByVal Value As ValueConstants) As Boolean
    On Local Error Resume Next
    BooleanPressed = Value = tbrPressed
End Function


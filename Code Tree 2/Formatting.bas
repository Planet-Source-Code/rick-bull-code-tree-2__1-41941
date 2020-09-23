Attribute VB_Name = "Formatting"
Option Explicit
Public Type FormattingComments
    Bold() As Boolean
    Colour() As Long
    EndString() As String
    Font() As String
    FontSize() As Integer
    Italic() As Boolean
    StartString() As String
    StrikeThru() As Boolean
    Underline() As Boolean
End Type
Public Type FormattingDefault
    Bold As Boolean
    Colour As Long
    Font As String
    FontSize As Integer
    Italic As Boolean
    StrikeThru As Boolean
    Underline As Boolean
End Type
Public Type FormattingKeywords
    Bold As Boolean
    Colour As Long
    Delimeter As String
    Font As String
    FontSize As Integer
    Italic As Boolean
    Keywords() As String
    StrikeThru As Boolean
    Underline As Boolean
End Type
Public Type FormattingStrings
    Bold() As Boolean
    Colour() As Long
    EndString() As String
    Font() As String
    FontSize() As Integer
    Italic() As Boolean
    StartString() As String
    StrikeThru() As Boolean
    Underline() As Boolean
End Type
Public Type FormattingDetails
    Comments As FormattingComments
    Default As FormattingDefault
    Keywords As FormattingKeywords
    Strings As FormattingStrings
End Type
Public bolCancel  As Boolean 'Whether to leave the loop running
Public bolLoopRunning  As Boolean 'Whether the loop is running

Public Function GetConfig(ByVal ConfigFile As String, _
    Optional ByVal FormatEscape As Boolean = True) As FormattingDetails
    On Local Error Resume Next
    Dim lngLoopCounter As Long 'For loops
    Dim strTemp As String, intTemp As Integer 'Temp vars for extracting vaues to
    
    'COMMENTS
    With GetConfig.Comments
        'Get the amount of strings and set the arrays length
        intTemp = GetINISetting(ConfigFile, "Comments", "Count", 1)
        ReDim .Bold(1 To intTemp)
        ReDim .Colour(1 To intTemp)
        ReDim .EndString(1 To intTemp)
        ReDim .Font(1 To intTemp)
        ReDim .FontSize(1 To intTemp)
        ReDim .Italic(1 To intTemp)
        ReDim .StartString(1 To intTemp)
        ReDim .StrikeThru(1 To intTemp)
        ReDim .Underline(1 To intTemp)
        
        'Loop for all strings
        For lngLoopCounter = 1 To intTemp
            'Bold
            .Bold(lngLoopCounter) = GetINISetting(ConfigFile, "Comments", "FontBold" & lngLoopCounter, False)
            'Colour
            .Colour(lngLoopCounter) = GetColourFromString(GetINISetting(ConfigFile, "Comments", _
                "Colour" & lngLoopCounter, RGB(128, 0, 128)))
            'End
            .EndString(lngLoopCounter) = GetINISetting(ConfigFile, "Comments", _
                "End" & lngLoopCounter, """")
            If FormatEscape Then .EndString(lngLoopCounter) = Unescape(.EndString(lngLoopCounter))
            'Font Name
            .Font(lngLoopCounter) = GetINISetting(ConfigFile, "Comments", "Font" & lngLoopCounter, "Courier New")
            'Font Size
            .FontSize(lngLoopCounter) = Val(GetINISetting(ConfigFile, "Comments", "FontSize" & lngLoopCounter, 10))
            'Italic
            .Italic(lngLoopCounter) = GetINISetting(ConfigFile, "Comments", "FontItalic" & lngLoopCounter, False)
            'Start
            .StartString(lngLoopCounter) = GetINISetting(ConfigFile, "Comments", _
                "Start" & lngLoopCounter, """")
            If FormatEscape Then .StartString(lngLoopCounter) = Unescape(.StartString(lngLoopCounter))
            'Strike
            .StrikeThru(lngLoopCounter) = GetINISetting(ConfigFile, "Comments", "FontStrikeThru" & lngLoopCounter, False)
            'Underline
            .Underline(lngLoopCounter) = GetINISetting(ConfigFile, "Comments", "FontUnderline" & lngLoopCounter, False)
        Next lngLoopCounter
    End With

    'DEFAULT
    With GetConfig.Default
        'Bold
        .Bold = GetINISetting(ConfigFile, "Default", "FontBold", False)
        'Colour
        .Colour = GetColourFromString(GetINISetting(ConfigFile, "Default", _
            "Colour", vbWindowText))
        'Font Name
        .Font = GetINISetting(ConfigFile, "Default", "Font", "Courier New")
        'Font Size
        .FontSize = Val(GetINISetting(ConfigFile, "Default", "FontSize", 10))
        'Italic
        .Italic = GetINISetting(ConfigFile, "Default", "FontItalic", False)
        'Strike
        .StrikeThru = GetINISetting(ConfigFile, "Default", "FontStrikeThru", False)
        'Underline
        .Underline = GetINISetting(ConfigFile, "Default", "FontUnderline", False)
    End With

    'KEYWORDS
    With GetConfig.Keywords
        'Bold
        .Bold = GetINISetting(ConfigFile, "Keywords", "FontBold", False)
        'Colour
        strTemp = LCase(Trim(GetINISetting(ConfigFile, "Keywords", "Colour", RGB(0, 0, 128))))
        .Colour = GetColourFromString(strTemp, RGB(0, 0, 128))
        'Delimeter
        .Delimeter = GetINISetting(ConfigFile, "Keywords", "Delimeter")
        If FormatEscape Then .Delimeter = Unescape(.Delimeter)
        'Font Name
        .Font = GetINISetting(ConfigFile, "Keywords", "Font", "Courier New")
        'Font Size
        .FontSize = Val(GetINISetting(ConfigFile, "Keywords", "FontSize", 10))
        'Italic
        .Italic = GetINISetting(ConfigFile, "Keywords", "FontItalic", False)
        'Keywords
        strTemp = GetINISetting(ConfigFile, "Keywords", "Keywords")
        If FormatEscape Then strTemp = Unescape(strTemp)
        .Keywords() = Split(strTemp, .Delimeter, , vbTextCompare)
        'Strike
        .StrikeThru = GetINISetting(ConfigFile, "Keywords", "FontStrikeThru", False)
        'Underline
        .Underline = GetINISetting(ConfigFile, "Keywords", "FontUnderline", False)
    End With
    
    'STRINGS
    With GetConfig.Strings
        'Get the amount of strings and set the arrays length
        intTemp = GetINISetting(ConfigFile, "Strings", "Count", 1)
        'Redim the vars to accept all values
        ReDim .Bold(1 To intTemp)
        ReDim .Colour(1 To intTemp)
        ReDim .EndString(1 To intTemp)
        ReDim .Font(1 To intTemp)
        ReDim .FontSize(1 To intTemp)
        ReDim .Italic(1 To intTemp)
        ReDim .StartString(1 To intTemp)
        ReDim .StrikeThru(1 To intTemp)
        ReDim .Underline(1 To intTemp)

        'Loop for all strings
        For lngLoopCounter = 1 To intTemp
            'Bold
            .Bold(lngLoopCounter) = GetINISetting(ConfigFile, "Strings", "FontBold" & lngLoopCounter, False)
            'Colour
            .Colour(lngLoopCounter) = GetColourFromString(GetINISetting(ConfigFile, "Strings", _
                "Colour" & lngLoopCounter, RGB(128, 0, 128)))
            'End
            .EndString(lngLoopCounter) = GetINISetting(ConfigFile, "Strings", _
                "End" & lngLoopCounter, """")
            If FormatEscape Then .EndString(lngLoopCounter) = Unescape(.EndString(lngLoopCounter))
            'Font Name
            .Font(lngLoopCounter) = GetINISetting(ConfigFile, "Strings", _
                "Font" & lngLoopCounter, "Courier New")
            'Font Size
            .FontSize(lngLoopCounter) = Val(GetINISetting(ConfigFile, "Strings", _
                "FontSize" & lngLoopCounter, 10))
            'Italic
            .Italic(lngLoopCounter) = GetINISetting(ConfigFile, "Strings", "FontItalic" & lngLoopCounter, False)
            'Start
            .StartString(lngLoopCounter) = GetINISetting(ConfigFile, "Strings", _
                "Start" & lngLoopCounter, """")
            If FormatEscape Then .StartString(lngLoopCounter) = Unescape(.StartString(lngLoopCounter))
            'Strike
            .StrikeThru(lngLoopCounter) = GetINISetting(ConfigFile, "Strings", "FontStrikeThru" & lngLoopCounter, False)
            'Underline
            .Underline(lngLoopCounter) = GetINISetting(ConfigFile, "Strings", "FontUnderline" & lngLoopCounter, False)
        Next lngLoopCounter
    End With
End Function

Public Function FormatText(ByVal RTFBox As RichTextBox, _
    ByVal ConfigFile As String) As Boolean
    On Local Error Resume Next
    bolCancel = False
    bolLoopRunning = True
    FormatText = False
    'Get the configuration
    Dim fmdConfig As FormattingDetails
    fmdConfig = GetConfig(ConfigFile)
    
    'Get the text for quicker processing
    Dim Text As String
    Text = RTFBox.Text

    Dim lngTemp1 As Long, lngTemp2 As Long, lngTemp3 As Long, lngTemp4 As Long, _
        lngLength As Long 'Temp longs
    Dim strTemp  As String 'Temp string
    Dim lngLoopCounter As Long, lngLoopCounter2 As Long 'For loops
    lngLoopCounter = 1
    
    With RTFBox
        'Default text formatting
        .SelStart = 0
        .SelLength = Len(Text)
        'Load the progress bar form
        Load frmProgress
        'Set the progress
        frmProgress.pgbProgress.Max = .SelLength
        frmProgress.pgbProgress.Value = 0
        'Show the form
        frmProgress.Show vbModeless, frmMain
        .SelAlignment = vbLeftJustify
        .SelBold = fmdConfig.Default.Bold
        .SelBullet = False
        .SelColor = fmdConfig.Default.Colour
        .SelFontName = fmdConfig.Default.Font
        .SelFontSize = fmdConfig.Default.FontSize
        .SelItalic = fmdConfig.Default.Italic
        .SelStrikeThru = fmdConfig.Default.StrikeThru
        .SelUnderline = fmdConfig.Default.Underline
        
        'Loop for all text
        Do While lngLoopCounter < Len(Text) And bolCancel = False
            Dim intStringNumber As Integer, intCommentsNumber As Integer
            'Find if we have a string/keyword
            intStringNumber = GetArrayIndex(Text, lngLoopCounter, _
                fmdConfig.Strings.StartString)
            intCommentsNumber = GetArrayIndex(Text, lngLoopCounter, _
                fmdConfig.Comments.StartString)
            
            'If we have a string
            If intStringNumber > -1 Then
                'Find the end of it
                lngTemp1 = InStr(lngLoopCounter + Len(fmdConfig.Strings.StartString(intStringNumber)), Text, fmdConfig.Strings.EndString(intStringNumber), vbTextCompare)
                'If it's not found make equal to the length of the text
                If lngTemp1 <= 0 Then lngTemp1 = Len(Text)
                
                'Select string
                .SelStart = lngLoopCounter + Len(fmdConfig.Strings.StartString(intStringNumber)) - 1
                .SelLength = lngTemp1 - .SelStart - 1
                'Apply the style to it
                .SelBold = fmdConfig.Strings.Bold(intStringNumber)
                .SelColor = fmdConfig.Strings.Colour(intStringNumber)
                .SelFontName = fmdConfig.Strings.Font(intStringNumber)
                .SelFontSize = fmdConfig.Strings.FontSize(intStringNumber)
                .SelItalic = fmdConfig.Strings.Italic(intStringNumber)
                .SelStrikeThru = fmdConfig.Strings.StrikeThru(intStringNumber)
                .SelUnderline = fmdConfig.Strings.Underline(intStringNumber)
                'Increment loop counter to take to the end of the string
                lngLoopCounter = lngTemp1 + Len(fmdConfig.Strings.EndString(intStringNumber))
    
            'If we have a comment
            ElseIf intCommentsNumber > -1 Then
                'Find the end of the comment
                lngTemp1 = InStr(lngLoopCounter + Len(fmdConfig.Comments.StartString(intCommentsNumber)), Text, fmdConfig.Comments.EndString(intCommentsNumber), vbTextCompare)
                'If it's not found make equal the length of the text
                If lngTemp1 <= 0 Then lngTemp1 = Len(Text)
                
                'Select it
                .SelStart = lngLoopCounter - 1 '+ Len(strCommentsStart(intCommentsNumber)) - 2
                .SelLength = lngTemp1 + Len(fmdConfig.Comments.EndString(intCommentsNumber)) - .SelStart - 1
                'Aookt the style to it
                .SelBold = fmdConfig.Comments.Bold(intCommentsNumber)
                .SelColor = fmdConfig.Comments.Colour(intCommentsNumber)
                .SelFontName = fmdConfig.Comments.Font(intCommentsNumber)
                .SelFontSize = fmdConfig.Comments.FontSize(intCommentsNumber)
                .SelItalic = fmdConfig.Comments.Italic(intCommentsNumber)
                .SelStrikeThru = fmdConfig.Comments.StrikeThru(intCommentsNumber)
                .SelUnderline = fmdConfig.Comments.Underline(intCommentsNumber)
                
                'Increment the loop counter to take to the end of the comment
                lngLoopCounter = lngTemp1 + Len(fmdConfig.Comments.EndString(intCommentsNumber))
            
            'Other
            ElseIf IsArray(fmdConfig.Keywords.Keywords) Then
                'If we are not at the start
                Dim bolShouldFormat As Boolean
                If lngLoopCounter > 1 Then
                    'We can format if this is not within a word
                    bolShouldFormat = Not IsAlphaNumeric(Mid(Text, lngLoopCounter - 1, 1))
                'If we are at the start then we can format anyway
                Else
                    bolShouldFormat = True
                End If
                
                'If we can format
                If bolShouldFormat Then
                    'Find the end character of the next word
                    lngLength = GetFirstNonAlphaNumeric(Text, lngLoopCounter) ' GetLowest(1, Len(Text) + 1, _
                        InStr(lngLoopCounter, Text, Space(1), vbTextCompare), _
                        InStr(lngLoopCounter, Text, vbNewLine, vbTextCompare), _
                        InStr(lngLoopCounter, Text, ",", vbTextCompare), _
                        InStr(lngLoopCounter, Text, ")", vbTextCompare), _
                        InStr(lngLoopCounter, Text, ">", vbTextCompare), _
                        InStr(lngLoopCounter, Text, ".", vbTextCompare), _
                        InStr(lngLoopCounter, Text, ";", vbTextCompare))
                    
                    
                    'Get the current word
                    strTemp = LCase(Mid(Text, lngLoopCounter, lngLength - lngLoopCounter))
                    'Loop for all keywords
                    For lngLoopCounter2 = 0 To UBound(fmdConfig.Keywords.Keywords)
                        'If we have a keyword
                        If LCase(fmdConfig.Keywords.Keywords(lngLoopCounter2)) = strTemp Then
                            'Select it
                            .SelStart = lngLoopCounter - 1
                            .SelLength = lngLength - lngLoopCounter
                            'Apply the style to it
                            .SelBold = fmdConfig.Keywords.Bold
                            .SelColor = fmdConfig.Keywords.Colour
                            .SelFontName = fmdConfig.Keywords.Font
                            .SelFontSize = fmdConfig.Keywords.FontSize
                            .SelItalic = fmdConfig.Keywords.Italic
                            .SelStrikeThru = fmdConfig.Keywords.StrikeThru
                            .SelUnderline = fmdConfig.Keywords.Underline
                            .SelText = fmdConfig.Keywords.Keywords(lngLoopCounter2)
                            
                            'Increment loop counter to take us to the end of this keyword
                            lngLoopCounter = (lngLoopCounter) + (lngLength - lngLoopCounter) ' used to be _
                                (lngLoopCounter - 1) + (lngLength - lngLoopCounter)
                            'Exit loop for effeciency - frequently used keywords should be first in the config file
                            Exit For
                        Else
                            If lngLoopCounter2 >= UBound(fmdConfig.Keywords.Keywords) Then
                            'Move the loop counter up one
                            lngLoopCounter = lngLoopCounter + 1
                            End If
                        End If
                    Next lngLoopCounter2
                Else
                    'Move the loop counter up one
                    lngLoopCounter = lngLoopCounter + 1
                End If
            End If
            Call frmProgress.SetValue(lngLoopCounter)
            bolCancel = IsButtonActive(VK_ESCAPE, frmProgress.hWnd)
            DoEvents
        Loop
    End With
    FormatText = True
    GoTo ExitFunction
    Exit Function

ExitFunction:
    Unload frmProgress
    bolLoopRunning = False
End Function

Private Function GetArrayIndex(ByVal Text As String, ByVal Start As Long, _
    ByRef StringsArray() As String) As Integer
    On Local Error Resume Next
    'Default return value = -1
    GetArrayIndex = -1
    'Loop for all array entries while we haven't found the selected one
    Dim intLoopCounter As Integer
    Do While intLoopCounter < UBound(StringsArray) And GetArrayIndex = -1
        'Increment the counter
        intLoopCounter = intLoopCounter + 1
        'If we have found the selected one in the array return it's array index
        If LCase(Mid(Text, Start, Len(StringsArray(intLoopCounter)))) = LCase(StringsArray(intLoopCounter)) Then _
            GetArrayIndex = intLoopCounter
    Loop
End Function

Public Function GetColourFromString(ByVal Text As String, _
    Optional ByVal Default As Long = 0&) As Long
    On Local Error GoTo ErrorHandler
    'Remove the spaces around the text
    Text = Trim(Text)
    'If it's a RGB colour
    If Left(Text, Len("rgb")) = "rgb" Then
        'Find the start/end of the brackets
        Dim lngFound1 As Long, lngFound2 As Long
        lngFound1 = InStr(Len("rgb"), Text, "(", vbTextCompare)
        lngFound2 = InStr(lngFound1, Text, ")", vbTextCompare)
        'If we've found both brackets
        If lngFound1 > 0 And lngFound2 > 0 Then
            'Split the bit between the brackets by commas (,)
            Dim strRGB() As String
            strRGB() = Split(Mid(Text, lngFound1 + 1, lngFound2), ",", , vbTextCompare)
            'If we have at least 3 values
            If UBound(strRGB) >= 2 Then
                'Get the long colour of the RGB values
                GetColourFromString = RGB(Val(strRGB(0)), Val(strRGB(1)), Val(strRGB(2)))
            'If we haven't got enough values
            Else
                'Default colour
                GetColourFromString = Default
            End If
        'If we haven't got both brackets
        Else
            'Default colour
            GetColourFromString = Default
        End If
    'Normal long colour
    Else
        'Get the value
        GetColourFromString = Val(Text)
    End If
    Exit Function

ErrorHandler:
    'Return the default on error
    GetColourFromString = Default
End Function

Private Function Unescape(ByVal Text As String) As String
    On Local Error Resume Next
    Dim lngLoopCounter As Long
    'Loop for length of string
    lngLoopCounter = 1
    Do While lngLoopCounter <= Len(Text)
        'If we have an  escape character
        If Mid(Text, lngLoopCounter, 1) = "\" Then
            'Choose which one
            Select Case LCase(Mid(Text, lngLoopCounter + 1, 1))
                '\f = Line Feed
                Case "f"
                    Unescape = Unescape & vbLf
                '\n = New Line
                Case "n"
                    Unescape = Unescape & vbNewLine
                '\r = Carriage Return
                Case "r"
                    Unescape = Unescape & vbCr
                '\t = Tab
                Case "t"
                    Unescape = Unescape & vbTab
                '\v = Vertical Tab
                Case "v"
                    Unescape = Unescape & vbVerticalTab
                '\' = '
                Case "'"
                    Unescape = Unescape & "'"
                '\" = "
                Case """"
                    Unescape = Unescape & """"
                '\\ = \
                Case "\"
                    Unescape = Unescape & "\"
            End Select
            'Increment past these characters
            lngLoopCounter = lngLoopCounter + 2
        
        'Not an escape character
        Else
            'Add this char and increment the counter by 1
            Unescape = Unescape & Mid(Text, lngLoopCounter, 1)
            lngLoopCounter = lngLoopCounter + 1
        End If
    Loop
End Function

Public Function IsAlphaNumeric(ByVal Text As String) As Boolean
    On Local Error Resume Next
    Dim intAscValue As Integer

    'Loop for all chars in text
    Dim lngLoopCounter As Integer
    For lngLoopCounter = 1 To Len(Text)
        'Get the ASCII number of the string
        intAscValue = Asc(Mid(Text, lngLoopCounter, 1))
        'Return true only if this char is A-Z, a-z or 0-9
        IsAlphaNumeric = (intAscValue >= vbKeyA And intAscValue <= vbKeyZ) Or _
            (intAscValue >= vbKeyA + 32 And intAscValue <= vbKeyZ + 32) Or _
            (intAscValue >= vbKey0 And intAscValue <= vbKey9)
        'If not AlphaNumeric exit the loop
        If IsAlphaNumeric = False Then Exit For
    Next lngLoopCounter
 End Function


Public Function GetConfigFiles(ByVal ConfigPath As String) As String()
    On Local Error GoTo ErrorHandler
    'What gets returned at the end
    Dim strReturn() As String
    
    'No arrays yet
    Dim bolInitialised As Boolean
    bolInitialised = False
    
    'Get the first file in the formatting directory
    Dim strFileName As String
    strFileName = Dir(ConfigPath & "*.*")
    'Loop for all files
    Do While strFileName <> vbNullString
        'If we don't already have some options loaded
        If bolInitialised = False Then
            'Create the array
            ReDim strReturn(0 To 0)
            'Initialised = True
            bolInitialised = True
        Else
            ReDim Preserve strReturn(0 To UBound(strReturn) + 1)
        End If
        'Set the new menu's caption and show it
        strReturn(UBound(strReturn)) = strFileName
        'Get the next file
        strFileName = Dir
    Loop
    GetConfigFiles = strReturn()
    Exit Function
ErrorHandler:
End Function

'Private Function GetLowest(ByVal Min As Long, _
'    ByVal Default As Long, ParamArray Number()) As Long
'    On Local Error Resume Next
'    GetLowest = Default
'    Dim intLoopCounter As Integer
'    For intLoopCounter = LBound(Number) To UBound(Number)
'        If GetLowest > Number(intLoopCounter) And _
'            Number(intLoopCounter) >= Min Then _
'            GetLowest = Number(intLoopCounter)
'    Next intLoopCounter
'End Function

Private Function GetFirstNonAlphaNumeric(Text As String, Start As Long) As Long
    On Local Error Resume Next
    'Loop Counter = start position
    Dim lngLoopCounter As Integer
    lngLoopCounter = Start
    'Default return val = length of text
    GetFirstNonAlphaNumeric = Len(Text)
    'Do while there is more text to start and we have not already found the char
    Do While lngLoopCounter < Len(Text) And GetFirstNonAlphaNumeric = Len(Text)
        'If the character isn't AlphaNumeric then return this position (this will make the loop stop)
        If IsAlphaNumeric(Mid(Text, lngLoopCounter, 1)) = False Then _
            GetFirstNonAlphaNumeric = lngLoopCounter
        'Increment the loopcounter
        lngLoopCounter = lngLoopCounter + 1
    Loop
End Function

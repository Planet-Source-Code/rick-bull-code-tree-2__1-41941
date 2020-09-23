Attribute VB_Name = "Case"
Option Explicit
Public Enum CaseConstants
    [lower case] = 1
    [UPPER CASE] = 2
    [tOGGLE cASE] = 3
    [Proper Case] = 4
    [Sentance case] = 5
    [VaRy cAsE 1] = 6
    [vArY CaSe 2] = 7
End Enum

Public Function ChangeCase(ByVal Text As String, _
    ByVal NewCase As CaseConstants, _
    Optional ByVal Start As Long = 1, _
    Optional ByVal Limit As Long = -1) As String
    On Local Error GoTo ErrorHandler
    'If no limit (no no limit, no no limit, no no there's not limit, ahem)
    'make it = the length of the text
    If Limit = -1 Then Limit = Len(Text)
    
    Dim strBefore As String, strMain As String, _
        strAfter As String 'Strings for below usage
    'Start = all before START
    strBefore = Mid(Text, 1, Start - 1)
    'Text to convert
    strMain = Mid(Text, Start, Limit - (Start - 1))
    'Text after conversion  text
    strAfter = Mid(Text, Limit + 1)
    
    Dim strTemp As String 'Temp string for various things
    Dim lngLoopCounter As Integer
    
    'Select the case
    Select Case NewCase
        'LOWER CASE
        Case [lower case]
            strMain = LCase(strMain)
        
        'UPPER CASE
        Case [UPPER CASE]
            strMain = UCase(strMain)
        
        'tOGGLE cASE
        Case [tOGGLE cASE]
            'Loop for all text
            For lngLoopCounter = 1 To Len(strMain)
                'If the text is already uppercase
                If UCase(Mid(strMain, lngLoopCounter, 1)) = Mid(strMain, lngLoopCounter, 1) Then
                    'Make it lowercase
                    strTemp = strTemp & LCase(Mid(strMain, lngLoopCounter, 1))
                'If it's lowercase
                Else
                    'Make it uppercase
                    strTemp = strTemp & UCase(Mid(strMain, lngLoopCounter, 1))
                End If
            Next lngLoopCounter
            'Make strMain = the converted text
            strMain = strTemp
            
        'Sentance Case
        Case [Sentance case]
            'Lower case all
            strMain = LCase(strMain)
            Dim bolCapitol As Boolean
            bolCapitol = True
            Dim strTemp2 As String
            Dim intTemp As Integer
            intTemp = Asc(strTemp2)
            'Loop for all chars
            For lngLoopCounter = 1 To Len(strMain)
                strTemp2 = Mid(strMain, lngLoopCounter, 1)
                'If this one need to be capitol (and not just a white space char)
                If bolCapitol And intTemp > 32 Then
                    'Uppercase it
                    strTemp = strTemp + UCase(strTemp2)
                    'Already capitoled, so don't capitol the next one
                    bolCapitol = False
                'Return this char (already lower case so don't do it again)
                Else
                    strTemp = strTemp + strTemp2
                End If
                'If this char is a full stop, question mark or exclamation mark then the next non-whitespace char should be uppercase
                If strTemp2 = "." Or strTemp2 = "?" Or strTemp2 = "!" Then bolCapitol = True
            Next lngLoopCounter
            'Return this new stuff
            strMain = strTemp
            
        'Proper case
        Case [Proper Case]
            strMain = StrConv(strMain, vbProperCase)
            
        'VaRy cAsE 1
        Case [VaRy cAsE 1]
            'Loop for all text, incrementing by 2 each time
            For lngLoopCounter = 1 To Len(strMain) Step 2
                'Return uppercase of the current char and lowercase of the one after (if present)
                strTemp = strTemp & UCase(Mid(strMain, lngLoopCounter, 1)) & _
                    IIf(lngLoopCounter + 1 <= Len(strMain), LCase(Mid(strMain, lngLoopCounter + 1, 1)), "")
            Next lngLoopCounter
            'Make strMain = the converted text
            strMain = strTemp
        
        'VaRy cAsE 1
        Case [vArY CaSe 2]
            'Loop for all text, incrementing by 2 each time
            For lngLoopCounter = 1 To Len(strMain) Step 2
                'Return lowercase of the current char and uppercase of the one after (if present)
                strTemp = strTemp & LCase(Mid(strMain, lngLoopCounter, 1)) & _
                    IIf(lngLoopCounter + 1 <= Len(strMain), UCase(Mid(strMain, lngLoopCounter + 1, 1)), "")
            Next lngLoopCounter
            'Make strMain = the converted text
            strMain = strTemp
        
        'Anything unrecognised
        Case Else
            'Return default text
            GoTo ErrorHandler
    End Select
    
    'Return the start, middle (converted text) and end text
    ChangeCase = strBefore & strMain & strAfter
    'Exit so we don't return original text
    Exit Function

ErrorHandler:
    'Return original text
    ChangeCase = Text
End Function
'Public Function IsAlphaNumeric(ByVal Text As String) As Boolean
'    'Uppercase text for ease of comparison
'    Text = UCase(Text)
'    Dim intAscValue As Integer 'ASCII value of characters
'    'Loop for all text
'    Dim lngLoopCounter As Long
'    For lngLoopCounter = 1 To Len(Text)
'        'Get the ASCII value of the current character
'        intAscValue = Asc(Mid(Text, lngLoopCounter, 1))
'        'Return true if A-Z or 0-9
'        IsAlphaNumeric = (intAscValue >= vbKeyA And intAscValue <= vbKeyZ) Or _
'            (intAscValue >= vbKey0 And intAscValue <= vbKey9)
'        'Exit if not Alpha as later loops may change this
'        If IsAlphaNumeric = False Then Exit For
'    Next lngLoopCounter
'End Function

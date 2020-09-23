Attribute VB_Name = "Sound"
'API Declarations
#If Win32 Then '32-Bit windows
    Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long 'The API that lets sound be played
#Else 'Other Windows
    Public Declare Function sndPlaySound Lib "MMSYSTEM.DLL" _
        (ByVal lpszSoundName As Any, ByVal wFlags As Integer) As Integer 'The API that lets sound be played
#End If
'API Constants Used for flags in PlaySound function
Public Enum SoundFlags
    SND_LOOP = &H8 ' Loop the sound until next sndPlaySound
    SND_ALIAS = &H10000 ' Name is a WIN.INI [sounds] entry
    SND_ALIAS_ID = &H110000 ' Name is a WIN.INI [sounds] entry identifier
    SND_ALIAS_START = 0 ' Must be > 4096 to keep strings in same section of resource file
    SND_APPLICATION = &H80 ' Look for application specific association
    SND_ASYNC = &H1 ' Play asynchronously
    SND_FILENAME = &H20000 ' Name is a file name
    SND_MEMORY = &H4 ' lpszSoundName points to a memory file
    SND_NODEFAULT = &H2 ' Silence not default, if sound not found
    SND_NOSTOP = &H10 ' Don't stop any currently playing sound
    SND_NOWAIT = &H2000 ' Don't wait if the driver is busy
    SND_PURGE = &H40 ' Purge non-static events for task
    SND_RESERVED = &HFF000000 ' In particular these flags are reserved
    SND_RESOURCE = &H40004 ' Name is a resource name or atom
    SND_SYNC = &H0 ' Play synchronously (default)
    SND_TYPE_MASK = &H170007
    SND_VALID = &H1F ' Valid flags          / ;Internal /
    SND_VALIDFLAGS = &H17201F ' Set of valid flag bits.  Anything outside
End Enum

Private Sub PlaySound(ByVal FileName As String, _
    Optional Flags As SoundFlags = SND_FILENAME Or SND_NOSTOP Or SND_ASYNC)
    On Local Error Resume Next
    'If the filename is not "" then play the sound
    If FileName <> "" Then Call sndPlaySound(FileName, Flags)
End Sub


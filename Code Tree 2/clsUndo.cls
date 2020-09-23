VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUndo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'PUT THIS IN A MODULE:
'Public Type UndoInfo 'Type for the undo info
'    SelStart As Long 'Start of selection
'    SelLength As Long 'Length of selection
'    Text As String 'The text
'End Type
Private intUndoUbound As Integer 'The upper bound of the udiUndo array
Private udiUndo() As UndoInfo 'Holds the undo info
Private intRedoUbound As Integer 'The upper bound of the udiRedo array
Private udiRedo() As UndoInfo 'Holds the redo info
Private bolUndoing As Boolean 'Whether the calling module is undoing
Private intMaxUndos As Integer 'How many undos to store

Friend Sub AddToBuffer(ByVal Text As String, ByVal SelStart As String, _
    ByVal SelLength As String)
    On Local Error Resume Next
    'Exit if we are undoing
    If bolUndoing Then Exit Sub
    'If the Undo buffer is unitialised
    If intUndoUbound <= -1 Then
        'Initialise it
        intUndoUbound = 0
        ReDim udiUndo(0 To intUndoUbound)
    'If it already has some details
    Else
        'Get the top array number + 1
        If intUndoUbound < intMaxUndos Then intUndoUbound = intUndoUbound + 1
        'Add a new array
        ReDim Preserve udiUndo(0 To intUndoUbound)
        'Loop backwards for all arrays minus the first
        Dim intLoopCounter As Integer
        For intLoopCounter = intUndoUbound To 1 Step -1
            'Make this one = the last one (i.e. move them all down one
            With udiUndo(intLoopCounter)
                .SelLength = udiUndo(intLoopCounter - 1).SelLength
                .SelStart = udiUndo(intLoopCounter - 1).SelStart
                .Text = udiUndo(intLoopCounter - 1).Text
            End With
        Next intLoopCounter
    End If
    'Add the new one to the top
    With udiUndo(0)
        .SelLength = SelLength
        .SelStart = SelStart
        .Text = Text
    End With
    'Kill the old Redos
    Erase udiRedo
    intRedoUbound = -1
End Sub

Friend Function GetUndo() As UndoInfo
    On Local Error Resume Next
    'Return the first undo in the array
    With GetUndo
        .SelLength = udiUndo(0).SelLength
        .SelStart = udiUndo(0).SelStart
        .Text = udiUndo(0).Text
    End With
    'Take one off of the undo's ubound
    intUndoUbound = intUndoUbound - 1
    'If we can shuffle the array
    If intUndoUbound > -1 Then
        'Loop for all the undos (minus the last as we are going to kill it)
        Dim intLoopCounter As Integer
        For intLoopCounter = 0 To intUndoUbound
            'Make this one = the next one, i.e. move them all back one
            With udiUndo(intLoopCounter)
                .SelLength = udiUndo(intLoopCounter + 1).SelLength
                .SelStart = udiUndo(intLoopCounter + 1).SelStart
                .Text = udiUndo(intLoopCounter + 1).Text
            End With
        Next intLoopCounter
        'Remove the last undo as it has the same as the previous one in it
        ReDim Preserve udiUndo(0 To intUndoUbound)
    
    'No more undos
    Else
        'Kill the undo buffer
        Erase udiUndo
    End If
    
    'Add a new space to the redo buffer
    intRedoUbound = intRedoUbound + 1
    ReDim Preserve udiRedo(0 To intRedoUbound)
    'Loop backwards for all but the first buffers
    For intLoopCounter = intRedoUbound To 1 Step -1
        'Male this one = the last one, i.e. move them all forward one place
        With udiRedo(intLoopCounter)
            .SelLength = udiRedo(intLoopCounter - 1).SelLength
            .SelStart = udiRedo(intLoopCounter - 1).SelStart
            .Text = udiRedo(intLoopCounter - 1).Text
        End With
    Next intLoopCounter
    'Add this undo to the first redo
    With udiRedo(0)
        .SelLength = GetUndo.SelLength
        .SelStart = GetUndo.SelStart
        .Text = GetUndo.Text
    End With
End Function

Friend Function GetRedo(Optional ByVal RemoveFromBuffer As Boolean = True) As UndoInfo
    On Local Error Resume Next
    'Return the first redo in the array
    With GetRedo
        .SelLength = udiRedo(0).SelLength
        .SelStart = udiRedo(0).SelStart
        .Text = udiRedo(0).Text
    
        'Remove this from the buffer is wanted
        If RemoveFromBuffer Then
            'Make undo array one bigger
            intUndoUbound = intUndoUbound + 1
            ReDim Preserve udiUndo(0 To intUndoUbound)
            
            'Loop backwards for all all undos minus the first
            Dim intLoopCounter As Integer
            For intLoopCounter = intUndoUbound To 1 Step -1
                'Make this one = the last one, i.e. move the all down one place
                With udiUndo(intLoopCounter)
                    .SelLength = udiUndo(intLoopCounter - 1).SelLength
                    .SelStart = udiUndo(intLoopCounter - 1).SelStart
                    .Text = udiUndo(intLoopCounter - 1).Text
                End With
            Next intLoopCounter
            'Make the first undo array = the redo returned
            With udiUndo(0)
                .SelLength = GetRedo.SelLength
                .SelStart = GetRedo.SelStart
                .Text = GetRedo.Text
            End With
            
            'Redo array is one smaller
            intRedoUbound = intRedoUbound - 1
            'If we can shuffle arrays
            If intRedoUbound > -1 Then
                'Loop for all redos
                For intLoopCounter = 0 To intRedoUbound
                        'Move this one = the next one, i.e. move all back one
                        With udiRedo(intLoopCounter)
                            .SelLength = udiRedo(intLoopCounter + 1).SelLength
                            .SelStart = udiRedo(intLoopCounter + 1).SelStart
                            .Text = udiRedo(intLoopCounter + 1).Text
                        End With
                    Next intLoopCounter
                'Remove the last array
                ReDim Preserve udiRedo(0 To intRedoUbound)
            
            'If we have no more redos left
            Else
                'Kill the redo array
                Erase udiRedo
            End If
        End If
    End With
End Function

Friend Sub Reset()
    On Local Error Resume Next
    'Kill the Undo/Redo arrays
    intUndoUbound = -1
    Erase udiUndo
    intRedoUbound = -1
    Erase udiRedo
End Sub

Private Sub Class_Initialize()
    On Local Error Resume Next
    'User not undoing, up to the programmer to use if they want
    bolUndoing = False
    'Initial ubound = uninitialised
    intUndoUbound = -1
    intRedoUbound = -1
    'Default Max undos = 99
    intMaxUndos = 99
End Sub

Friend Function CanUndo() As Boolean
    On Local Error Resume Next
    'Return whether we can undo
    CanUndo = intUndoUbound >= (0)
End Function

Friend Function CanRedo() As Boolean
    On Local Error Resume Next
    'Return whether we can undo
    CanRedo = intRedoUbound >= (0)
End Function

Public Property Get Undoing() As Boolean
    On Local Error Resume Next
    'Return whether we are in the middle of undoing, so don't add to buffer
    Undoing = bolUndoing
End Property

Public Property Let Undoing(ByVal New_Undoing As Boolean)
    On Local Error Resume Next
    'Set the new undoing
    bolUndoing = New_Undoing
End Property

Public Property Get MaxUndos() As Integer
    On Local Error Resume Next
    'Return how many undos we can perform
    MaxUndos = intMaxUndos
End Property

Public Property Let MaxUndos(ByVal New_MaxUndos As Integer)
    On Local Error Resume Next
    'Set how many undos we can perform
    intMaxUndos = New_MaxUndos
    'Delete the top arrays if too many
    If UBound(udiRedo) > New_MaxUndos Then ReDim udiRedo(0 To New_MaxUndos)
    If UBound(udiUndo) > New_MaxUndos Then ReDim udiUndo(0 To New_MaxUndos)
End Property

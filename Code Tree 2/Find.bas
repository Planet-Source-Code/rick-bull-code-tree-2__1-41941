Attribute VB_Name = "Find"
Option Explicit
Public Enum SearchInConsts
    [Tree] = 0
    [Code] = 1
    [Notes] = 2
    [Bookmarks] = 3
End Enum
Private cpmLastCompareMethod As VbCompareMethod
Private strLastFind As String, strLastReplace As String
Public bolBeepOnFind As Boolean

Public Sub FindIn(ByVal SearchIn As SearchInConsts, _
    Optional ByVal Find As String = vbNullChar, _
    Optional ByVal Replace As String = vbNullChar, _
    Optional ByVal CompareMethod As VbCompareMethod = -1)
    On Local Error Resume Next
    Dim lngLoopCounter As Long, lngFound As Long
    Static lngLastTree As Long, lngLastCode As Long, lngLastNotes As Long, lngLastBookmarks As Long
    If Find <> vbNullChar Then strLastFind = Find
    If Replace <> vbNullChar Then strLastReplace = Replace
    If CompareMethod <> -1 Then cpmLastCompareMethod = CompareMethod
    
    'Dim intMatchMode As VbCompareMethod
Start:
    'Get the case compare mode
    'intMatchMode = IIf(chkMatchCase.Value = vbChecked, vbBinaryCompare, vbTextCompare)
    'Tree
    If SearchIn = Tree Then
        'Increment the node to search from so we don't stay on the same one
        lngLastTree = lngLastTree + 1
        With frmMain.tvwCodes
            'If all nodes have been searched
            If lngLastTree >= .Nodes.Count Then
                'Reset and tell user
                lngLastTree = 1
                Call MsgBox("Finished searching through tree nodes with no match.", _
                    vbOKOnly Or vbInformation, "Finished")
                Exit Sub
            End If
            'Loop for all nodes from the last one found (+1)
            For lngLoopCounter = lngLastTree To .Nodes.Count
                'Search for the string
                lngFound = InStr(1, .Nodes(lngLoopCounter).Text, strLastFind, cpmLastCompareMethod)
                'ADD WHITE SPACE CHECK HERE!
                'If the search string is found
                If lngFound > 0 Then
                    'Set this one as the last one and select the item
                    lngLastTree = lngLoopCounter
                    .Nodes(lngLoopCounter).Expanded = True
                    .Nodes(lngLoopCounter).Selected = True
                    .Nodes(lngLoopCounter).EnsureVisible
                    Call frmMain.tvwCodes_Expand(.Nodes(lngLoopCounter))
                    Call frmMain.tvwCodes_NodeClick(.Nodes(lngLoopCounter))
                    'Beep if wanted
                    If bolBeepOnFind Then Beep
                    Exit Sub
                
                'If all nodes have been searched
                ElseIf lngLoopCounter >= .Nodes.Count Then
                    'If we didn't start from the first node
                    If lngLastTree >= 1 And lngLastTree < .Nodes.Count Then
                        'Search from the start if wanted
                        If MsgBox("Finished searching through tree nodes, continue search from the start?", _
                            vbYesNo Or vbQuestion, "Restart Search?") = vbYes Then
                            lngLastTree = 1
                            GoTo Start
                        End If
                    
                    'If we have searched all nodes tell user
                    Else
                        lngLastTree = 1
                        Call MsgBox("Finished searching through tree nodes with no match.", _
                            vbOKOnly Or vbInformation, "Finished")
                        Exit Sub
                    End If
                End If
            'Onto next node
            Next lngLoopCounter
        End With
    
    'Code/Notes
    ElseIf SearchIn = Code Or SearchIn = Notes Then
        If SearchIn = Code Then
            lngLastCode = frmMain.rtfCode.Find(strLastFind, lngLastCode, , IIf(cpmLastCompareMethod = vbBinaryCompare, rtfMatchCase, 0))
            If lngLastCode < 0 Then
                Call MsgBox("Finished searching through the code with no match.", vbOKOnly Or vbInformation, "Finished")
                lngLastCode = 0
            Else
                'Beep if wanted
                If bolBeepOnFind Then Beep
                If lngLastCode > Len(frmMain.rtfCode.Text) Then
                    lngLastCode = 0
                Else
                    lngLastCode = lngLastCode + 1
                End If
            End If
        Else
            lngLastCode = frmMain.rtfNotes.Find(strLastFind, lngLastNotes)
            If lngLastNotes < 0 Then
                Call MsgBox("Finished searching through the notes with no match.", vbOKOnly Or vbInformation, "Finished")
                lngLastNotes = 0
            ElseIf lngLastNotes > Len(frmMain.rtfNotes.Text) Then
                lngLastNotes = 0
            Else
                lngLastNotes = lngLastNotes + 1
            End If
        End If

    'Bookmarks
    ElseIf SearchIn = Bookmarks Then
        'Increment the node to search from so we don't stay on the same one
        lngLastBookmarks = lngLastBookmarks + 1
        With frmMain.lvwBookmarks
            'If all nodes have been searched
            If lngLastBookmarks >= .ListItems.Count Then
                'Reset and tell user
                lngLastBookmarks = 1
                Call MsgBox("Finished searching through bookmark items with no match.", _
                    vbOKOnly Or vbInformation, "Finished")
                Exit Sub
            End If
            'Loop for all nodes from the last one found (+1)
            For lngLoopCounter = lngLastBookmarks To .ListItems.Count
                'Search for the string
                lngFound = InStr(1, .ListItems(lngLoopCounter).Text, strLastFind, cpmLastCompareMethod)
                'ADD WHITE SPACE CHECK HERE!
                'If the search string is found
                If lngFound > 0 Then
                    'Set this one as the last one and select the item
                    lngLastBookmarks = lngLoopCounter
                    .ListItems(.SelectedItem.Key).Selected = False
                    .ListItems(lngLoopCounter).Selected = True
                    'Beep if wanted
                    If bolBeepOnFind Then Beep
                    Exit Sub
                
                'If all nodes have been searched
                ElseIf lngLoopCounter >= .ListItems.Count Then
                    'If we didn't start from the first node
                    If lngLastBookmarks >= 1 And lngLastBookmarks < .ListItems.Count Then
                        'Search from the start if wanted
                        If MsgBox("Finished searching through bookmark items, continue search from the start?", _
                            vbYesNo Or vbQuestion, "Restart Search?") = vbYes Then
                            lngLastBookmarks = 1
                            GoTo Start
                        End If
                    
                    'If we have searched all nodes tell user
                    Else
                        lngLastBookmarks = 1
                        Call MsgBox("Finished searching through bookmark items with no match.", _
                            vbOKOnly Or vbInformation, "Finished")
                        Exit Sub
                    End If
                End If
            'Onto next node
            Next lngLoopCounter
        End With
    End If
End Sub

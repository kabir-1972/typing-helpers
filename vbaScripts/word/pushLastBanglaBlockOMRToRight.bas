Sub PushLastBanglaBlockOMRToRight()
    Dim para As Paragraph
    Dim doc As Document
    Dim i As Long
    Dim charRange As Range
    Dim lastBanglaPos As Long
    Dim targetX As Single, currentX As Single
    Dim insertPos As Range
    Dim spacesAdded As Long
    Dim bookmarkName As String
    
    On Error GoTo ErrorHandler
    
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    For Each para In doc.Paragraphs
        ' Skip if paragraph is too short
        If para.Range.Characters.Count < 2 Then GoTo nextPara
        
        ' Find last BanglaBlockOMR character by scanning backwards
        lastBanglaPos = -1
        For i = para.Range.Characters.Count - 1 To 1 Step -1
            On Error Resume Next
            Set charRange = para.Range.Characters(i)
            If Err.Number = 0 Then
                If charRange.Font.Name = "BanglaBlockOMR" Then
                    lastBanglaPos = i
                    Exit For
                End If
            End If
            On Error GoTo ErrorHandler
        Next i
        
        ' Skip if no BanglaBlockOMR found
        If lastBanglaPos = -1 Then GoTo nextPara
        
        ' Get position of last BanglaBlockOMR character FIRST
        Set charRange = para.Range.Characters(lastBanglaPos)
        charRange.Collapse wdCollapseEnd
        currentX = GetHorizontalPosition(charRange)
        
        If currentX <= 0 Then GoTo nextPara ' Can't measure position
        
        ' Now get the target position based on which column we're in
        targetX = GetColumnRightEdge(para, currentX)
        
        If targetX <= 0 Then GoTo nextPara ' Invalid target
        
        ' Debug info
        Debug.Print "Para " & para.Range.Start & ": Current=" & currentX & ", Target=" & targetX
        
        ' If already near right edge, skip
        If currentX >= targetX - 10 Then GoTo nextPara
        
        ' Mark the position with a bookmark for stable reference
        Set insertPos = para.Range.Characters(lastBanglaPos)
        bookmarkName = "TempBangla" & para.Range.Start
        
        ' Delete bookmark if it exists
        On Error Resume Next
        doc.Bookmarks(bookmarkName).Delete
        On Error GoTo ErrorHandler
        
        ' Create bookmark at the BanglaBlockOMR character
        doc.Bookmarks.Add Name:=bookmarkName, Range:=insertPos
        
        spacesAdded = 0
        Dim previousX As Single
        previousX = currentX
        
        Do While currentX < targetX - 5 And spacesAdded < 500
            ' Safety check - make sure bookmark still exists
            On Error Resume Next
            Set insertPos = doc.Bookmarks(bookmarkName).Range
            If Err.Number <> 0 Then
                Debug.Print "  Bookmark lost, stopping"
                Exit Do
            End If
            On Error GoTo ErrorHandler
            
            ' Insert space before the bookmarked character
            insertPos.Collapse wdCollapseStart
            insertPos.InsertBefore " "
            
            spacesAdded = spacesAdded + 1
            
            ' Remeasure using the bookmark
            On Error Resume Next
            Set charRange = doc.Bookmarks(bookmarkName).Range
            If Err.Number <> 0 Then Exit Do
            On Error GoTo ErrorHandler
            
            charRange.Collapse wdCollapseEnd
            previousX = currentX
            currentX = GetHorizontalPosition(charRange)
            
            ' If measurement failed, stop
            If currentX <= 0 Then Exit Do
            
            ' CRITICAL: If position moved backwards or moved to a new line, we've wrapped - stop immediately!
            If currentX < previousX - 10 Then
                Debug.Print "  Text wrapped to next line! Stopping and removing last space."
                ' Remove the space we just added
                insertPos.Collapse wdCollapseStart
                insertPos.MoveStart wdCharacter, -1
                insertPos.Delete
                spacesAdded = spacesAdded - 1
                Exit Do
            End If
            
            ' Every 10 spaces, check if we're making progress
            If spacesAdded Mod 10 = 0 Then
                Debug.Print "  Added " & spacesAdded & " spaces, now at " & currentX & " (target=" & targetX & ")"
            End If
        Loop
        
        ' Clean up bookmark
        On Error Resume Next
        doc.Bookmarks(bookmarkName).Delete
        On Error GoTo ErrorHandler
        
        Debug.Print "  Final: Added " & spacesAdded & " spaces"
        
nextPara:
    Next para
    
    Application.ScreenUpdating = True
    MsgBox "Done! Check Immediate Window (Ctrl+G) for debug info.", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error occurred: " & Err.Description & vbCrLf & "Error number: " & Err.Number, vbCritical
    Debug.Print "ERROR: " & Err.Description & " (Number: " & Err.Number & ")"
End Sub

Function GetHorizontalPosition(rng As Range) As Single
    On Error Resume Next
    GetHorizontalPosition = rng.Information(wdHorizontalPositionRelativeToPage)
    If Err.Number <> 0 Or GetHorizontalPosition = 0 Then
        GetHorizontalPosition = -1
    End If
    On Error GoTo 0
End Function

Function GetColumnRightEdge(para As Paragraph, currentX As Single) As Single
    Dim pgSetup As PageSetup
    Dim pgWidth As Single, leftM As Single, rightM As Single
    Dim numCols As Long, colSpacing As Single, colWidth As Single
    Dim rightIndent As Single
    Dim currentCol As Long
    
    On Error GoTo SafeExit
    
    Set pgSetup = para.Range.Sections(1).PageSetup
    
    With pgSetup
        pgWidth = .PageWidth
        leftM = .leftMargin
        rightM = .rightMargin
        numCols = .TextColumns.Count
        If numCols < 1 Then numCols = 1
        colSpacing = .TextColumns.Spacing
    End With
    
    ' Calculate column width
    If numCols > 1 Then
        colWidth = (pgWidth - leftM - rightM - colSpacing * (numCols - 1)) / numCols
    Else
        colWidth = pgWidth - leftM - rightM
    End If
    
    ' Account for paragraph indents
    rightIndent = para.Format.rightIndent
    
    ' Detect which column we're in by the current horizontal position
    If numCols > 1 And currentX > 0 Then
        ' Calculate which column based on horizontal position
        currentCol = Int((currentX - leftM) / (colWidth + colSpacing)) + 1
        If currentCol > numCols Then currentCol = numCols
        If currentCol < 1 Then currentCol = 1
        
        Debug.Print "  Detected column " & currentCol & " of " & numCols
        
        ' Calculate right edge of the detected column
        GetColumnRightEdge = leftM + (currentCol * colWidth) + ((currentCol - 1) * colSpacing) - rightIndent
    Else
        ' Single column or couldn't detect
        GetColumnRightEdge = leftM + colWidth - rightIndent
    End If
    
    Exit Function
    
SafeExit:
    GetColumnRightEdge = -1
End Function

'Works throughout the Doc automatically, no selection needed
Sub SendBracketedTextFarRight()
    Dim rngSearch As Range
    Dim rngFound As Range
    Dim foundRanges As Collection
    Dim para As Paragraph
    Dim r As Range
    Dim pgW As Single, leftMargin As Single, rightMargin As Single
    Dim colSpacing As Single, numCols As Long, colWidth As Single
    Dim usableWidth As Single
    
    Set foundRanges = New Collection
    Set rngSearch = ActiveDocument.Content
    
    ' === Step 1: Collect all bracketed ranges ===
    With rngSearch.Find
        .ClearFormatting
        .Text = "\[*\]"
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindStop
        
        Do While .Execute
            Set rngFound = rngSearch.Duplicate
            foundRanges.Add rngFound.Duplicate
            rngSearch.Start = rngFound.End
            rngSearch.End = ActiveDocument.Content.End
        Loop
    End With
    
    ' === Step 2: Process each match AFTER collecting ===
    For Each r In foundRanges
        Set para = r.Paragraphs(1)
        
        ' Page margins
        With para.Range.Sections(1).PageSetup
            pgW = .PageWidth
            leftMargin = .LeftMargin
            rightMargin = .RightMargin
        End With
        
        ' Column setup
        With para.Range.Sections(1).PageSetup.TextColumns
            numCols = .Count
            colSpacing = .Spacing
        End With
        If numCols < 1 Then numCols = 1
        
        usableWidth = pgW - leftMargin - rightMargin
        If usableWidth > 0 Then
            If numCols > 1 Then
                colWidth = (usableWidth - colSpacing * (numCols - 1)) / numCols
            Else
                colWidth = usableWidth
            End If
            
            ' Reset tabs for this paragraph
            para.Format.TabStops.ClearAll
            para.Format.TabStops.Add Position:=colWidth, Alignment:=wdAlignTabRight
            
            ' Insert tab before the bracketed text
            r.Collapse Direction:=wdCollapseStart
            r.InsertBefore vbTab
        End If
    Next r
End Sub

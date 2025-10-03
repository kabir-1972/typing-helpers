Sub SendSelectionFarRight()
    Dim sel As Range
    Dim para As Paragraph
    Dim selPortion As Range
    Dim pgW As Single, leftMargin As Single, rightMargin As Single
    Dim colSpacing As Single, numCols As Long, colWidth As Single
    Dim usableWidth As Single, leftPagePos As Single, leftRelative As Single
    Dim colIndex As Long, colStart As Single, colEnd As Single
    
    ' Require text selection
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select some text first.", vbInformation
        Exit Sub
    End If
    
    Set sel = Selection.Range.Duplicate
    If sel Is Nothing Then Exit Sub
    
    For Each para In sel.Paragraphs
        If Len(Trim(para.Range.Text)) <= 1 Then GoTo NextPara
        
        ' Page margins
        With para.Range.Sections(1).PageSetup
            pgW = .PageWidth
            leftMargin = .leftMargin
            rightMargin = .rightMargin
        End With
        
        ' Column setup
        With para.Range.Sections(1).PageSetup.TextColumns
            numCols = .Count
            colSpacing = .Spacing
        End With
        If numCols < 1 Then numCols = 1
        
        usableWidth = pgW - leftMargin - rightMargin
        If usableWidth <= 0 Then GoTo NextPara
        
        If numCols > 1 Then
            colWidth = (usableWidth - colSpacing * (numCols - 1)) / numCols
        Else
            colWidth = usableWidth
        End If
        
        ' Find left position of this paragraph (relative to page)
        On Error Resume Next
        leftPagePos = para.Range.Characters(1).Information(wdHorizontalPositionRelativeToPage)
        On Error GoTo 0
        
        leftRelative = leftPagePos - leftMargin
        If leftRelative < 0 Then leftRelative = 0
        
        ' Which column are we in?
        colIndex = Int(leftRelative / (colWidth + colSpacing)) + 1
        If colIndex < 1 Then colIndex = 1
        If colIndex > numCols Then colIndex = numCols
        
        ' Column start & end (absolute positions on page)
        colStart = leftMargin + (colIndex - 1) * (colWidth + colSpacing)
        colEnd = colStart + colWidth
        
        ' Reset tabs for this paragraph
        para.Format.TabStops.ClearAll
        
        ' Add a right tab relative to column (NOT the page)
        para.Format.TabStops.Add Position:=colWidth, Alignment:=wdAlignTabRight
        
        ' Portion of paragraph that is inside selection
        Set selPortion = para.Range.Duplicate
        If selPortion.Start < sel.Start Then selPortion.Start = sel.Start
        If selPortion.End > sel.End Then selPortion.End = sel.End
        If selPortion.Start >= selPortion.End Then GoTo NextPara
        
        ' Collapse to start and insert tab
        selPortion.Collapse Direction:=wdCollapseStart
        selPortion.InsertBefore vbTab
        
NextPara:
    Next para
End Sub
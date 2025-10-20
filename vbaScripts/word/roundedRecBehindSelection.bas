Sub PutRoundedRectangleBehindSelection()
    On Error GoTo ErrHandler
    Dim selRange As Range
    Dim shp As Shape
    Dim leftPos As Single, topPos As Single
    Dim shpWidth As Single, shpHeight As Single
    
    ' Fixed size in points (1 inch = 72 points)
    shpWidth = 0.57 * 72
    shpHeight = 0.2 * 72
    
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select some text first.", vbExclamation
        Exit Sub
    End If
    
    Set selRange = Selection.Range.Duplicate
    
    ' Get top-left coordinate of selection on the page (points)
    leftPos = selRange.Information(wdHorizontalPositionRelativeToPage)
    topPos = selRange.Information(wdVerticalPositionRelativeToPage)
    
    ' Insert rounded rectangle anchored to selection
    Set shp = ActiveDocument.Shapes.AddShape( _
                Type:=msoShapeRoundedRectangle, _
                Left:=leftPos, Top:=topPos, _
                Width:=shpWidth, Height:=shpHeight, _
                Anchor:=selRange)
    
    With shp
        .WrapFormat.Type = wdWrapBehind
        .ZOrder msoSendBehindText
        
        ' Style
        .Fill.Visible = msoFalse
        .Line.Visible = msoTrue
        .Line.Weight = 1
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        
        ' Rounded corners
        On Error Resume Next
        .Adjustments.Item(1) = 0.5
        On Error GoTo ErrHandler
    End With

    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub
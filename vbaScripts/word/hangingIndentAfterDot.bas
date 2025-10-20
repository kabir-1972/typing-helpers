Sub HangingIndentAndTabAfterDot()
    Dim sel As Range
    Dim para As Paragraph
    Dim pos As Long
    
    ' Work with the current selection
    Set sel = Selection.Range
    
    ' Apply hanging indent = 0.31"
    For Each para In sel.Paragraphs
        With para.Format
            .LeftIndent = InchesToPoints(0.31)
            .FirstLineIndent = -InchesToPoints(0.31) ' pulls the first line back to margin
        End With
    Next para
    
    ' Find the first dot and add a tab after it
    pos = InStr(sel.Text, ". ")
    If pos > 0 Then
        sel.Characters(pos + 1).InsertAfter vbTab
    End If
End Sub
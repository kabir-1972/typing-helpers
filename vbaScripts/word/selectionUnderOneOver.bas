Sub PutSelectionUnderOneOver()
    If Selection.Type = wdNoSelection Or Trim(Selection.Range.Text) = "" Then
        MsgBox "Please select some text first.", vbExclamation
        Exit Sub
    End If
    
    Dim selText As String
    selText = Replace(Trim(Selection.Range.Text), vbCr, " ")
    
    ' Replace selection with "1/(selText)"
    Selection.Text = "1/(" & selText & ")"
    
    ' Convert THAT text into an equation
    Dim rng As Range
    Set rng = Selection.Range
    rng.OMaths.Add rng
    rng.OMaths(1).BuildUp
End Sub
Sub ToggleFont()
    Dim rng As Range
    
    ' If nothing selected, use current insertion point
    If Selection.Type = wdNoSelection Then
        Set rng = Selection.Range
    Else
        Set rng = Selection.Range
    End If
    
    ' Loop through words to handle mixed fonts
    Dim i As Long
    For i = 1 To rng.Words.Count
        With rng.Words(i).Font
            Select Case .Name
                Case "Times New Roman"
                    .Name = "SutonnyMJ"
                Case "SutonnyMJ"
                    .Name = "Times New Roman"
                Case Else
                    .Name = "Times New Roman"
            End Select
        End With
    Next i
End Sub
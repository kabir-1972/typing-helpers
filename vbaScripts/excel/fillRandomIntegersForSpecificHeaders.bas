Sub FillRandomIntegersForSpecificHeaders()
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim col As Long
    Dim header As String
    Dim maxVal As Long
    Dim r As Long
    
    Set ws = ActiveSheet
    
    ' Find the last used column in row 3
    lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column
    
    Randomize ' Seed the random generator
    
    ' Loop through all columns with data in row 3
    For col = 1 To lastCol
        header = Trim(ws.Cells(3, col).Value)
        
        ' Check if the header matches any of the target names
        If header = "Creative" Or header = "MCQ" Or header = "Assignment" Or header = "Hygiene" Then
            ' Get the max value from row 2
            If IsNumeric(ws.Cells(2, col).Value) And ws.Cells(2, col).Value >= 0 Then
                maxVal = ws.Cells(2, col).Value
                
                ' Fill rows 4 to 43 with random integers
                For r = 4 To 43
                    ws.Cells(r, col).Value = Int((maxVal + 1) * Rnd)
                Next r
            End If
        End If
    Next col
End Sub


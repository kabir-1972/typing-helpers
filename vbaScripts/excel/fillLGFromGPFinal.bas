Sub FillLGFromGPFinal()
    Dim ws As Worksheet
    Dim selRange As Range
    Dim rowNum As Long
    Dim GPval As Double
    
    Set ws = ActiveSheet
    Set selRange = Selection
    
    ' Make sure a single column is selected
    If selRange.Columns.Count <> 1 Then
        MsgBox "Please select a single LG column."
        Exit Sub
    End If
    
    ' Loop through each row in the selected column
    For Each cell In selRange
        rowNum = cell.Row
        
        ' Read GP value from the previous column
        GPval = ws.Cells(rowNum, cell.Column - 1).Value
        
        ' Check if GP is numeric
        If IsNumeric(GPval) Then
            Select Case GPval
                Case 5
                    cell.Value = "A+"
                Case 4 To 4.99
                    cell.Value = "A"
                Case 3.5 To 3.99
                    cell.Value = "A-"
                Case 3 To 3.49
                    cell.Value = "B"
                Case 2 To 2.99
                    cell.Value = "C"
                Case 1 To 1.99
                    cell.Value = "D"
                Case 0
                    cell.Value = "F"
                Case Else
                    cell.Value = ""
            End Select
        Else
            cell.Value = ""  ' blank if GP not numeric
        End If
    Next cell
End Sub


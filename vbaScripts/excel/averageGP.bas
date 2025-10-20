Sub AverageGP()
    Dim ws As Worksheet
    Dim lastCol As Long, col As Variant
    Dim rowNum As Long
    Dim totalGP As Double, countGP As Long
    Dim selRange As Range
    Dim headerRow As Long
    Dim GPCols As Collection
    Dim cell As Range
    
    Set ws = ActiveSheet
    Set selRange = Selection
    
    ' Check selection
    If selRange.Columns.Count <> 1 Then
        MsgBox "Please select a single column to fill the GP average."
        Exit Sub
    End If
    
    ' Set the row that contains your GP headers
    headerRow = 2   ' change if your "GP" labels are in a different row
    
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    
    ' Detect all GP columns
    Set GPCols = New Collection
    For col = 1 To lastCol
        If UCase(Trim(ws.Cells(headerRow, col).Value)) = "GP" Then
            GPCols.Add col
        End If
    Next col
    
    If GPCols.Count = 0 Then
        MsgBox "No GP columns found in row " & headerRow
        Exit Sub
    End If
    
    ' Loop through each row in the selected column
    For Each cell In selRange
        rowNum = cell.Row
        totalGP = 0
        countGP = 0
        
        ' Sum all GP values in this row
        For Each col In GPCols
            If IsNumeric(ws.Cells(rowNum, col).Value) Then
                totalGP = totalGP + ws.Cells(rowNum, col).Value
                countGP = countGP + 1
            End If
        Next col
        
        ' Calculate average and write into the selected column cell
        If countGP > 0 Then
            cell.Value = Round(totalGP / countGP, 2)  ' rounded to 2 decimals
        Else
            cell.Value = ""  ' no GP values found
        End If
    Next cell
End Sub


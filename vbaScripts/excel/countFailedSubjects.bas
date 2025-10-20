Sub CountFailedSubjects()
    Dim ws As Worksheet
    Dim lastCol As Long, col As Variant
    Dim rowNum As Long
    Dim Fcount As Long
    Dim selRange As Range
    Dim headerRow As Long
    Dim LGCols As Collection
    Dim cell As Range
    
    Set ws = ActiveSheet
    Set selRange = Selection
    
    ' Check selection
    If selRange.Columns.Count <> 1 Then
        MsgBox "Please select a single column to fill the F count."
        Exit Sub
    End If
    
    ' Set the row that contains your LG headers
    headerRow = 2   ' change this if your "LG" labels are in a different row
    
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    
    ' Detect all LG columns
    Set LGCols = New Collection
    For col = 1 To lastCol
        If UCase(Trim(ws.Cells(headerRow, col).Value)) = "LG" Then
            LGCols.Add col
        End If
    Next col
    
    If LGCols.Count = 0 Then
        MsgBox "No LG columns found in row " & headerRow
        Exit Sub
    End If
    
    ' Loop through each row in the selected column
    For Each cell In selRange
        rowNum = cell.Row
        Fcount = 0
        
        ' Check all LG columns for "F" in this row
        For Each col In LGCols
            If UCase(Trim(ws.Cells(rowNum, col).Value)) = "F" Then
                Fcount = Fcount + 1
            End If
        Next col
        
        ' Write the count into the selected column cell
        cell.Value = Fcount
    Next cell
End Sub


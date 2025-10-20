Sub FillMeritOrderWithCardinals()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim totalCol As Long, meritCol As Long
    Dim i As Long, j As Long
    Dim studentMarks() As Variant
    Dim studentRows() As Long
    Dim n As Long
    Dim tempMarks As Double
    Dim tempRow As Long
    
    Set ws = ActiveSheet
    
    ' Find last row and last column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Find columns
    totalCol = 0
    meritCol = 0
    For i = 1 To lastCol
        If Trim(UCase(ws.Cells(1, i).Value)) = "TOTAL OBTAINED MARKS" Then totalCol = i
        If Trim(UCase(ws.Cells(1, i).Value)) = "MERIT ORDER" Then meritCol = i
    Next i
    
    If totalCol = 0 Or meritCol = 0 Then
        MsgBox "Could not find 'Total Obtained Marks' or 'Merit Order' column."
        Exit Sub
    End If
    
    n = lastRow - 1 ' assuming row 1 is header
    
    ' Read student marks into array
    ReDim studentMarks(1 To n)
    ReDim studentRows(1 To n)
    
    For i = 2 To lastRow
        studentMarks(i - 1) = ws.Cells(i, totalCol).Value
        studentRows(i - 1) = i
    Next i
    
    ' Sort the array in descending order
    For i = 1 To n - 1
        For j = i + 1 To n
            If studentMarks(i) < studentMarks(j) Then
                ' Swap marks
                tempMarks = studentMarks(i)
                studentMarks(i) = studentMarks(j)
                studentMarks(j) = tempMarks
                ' Swap row numbers
                tempRow = studentRows(i)
                studentRows(i) = studentRows(j)
                studentRows(j) = tempRow
            End If
        Next j
    Next i
    
    ' Fill Merit Order with cardinal numbers
    For i = 1 To n
        ws.Cells(studentRows(i), meritCol).Value = GetOrdinal(i)
    Next i
    
    MsgBox "Merit order filled successfully with cardinal numbers."
End Sub

' Function to convert number to ordinal (1 ? 1st, 2 ? 2nd, etc.)
Function GetOrdinal(ByVal num As Long) As String
    Select Case num Mod 100
        Case 11, 12, 13
            GetOrdinal = num & "th"
        Case Else
            Select Case num Mod 10
                Case 1: GetOrdinal = num & "st"
                Case 2: GetOrdinal = num & "nd"
                Case 3: GetOrdinal = num & "rd"
                Case Else: GetOrdinal = num & "th"
            End Select
    End Select
End Function


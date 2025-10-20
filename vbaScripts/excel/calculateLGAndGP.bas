Sub CalculateLGandGP()
    Dim ws As Worksheet
    Dim lastCol As Long, col As Long, r As Long
    Dim colHeader As String, prevHeader As String, prevHeaderForGP As String
    Dim totalVal As Double
    Dim totalCol As Long
    
    Set ws = ActiveSheet
    lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column
    
    ' Loop through all columns to find LG or GP columns
    For col = 3 To lastCol
        colHeader = Trim(ws.Cells(2, col).Value)
        prevHeader = Trim(ws.Cells(3, col - 1).Value)
        
        If col > 3 Then
            prevHeaderForGP = Trim(ws.Cells(3, col - 2).Value)
        Else
            prevHeaderForGP = ""
        End If
        
        ' Only proceed if row 2 = LG or GP and find the Total column
        If colHeader = "LG" And prevHeader = "Total" Then
            totalCol = col - 1
        ElseIf colHeader = "GP" And prevHeaderForGP = "Total" Then
            totalCol = col - 2
        Else
            totalCol = 0
        End If
        
        If totalCol > 0 Then
            For r = 4 To 43
                If IsNumeric(ws.Cells(r, totalCol).Value) And ws.Cells(r, totalCol).Value <> "" Then
                    totalVal = ws.Cells(r, totalCol).Value
                    
                    Select Case totalVal
                        Case Is >= 80
                            If colHeader = "LG" Then ws.Cells(r, col).Value = "A+"
                            If colHeader = "GP" Then ws.Cells(r, col).Value = 5#
                        Case Is >= 70
                            If colHeader = "LG" Then ws.Cells(r, col).Value = "A"
                            If colHeader = "GP" Then ws.Cells(r, col).Value = 4#
                        Case Is >= 60
                            If colHeader = "LG" Then ws.Cells(r, col).Value = "A-"
                            If colHeader = "GP" Then ws.Cells(r, col).Value = 3.5
                        Case Is >= 50
                            If colHeader = "LG" Then ws.Cells(r, col).Value = "B"
                            If colHeader = "GP" Then ws.Cells(r, col).Value = 3#
                        Case Is >= 40
                            If colHeader = "LG" Then ws.Cells(r, col).Value = "C"
                            If colHeader = "GP" Then ws.Cells(r, col).Value = 2#
                        Case Is >= 33
                            If colHeader = "LG" Then ws.Cells(r, col).Value = "D"
                            If colHeader = "GP" Then ws.Cells(r, col).Value = 1#
                        Case Else
                            If colHeader = "LG" Then ws.Cells(r, col).Value = "F"
                            If colHeader = "GP" Then ws.Cells(r, col).Value = 0#
                    End Select
                Else
                    ws.Cells(r, col).ClearContents
                End If
            Next r
        End If
    Next col
End Sub

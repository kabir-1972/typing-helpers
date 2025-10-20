Option Explicit

' --- Helper: check if char is punctuation or space ---
Function IsPunctuationOrSpace(ch As String) As Boolean
    Select Case ch
        Case " ", ",", ".", ":", ";", "!", "?", "-", _
             "(", ")", "[", "]", "{", "}", "'", """"
            IsPunctuationOrSpace = True
        Case Else
            IsPunctuationOrSpace = False
    End Select
End Function


Sub BlockBijoyConvert_MultiParagraph()
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    Dim selRange As Range
    Set selRange = Selection.Range

    Dim para As Paragraph
    Dim ch As Range
    Dim buffer As String
    Dim bufferRange As Range
    Dim i As Long
    
    buffer = ""
    Set bufferRange = Nothing

    For Each para In selRange.Paragraphs
        Dim chars As Characters
        Set chars = para.Range.Characters
        
        For i = 1 To chars.Count
            Set ch = chars(i)
            
            ' Non-TNR OR punctuation/space ? add to buffer
            If (InStr(1, ch.Font.Name, "Times New Roman", vbTextCompare) = 0) _
               Or IsPunctuationOrSpace(ch.Text) Then
               
               If buffer = "" Then
                   ' Start new buffer range
                   Set bufferRange = ch.Duplicate
               Else
                   bufferRange.End = ch.End
               End If
               
               buffer = buffer & ch.Text
               
            Else
                ' --- Hit TNR ? flush buffer if not empty ---
                If buffer <> "" Then
                    Call FlushBuffer(bufferRange, buffer, http)
                    buffer = ""
                    Set bufferRange = Nothing
                End If
            End If
        Next i
    Next para
    
    ' Flush leftover buffer
    If buffer <> "" Then
        Call FlushBuffer(bufferRange, buffer, http)
    End If
    
    MsgBox "Done converting runs across multiple paragraphs!", vbInformation
End Sub


Private Sub FlushBuffer(ByVal r As Range, ByVal txt As String, ByVal http As Object)
    Dim result As String
    Dim origBold As Long, origItalic As Long, origUnderline As Long
    
    ' Save formatting
    origBold = r.Font.Bold
    origItalic = r.Font.Italic
    origUnderline = r.Font.Underline
    
    ' --- Send to server ---
    http.Open "POST", "http://localhost:1337/", False
    http.setRequestHeader "Content-Type", "text/plain; charset=utf-8"
    http.send txt
    result = http.responseText
    
    ' --- Apply replacement rules ---
    result = Replace(result, ChrW(&HAD), "")
    result = Replace(result, "h" & ChrW(&H2021) & ChrW(&H9BC), "‡q")
    result = Replace(result, "Ww" & ChrW(&H9BC), "wo")
    result = Replace(result, "W" & ChrW(&H9BC), "o")
    result = Replace(result, "h" & ChrW(&H9BC), "q")
    
    ' --- Replace text ---
    r.Text = result
    r.Font.Name = "SutonnyMJ"
    
    ' Restore formatting
    r.Font.Bold = origBold
    r.Font.Italic = origItalic
    r.Font.Underline = origUnderline
    
    Debug.Print "Flushed buffer: " & Left(txt, 40) & "..."
End Sub



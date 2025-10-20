Option Explicit

' --- Helper: check if char is punctuation (excluding space) ---
Function IsPunctuation(ch As String) As Boolean
    Select Case ch
        Case ",", ".", ":", ";", "!", "?", "-", _
             "(", ")", "[", "]", "{", "}", "'", """", "?"
            IsPunctuation = True
        Case Else
            IsPunctuation = False
    End Select
End Function

' --- Helper: check if string has Bangla characters (including danda) ---
Function HasBanglaCharacters(txt As String) As Boolean
    Dim i As Long, code As Long
    For i = 1 To Len(txt)
        code = AscW(Mid$(txt, i, 1))
        If (code >= &H980 And code <= &H9FF) Or code = &H964 Then
            HasBanglaCharacters = True
            Exit Function
        End If
    Next i
    HasBanglaCharacters = False
End Function

' --- Main macro ---
Sub ConvertBanglaPreserveEnglish()
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    Dim selRange As Range
    Set selRange = Selection.Range

    Dim para As Paragraph
    Dim i As Long
    Dim chars As Characters
    Dim ch As Range
    Dim buffer As String
    Dim bufferStart As Long, bufferEnd As Long
    Dim insideBuffer As Boolean

    Dim segments As Collection
    Dim segData As Variant
    Dim j As Long

    For Each para In selRange.Paragraphs
        Set chars = para.Range.Characters
        buffer = ""
        insideBuffer = False
        Set segments = New Collection

        ' --- Step 1: Collect all Bangla segments in this paragraph ---
        For i = 1 To chars.count
            Set ch = chars(i)

            If HasBanglaCharacters(ch.Text) Or IsPunctuation(ch.Text) Then
                If Not insideBuffer Then
                    bufferStart = ch.Start
                    buffer = ch.Text
                    insideBuffer = True
                Else
                    buffer = buffer & ch.Text
                End If
                bufferEnd = ch.End

            ElseIf ch.Text = " " Then
                ' Include spaces in buffer if already inside
                If insideBuffer Then
                    buffer = buffer & ch.Text
                    bufferEnd = ch.End
                End If

            Else
                ' Flush buffer to segments collection
                If insideBuffer Then
                    segments.Add Array(bufferStart, bufferEnd, buffer)
                    buffer = ""
                    insideBuffer = False
                End If
            End If
        Next i

        ' Flush remaining buffer at end of paragraph
        If insideBuffer Then
            segments.Add Array(bufferStart, bufferEnd, buffer)
            buffer = ""
            insideBuffer = False
        End If

        ' --- Step 2: Replace segments in reverse order ---
        For j = segments.count To 1 Step -1
            segData = segments(j)
            ' --- Force segData(2) to String to avoid ByRef mismatch ---
            Call FlushBufferWord(selRange.Document.Range(segData(0), segData(1)), CStr(segData(2)), http)
        Next j
    Next para

    MsgBox "Bangla text converted. English text preserved.", vbInformation
End Sub

' --- Flush a single buffer to server and replace text ---
Private Sub FlushBufferWord(r As Range, txt As String, http As Object)
    Dim result As String
    Dim origBold As Long, origItalic As Long, origUnderline As Long
    Dim trailingSpaces As String
    Dim lastChar As Long
    
    ' Count trailing spaces
    trailingSpaces = ""
    For lastChar = Len(txt) To 1 Step -1
        If Mid$(txt, lastChar, 1) = " " Then
            trailingSpaces = " " & trailingSpaces
        Else
            Exit For
        End If
    Next lastChar
    
    ' Remove trailing spaces from buffer before sending
    If trailingSpaces <> "" Then
        txt = Left$(txt, Len(txt) - Len(trailingSpaces))
    End If

    ' Save formatting
    origBold = r.Font.Bold
    origItalic = r.Font.Italic
    origUnderline = r.Font.Underline

    ' Only send if there is Bangla
    If Not HasBanglaCharacters(txt) Then Exit Sub

    ' Send to server
    http.Open "POST", "http://localhost:1337/", False
    http.setRequestHeader "Content-Type", "text/plain; charset=utf-8"
    http.send txt
    result = http.responseText

    ' Custom replacements
    result = Replace(result, ChrW(&HAD), "")
    result = Replace(result, "h" & ChrW(&H2021) & ChrW(&H9BC), "‡q")
    result = Replace(result, "Ww" & ChrW(&H9BC), "wo")
    result = Replace(result, "W" & ChrW(&H9BC), "o")
    result = Replace(result, "h" & ChrW(&H9BC), "q")

    ' Append trailing spaces back
    result = result & trailingSpaces

    ' Replace text in Word
    r.Text = result
    r.Font.Name = "SutonnyMJ"

    ' Restore formatting
    r.Font.Bold = origBold
    r.Font.Italic = origItalic
    r.Font.Underline = origUnderline
End Sub






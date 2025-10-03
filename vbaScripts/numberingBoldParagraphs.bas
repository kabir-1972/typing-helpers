Sub NumberBoldParagraphs()
    Dim para As Paragraph
    Dim firstChar As Range
    Dim firstWord As String
    Dim listTemplate As listTemplate
    
    ' Create a new list template for continuous numbering
    Set listTemplate = ActiveDocument.ListTemplates.Add(OutlineNumbered:=False)
    With listTemplate.ListLevels(1)
        .NumberFormat = "%1."
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0)      ' Number at left margin
        .TextPosition = InchesToPoints(0.31)    ' Hanging indent
        .TabPosition = InchesToPoints(0.31)
        .ResetOnHigher = 0
        .LinkedStyle = ""
    End With
    
    ' Loop through all paragraphs
    For Each para In ActiveDocument.Paragraphs
        ' Skip empty paragraphs
        If Len(Trim(para.Range.Text)) > 0 Then
            Set firstChar = para.Range.Characters(1)
            firstWord = Trim(para.Range.Words(1).Text)
            firstWord = Replace(firstWord, vbCr, "")
            firstWord = Replace(firstWord, " ", "")
            
            ' Apply numbering if first char is bold AND first word is not "Ans"
            If firstChar.Font.Bold = True And StrComp(firstWord, "Ans", vbTextCompare) <> 0 Then
                para.Range.ListFormat.ApplyListTemplate _
                    listTemplate:=listTemplate, _
                    ContinuePreviousList:=True, _
                    ApplyTo:=wdListApplyToWholeList
            End If
        End If
    Next para
End Sub
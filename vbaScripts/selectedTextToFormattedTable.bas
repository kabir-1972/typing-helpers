Sub ConvertSelectionToFormattedTable()
    Dim tbl As Table
    
    ' Make sure something is selected
    If Selection.Range.Text = "" Then
        MsgBox "Please select some text first."
        Exit Sub
    End If
    
    ' Convert the selection to a table (split by tabs by default)
    Set tbl = Selection.ConvertToTable( _
        Separator:=wdSeparateByTabs, _
        AutoFitBehavior:=wdAutoFitContent)
    
    ' Show table borders
    tbl.Borders.Enable = True
    
    ' Center the table on the page
    tbl.Rows.Alignment = wdAlignRowCenter
    
    ' Apply font formatting (bold, size 12) to the whole table
    tbl.Range.Font.Size = 12
    tbl.Range.Font.Bold = True
    
    ' Make sure table fits to contents, not page width
    tbl.AutoFitBehavior (wdAutoFitContent)
End Sub
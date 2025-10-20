Sub SendBanglaTextUTF8()
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    Dim selText As String
    selText = Selection.Text
    
    ' Prepare UTF-8 bytes using ADODB.Stream
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' text
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText selText
    stream.Position = 0
    stream.Type = 1 ' binary
    Dim byteArr() As Byte
    byteArr = stream.Read
    stream.Close
    
    http.Open "POST", "http://localhost:1337/", False
    http.setRequestHeader "Content-Type", "text/plain; charset=utf-8"
    
    ' Use a variant to wrap byte array
    Dim sendVar As Variant
    sendVar = byteArr
    http.send sendVar
    
    ' Get response
    Dim result As String
    result = http.responseText
    
    FindText = "h" & ChrW(&H9BC)
    findText2 = "h" & ChrW(&H2021) & ChrW(&H9BC)
    
    findText2r = "W" & ChrW(&H9BC)
    findText2r2 = "W" & "w" & ChrW(&H9BC)
    
    
    findText2dr = "W" & "w" & ChrW(&H9BC)
    
    result = Replace(result, findText2, "‡q")
    result = Replace(result, findText2r2, "wo")
    result = Replace(result, FindText, "q")
    result = Replace(result, findText2r, "o")
    Debug.Print result
    Selection.Text = result
    Selection.Font.Name = "SutonnyMJ"
End Sub

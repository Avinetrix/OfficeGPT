Attribute VB_Name = "GPT4omini"
Sub GPT35Turbo()
    Dim selectedText As String
    Dim response As String
    Dim http As Object
    Dim url As String
    Dim apiKey As String
    Dim jsonData As String
    
    ' Get the selected text
    selectedText = Selection.Text
    
    ' Check if text is selected
    If Len(Trim(selectedText)) = 0 Then
        MsgBox "Please select some text first.", vbExclamation
        Exit Sub
    End If
    
    ' Set your API key here
    apiKey = "INSERTAPIKEYHERE"
    
    ' Set the API endpoint URL
    url = "https://api.openai.com/v1/chat/completions"
    
    ' Prepare the JSON data
    jsonData = "{""model"": ""gpt-4o-mini"", ""messages"": [{""role"": ""user"", ""content"": " & jsonString(selectedText) & "}]}"
    
    ' Create HTTP object
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' Open connection
    http.Open "POST", url, False
    
    ' Set headers
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    
    ' Send request
    On Error Resume Next
    http.send jsonData
    
    ' Check for errors
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Check the response status
    If http.Status <> 200 Then
        MsgBox "Error: " & http.Status & " " & http.StatusText & vbNewLine & http.responseText, vbCritical
        Exit Sub
    End If
    
    ' Debug: Print full response
    Debug.Print "Full API Response:"
    Debug.Print http.responseText
    
    ' Get the response
    response = ParseJsonResponse(http.responseText)
    
    ' Insert response into document
    Selection.EndOf Unit:=wdLine
    Selection.TypeParagraph
    Selection.TypeText Text:=response
    
    ' Clean up
    Set http = Nothing
End Sub

Function ParseJsonResponse(jsonString As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    startPos = InStr(jsonString, """content"": """)
    If startPos > 0 Then
        startPos = startPos + 12 ' Length of """content"": """
        endPos = InStr(startPos, jsonString, """")
        If endPos > 0 Then
            ParseJsonResponse = Mid(jsonString, startPos, endPos - startPos)
            ParseJsonResponse = Replace(ParseJsonResponse, "\""", """")
            ParseJsonResponse = Replace(ParseJsonResponse, "\n", vbNewLine)
        Else
            ParseJsonResponse = "Error: Unable to find end of content"
        End If
    Else
        ParseJsonResponse = "Error: Unable to find content in response"
    End If
End Function

Function jsonString(str As String) As String
    Dim result As String
    result = Replace(str, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    jsonString = """" & result & """"
End Function

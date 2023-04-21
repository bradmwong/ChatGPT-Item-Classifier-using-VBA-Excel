Attribute VB_Name = "ExcelCellClassifier"
' Reference: Microsoft XML, v6.
' Reference: Microsoft Scripting Runtime

Sub ChatWithGPT()

    Dim apiKey As String
    Dim apiUrl As String
    Dim inputValue As String
    Dim response As String
        
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
       
    ' Set your API key and API URL
    apiKey = "YourAPIKeyHere"
    apiUrl = "https://api.openai.com/v1/completions"
    
    ' Get the user's message
    inputValue = ws.Range("A2").value
    
    ' Send a request to the ChatGPT API
    response = SendRequest(apiUrl, apiKey, inputValue)
    
    ' Parse the JSON response and extract the response message
    response = ParseResponse(response)
    
    ' Display the response message in a cell
    ws.Range("B2").value = response
    
End Sub

Function SendRequest(apiUrl As String, apiKey As String, inputValue)
' As String) As String

    ' Create an HTTP request with the input value and API key
    Dim http As New MSXML2.XMLHTTP60
    Dim response As String
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    
    ' Set the request body with the user's message
    Dim prompt As String
    Dim request As String
    prompt = "I have a(n) 'snake fruit'. Is it a fruit, vegetable, or other? Provide only 1 word as an answer, without punctuation, all lowercase. Do not provide answers other than 'fruit', 'vegetable', or 'other'."
    request = "{""model"": ""text-davinci-003"", ""prompt"": """ & prompt & """, ""max_tokens"": 20, ""temperature"": 0}"
    Debug.Print request
    
    ' Send the request and get the response
    http.send request
    response = http.responseText
    
    ' Return the response
    SendRequest = response
    
End Function

Function ParseResponse(response As String) As String

    ' Parse the JSON response and extract the response message
    Dim json As Object
    Set json = JsonConverter.ParseJson(response)
    
    Debug.Print response
'    Debug.Print TypeName(json("choices"))
'    Debug.Print json("choices").Count
'
'    Dim choice As Object
'    Set choice = json("choices")(1)
'    Debug.Print TypeName(choice)
'    Debug.Print choice("text")
    
    ParseResponse = json("choices")(1)("text")
    ParseResponse = LCase(ParseResponse)
    ParseResponse = Trim(ParseResponse)
    ParseResponse = Replace(ParseResponse, vbNewLine, "")
    Debug.Print ParseResponse
    
End Function


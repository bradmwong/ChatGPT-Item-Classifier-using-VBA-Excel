Attribute VB_Name = "ExcelCellClassifier"
' Reference: Microsoft XML, v6.
' reference: Microsoft Scripting Runtime

Sub ChatWithGPT()
    Dim apiKey As String
    Dim apiUrl As String
    Dim message As String
    Dim response As String
    
    ' Set your API key and API URL
    apiKey = "YourAPIKeyHere"
    apiUrl = "https://api.openai.com/v1/completions"
    
    ' Get the user's message
    message = InputBox("Enter your message:")
'    message = "Say this is a test."
    
    ' Send a request to the ChatGPT API
    response = SendRequest(apiUrl, apiKey, message)
    
    ' Parse the JSON response and extract the response message
    response = ParseResponse(response)
    
    ' Display the response message in a cell
    Range("A1").Value = response
End Sub

Function SendRequest(apiUrl As String, apiKey As String, message As String) As String
    ' Create an HTTP request with the user's message and API key
    Dim http As New MSXML2.XMLHTTP60
    Dim response As String
    
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    
    ' Set the request body with the user's message
    Dim request As String
    request = "{""model"": ""text-davinci-003"", ""prompt"": """ & message & """, ""max_tokens"": 20, ""temperature"": 0}"
    
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
    Debug.Print ParseResponse
    
End Function


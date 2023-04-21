Attribute VB_Name = "ExcelCellClassifier"
' Reference: Microsoft XML, v6.
' Reference: Microsoft Scripting Runtime

Option Compare Binary
Option Explicit


Sub Main()

    Dim apiKey As String
    apiKey = "YourAPIKeyHere"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Dim inputValue As String
    Dim outputOptions(1 To 2) As String
    outputOptions(1) = "fruit"
    outputOptions(2) = "vegetable"
    
    inputValue = ws.Range("A2").value

'    Call testing(outputOptions())
    
    ws.Range("B2").value = ClassifyTypeWithGPT(inputValue, outputOptions(), apiKey)
    
End Sub


Public Function ClassifyTypeWithGPT(inputValue, outputsArray() As String, apiKey As String) As String

    Dim apiUrl As String
    Dim response As String
    Dim prompt As String
    
    
    prompt = GenerateGPTPromptString(inputValue, outputsArray())
    
    ' Set your API key and API URL
    apiUrl = "https://api.openai.com/v1/completions"
       
    ' Send a request to the ChatGPT API
    response = SendRequest(apiUrl, apiKey, prompt)
    
    ' Parse the JSON response and extract the response message
    response = ParseResponse(response)
    
   
    ClassifyTypeWithGPT = response

End Function


' Function to generate a GPT prompt string
' Input: inputValue - the value to be identified
'        outputsArray() - an array of possible types the value could be
' Output: a string prompt that asks GPT to identify the type of the input value from the list of possible options
Private Function GenerateGPTPromptString(inputValue, outputsArray() As String) As String
    
    ' Join the elements of the outputsArray with a comma and add "or other" to the end
    Dim outputsTextPrompt As String
    outputsTextPrompt = Join(outputsArray, ", ")
    outputsTextPrompt = outputsTextPrompt & ", or other"
    
    ' Return the prompt string
    GenerateGPTPromptString = "I have a(n) '" & inputValue & "'. " & _
        "Is it a(n) " & outputsTextPrompt & "? " & _
        "Provide only 1 word as an answer, without punctuation, all lowercase. " & _
        "Do not provide answers other than " & outputsTextPrompt & "."
    
End Function


Private Function SendRequest(apiUrl As String, apiKey As String, prompt As String) As String

    ' Create an HTTP request with the input value and API key
    Dim http As New MSXML2.XMLHTTP60
    Dim response As String
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    
    ' Set the request body with the user's message
    Dim request As String
    request = "{""model"": ""text-davinci-003"", ""prompt"": """ & prompt & """, ""max_tokens"": 20, ""temperature"": 0}"
    Debug.Print request
    
    ' Send the request and get the response
    http.send request
    response = http.responseText
    
    ' Return the response
    SendRequest = response
    
End Function

Private Function ParseResponse(response As String) As String

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





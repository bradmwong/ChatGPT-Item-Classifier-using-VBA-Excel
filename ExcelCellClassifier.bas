Attribute VB_Name = "ExcelCellClassifier"
' Reference: Microsoft XML, v6.
' Reference: Microsoft Scripting Runtime

Option Compare Binary
Option Explicit


Sub Main()

    Const modelName As String = "gpt-3.5-turbo"



    Dim apikey As String
    apikey = "YourAPIKeyHere"
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Dim inputValue As String
    Dim outputOptions(1 To 2) As String
    outputOptions(1) = "fruit"
    outputOptions(2) = "vegetable"
    
    inputValue = ws.Range("A2").value

    ws.Range("B2").value = ClassifyTypeWithGPT(inputValue, outputOptions(), apikey, modelName)

    
    
End Sub


Public Function ClassifyTypeWithGPT(inputValue As String, outputsArray() As String, apikey As String, Optional modelName As String = "gpt-3.5-turbo") As String

    
    Dim prompt As String
    Dim response As String
    
    ' Generate GPT text prompt
    prompt = GenerateGPTPromptString(inputValue, outputsArray())

    ' Send a request to the ChatGPT API
    response = SendGPTRequest(prompt, apikey, modelName)
    
    ' Parse the JSON response and extract the response message
    response = ParseResponse(response, modelName)
    
    ' Check if response the is a possible output
    
'    Debug.Print response
   
    ' Return response
    ClassifyTypeWithGPT = response

End Function




Private Function ParseResponse(response As String, modelName As String) As String

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

    ' If response contains an error, return error message
    If json.Exists("error") Then
    
        Debug.Print "ERROR FOUND"
        GoTo errorHandler

    End If
    
    
    ' Determine how to parse the output data
    Select Case modelName
    
        ' Chat completion - GPT-4
        Case "gpt-4", "gpt-4-0314", "gpt-4-32k", "gpt-4-32k-0314"
            ParseResponse = json("choices")(1)("message")("content")
            
        ' Chat completion - GPT-3.5
        Case "gpt-3.5-turbo", "gpt-3.5-turbo-0301"
            ParseResponse = json("choices")(1)("message")("content")

        ' Text completion
        Case "text-davinci-003", "text-davinci-002", "text-curie-001", "text-babbage-001", "text-ada-001"
            ParseResponse = json("choices")(1)("text")

        Case Else

            ' Handle Error here

    End Select
    
    ParseResponse = LCase(ParseResponse)
    ParseResponse = Trim(ParseResponse)
    ParseResponse = Replace(ParseResponse, vbNewLine, "")
    Debug.Print ParseResponse
    
    Exit Function
    
errorHandler:

    Debug.Print json("error")("message")
    ParseResponse = json("error")("message")
    
End Function



' Function to generate a GPT prompt string
' Input: inputValue - the value to be identified
'        outputsArray() - an array of possible types the value could be
' Output: a string prompt that asks GPT to identify the type of the input value from the list of possible options
Private Function GenerateGPTPromptString(inputValue As String, outputsArray() As String) As String
    
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


' Function to send an HTTP request to OpenAI's API
' Input: prompt - the prompt value for GPT to answer
'        apikey - API key to allow usage of OpenAI's API
'        modelName - OpenAI's API model
' Output: JSON response
' Output: a string prompt that asks GPT to identify the type of the input value from the list of possible options
Private Function SendGPTRequest(prompt As String, apikey As String, modelName As String) As String
    
    ' Define const variables
    Const MAX_TOKENS_PARAM As Integer = 20
    Const TEMPERATURE_PARAM As Single = 0
    
    ' Define variables
    Dim http As New MSXML2.XMLHTTP60
    Dim apiUrl As String
    Dim request As String
    Dim response As String
    
    ' Determine the appropriate API details for the request
    Select Case modelName
    
        ' Chat completion - GPT-4
        Case "gpt-4", "gpt-4-0314", "gpt-4-32k", "gpt-4-32k-0314"
            apiUrl = "https://api.openai.com/v1/chat/completions"
            request = "{""model"": """ & modelName & """, ""messages"": [{""role"": ""user"", ""content"": """ & prompt & """}]}"

        ' Chat completion - GPT-3.5
        Case "gpt-3.5-turbo", "gpt-3.5-turbo-0301"
            apiUrl = "https://api.openai.com/v1/chat/completions"
            request = "{""model"": """ & modelName & """, ""messages"": [{""role"": ""user"", ""content"": """ & prompt & """}]}"
        
        ' Text completion
        Case "text-davinci-003", "text-davinci-002", "text-curie-001", "text-babbage-001", "text-ada-001"
            apiUrl = "https://api.openai.com/v1/completions"
            request = "{""model"": """ & modelName & """, ""prompt"": """ & prompt & """, ""max_tokens"": " & MAX_TOKENS_PARAM & ", ""temperature"": " & TEMPERATURE_PARAM & "}"
            
        Case Else
        
            ' Handle Error here
        
    End Select
    
    ' Create an HTTP request with the API url and API key
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apikey

    ' Send the request and get the response
    http.send request
    response = http.responseText
    
    ' Return the response
    SendGPTRequest = response

End Function







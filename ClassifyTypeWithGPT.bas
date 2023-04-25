Attribute VB_Name = "ClassifyTypeWithGPT"
Option Explicit
Option Compare Binary


Sub ExampleMacro()

    Dim inputValue As String
    Dim outputOptions(1) As String
    Dim apiKey As String
    Dim modelName As String
    Dim result As String
    
    inputValue = "apple"
    outputOptions(0) = "fruit"
    outputOptions(1) = "vegetable"
    apiKey = "your-api-key"
    modelName = "text-davinci-003"
    
    result = ClassifyTypeWithGPT(inputValue, outputOptions, apiKey, modelName)
    
    MsgBox "The item '" & inputValue & "' was classified as: " & result

End Sub


' Function to classify an input as one of several predefined outputs
' Input: inputValue - the value to be identified
'        outputOptions() - an array of possible types the value could be classified as
'        apikey - API key to allow usage of OpenAI's API
'        modelName - OpenAI's API model
Public Function ClassifyTypeWithGPT(inputValue As String, outputOptions() As String, apiKey As String, Optional modelName As String = "gpt-3.5-turbo") As String
    
    Dim prompt As String
    Dim response As String
    
    ' Check for function parameter errors
    Call CheckInputParameters(inputValue, outputOptions(), apiKey, modelName)
    
    ' Generate GPT text prompt
    prompt = GenerateGPTPromptString(inputValue, outputOptions())

    ' Send a request to the ChatGPT API
    response = SendGPTRequest(prompt, apiKey, modelName)
    
    ' Parse the JSON response and extract the response message
    response = ParseResponse(response, modelName)
    
    ' Check if response the is a possible output
    response = VerifyResponse(response, outputOptions())
    
    ' Return response
    ClassifyTypeWithGPT = response
    
    Exit Function

End Function


' Sub to check for any immediate user entered parameter errors
' Input: inputValue - the value to be identified
'        outputOptions() - an array of possible types the value could be classified as
'        apikey - API key to allow usage of OpenAI's API
'        modelName - OpenAI's API model
Private Sub CheckInputParameters(inputValue As String, outputOptions() As String, apiKey As String, modelName As String)

    Dim errorMsg As String
    Dim i As Long

    ' Check that input value is not empty
    If Len(inputValue) = 0 Then
        errorMsg = "'inputValue' cannot be empty"
        GoTo err_input
    End If
    
    ' Check output array length
    If UBound(outputOptions) - LBound(outputOptions) + 1 < 2 Then
        errorMsg = "'outputOptions()' requires at least 2 option elements"
        GoTo err_input
    End If
    ' Check for empty array elements
    For i = LBound(outputOptions) To UBound(outputOptions)
        If Len(outputOptions(i)) = 0 Then
            errorMsg = "'outputOptions()' cannot contain empty values"
            GoTo err_input
        End If
    Next i
    
    ' Check that API key value is not empty
    If Len(apiKey) = 0 Then
        errorMsg = "'apiKey' cannot be empty"
        GoTo err_input
    End If
    
    ' Check that model value is not empty
    If Len(modelName) = 0 Then
        errorMsg = "'modelName' cannot be empty"
        GoTo err_input
    End If

    Exit Sub

err_input:
    Err.Raise 1011, "Input", errorMsg

End Sub


' Function to generate a GPT prompt string
' Input: inputValue - the value to be identified
'        outputOptions() - an array of possible types the value could be classified as
' Output: a string prompt that asks GPT to identify the type of the input value from the list of possible options
Private Function GenerateGPTPromptString(inputValue As String, outputOptions() As String) As String
    
    ' Join the elements of the outputOptions with a comma and add "or other" to the end
    Dim outputsTextPrompt As String
    outputsTextPrompt = Join(outputOptions, ", ")
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
Private Function SendGPTRequest(prompt As String, apiKey As String, modelName As String) As String
    
    ' Define const variables
    Const MAX_TOKENS_PARAM As Integer = 50
    Const TEMPERATURE_PARAM As Single = 0
    
    ' Define variables
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
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
            GoTo err_invalidModel
        
    End Select
    
    ' Create an HTTP request with the API url and API key
    http.Open "POST", apiUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    
    ' Send the request and get the response
    http.send request
    response = http.responseText
    
    ' Check if status is not 200 OK
    If http.Status <> 200 Then
        GoTo err_response
    End If
    
    ' Return the response
    SendGPTRequest = response
    
    Exit Function

err_invalidModel:
    Err.Raise 1011, "Model", modelName & " is not a valid model"
    
err_response:
    Err.Raise 2011, "Response", "APIResponseError: " & http.Status & " - " & ParseErrorResponse(response)

End Function


' Function to parse JSON response
' Input: response - JSON response
'        modelName - OpenAI's API model
' Output: the response content from the API request
Private Function ParseResponse(response As String, modelName As String) As String

    ' Parse the JSON response and extract the response message
    Dim json As Object
    Set json = JsonConverter.ParseJson(response)
    
    Debug.Print response
    
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
            GoTo err_invalidModel

    End Select
    
    ParseResponse = Trim(ParseResponse)
    ParseResponse = LCase(ParseResponse)
    ParseResponse = Replace(ParseResponse, vbNewLine, "")
    
    Debug.Print ParseResponse
    
    Exit Function

err_invalidModel:
    Err.Raise 1011, "Input", modelName & " is not a valid model"
    
End Function


' Function to parse JSON error response
' Input: response - JSON response
' Output: the response content from the API request
Private Function ParseErrorResponse(response As String) As String
    
    ' Parse the JSON response and extract the response message
    Dim json As Object
    Set json = JsonConverter.ParseJson(response)
    
    ParseErrorResponse = json("error")("message")

End Function


' Function to check if the response is one of the possible output options
' Input: response - formatted response
'        outputOptions() - an array of possible types the value could be classified as
' Output: the validated response
Private Function VerifyResponse(response As String, outputOptions() As String) As String

    If response <> "other" Then
        Dim i As Long
        For i = LBound(outputOptions) To UBound(outputOptions)
            If outputOptions(i) = response Then
                VerifyResponse = response
                Exit Function
            End If
        Next i
    End If
    VerifyResponse = "Unable to determine type classification"

End Function



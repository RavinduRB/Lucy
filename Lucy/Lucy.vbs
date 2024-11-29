Option Explicit

Dim userResponse, botResponse

' Start the chatbot
MsgBox "Welcome to Princess Lucy! Say 'bye' to end the chat.", vbInformation, "Princess Lucy"

Do
    ' Get user input
    userResponse = LCase(InputBox("You: ", "Princess Lucy Chat"))

    ' Exit loop if user says "bye"
    If userResponse = "bye" Then
        botResponse = "Bye! Take care! You're amazing! You are loved!"
        MsgBox "Lucy: " & botResponse, vbInformation, "Princess Lucy"
        Exit Do
    End If

    ' Generate bot response based on keywords
    botResponse = GetBotResponse(userResponse)
    
    ' Display bot response
    MsgBox "Lucy: " & botResponse, vbInformation, "Princess Lucy"

Loop

' Function to generate responses
Function GetBotResponse(userInput)
    Select Case True
        Case InStr(userInput, "hello") > 0 Or InStr(userInput, "hi") > 0 Or InStr(userInput, "hey") > 0
            GetBotResponse = "Hi! You brighten my day! I love you!"
        Case InStr(userInput, "i love you") > 0
            GetBotResponse = "I love you too! You're wonderful!"
        Case InStr(userInput, "miss you") > 0
            GetBotResponse = "I miss you too! You are doing great!"
        Case InStr(userInput, "what are you doing now") > 0
            GetBotResponse = "I'm looking at you! Keep smiling!"
        Case Else
            GetBotResponse = "I don't quite understand. You are amazing! Keep being you!"
    End Select
End Function

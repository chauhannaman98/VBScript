Const SITE_TITLE = ">> CodingGears.io"
Dim input1, input2, sum, name

name = InputBox("Enter your name: ")
input1 = InputBox("Enter the first number: ", SITE_TITLE, "Enter input here")
input2 = InputBox("Enter the second number: ", SITE_TITLE, "Enter input here", 1000, 5000)

' Processing
sum = CInt(input1) + CInt(input2)

' Displaying a message
MsgBox "Hello: " & name & " !!! ", 0, SITE_TITLE
MsgBox "The sum of the two numbers: " & sum, 64, SITE_TITLE

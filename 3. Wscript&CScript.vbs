' WScript.exe ==> Windows apps
' CSCript.exe ==> Console apps

' WScript.exe --> Executable for WSH that interprests VBScript lang
' WScript --> An object in Core object model in WSH

MsgBox("1 Hello World")
WScript.Echo "2 Hello World"

' Sleep for 5 seconds
WScript.Sleep 5000
WScript.Echo "3 Hello World"
WScript.Quit
WScript.Echo "4 Hello World"


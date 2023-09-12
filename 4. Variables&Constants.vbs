' Dim keyword is short for Dimension
' & for concatenation
' &_ for line concatenation

Dim num1, num2
Dim sum

Const bonus = 5

num1 = 10
num2 = 20

sum = num1 + num2

MsgBox sum
MsgBox "The sum of " & num1 & " and " & num2 & " is " & sum
MsgBox "==> The sum of " & num1 & " and " & num2 _
        & " is " & sum

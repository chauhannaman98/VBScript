option explicit

Dim PhoneBook()

ReDim PhoneBook(2, 0)
PhoneBook(0,0) = "Peter"
PhoneBook(1,0) = "Boston"
PhoneBook(2,0) = "111-111-0000"

MsgBox "1: " & PhoneBook(0,0)

ReDim Preserve PhoneBook(2, 2)   '3 cols, 3 rows

MsgBox "2: " & PhoneBook(0,0)

' Following line throws exception as the last dimension
' in PhoneBook(5, 2) cannot be modified.
' ReDim Preserve PhoneBook(5, 2)   '6 cols, 3 rows

ReDim Preserve PhoneBook(2, 5)   '6 cols, 3 rows

MsgBox "3: " & PhoneBook(0,0)

ReDim PhoneBook(12, 20)   '13 cols, 21 rows

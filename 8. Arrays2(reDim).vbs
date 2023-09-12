' ReDim - Change the size of last dimension

' error handling
option explicit

' Variables
Const SITE_TITLE = "www.CodingGears.com"

Dim arrFriends()
ReDim Preserve arrFriends(2)    ' use Preserve to retain the old values after redim
arrFriends(0) = "John"
arrFriends(1) = "Peter"
arrFriends(2) = "Jimmy"

MsgBox arrFriends(1), 0, SITE_TITLE
' MsgBox arrFriends(4), 0, SITE_TITLE ' This will throw exception

ReDim Preserve arrFriends(4)
arrFriends(3) = "Mike"
arrFriends(4) = "Kenny"

MsgBox arrFriends(4), 0, SITE_TITLE

Redim Preserve arrFriends(9)
MsgBox arrFriends(8), 0, SITE_TITLE

option explicit

Const SITE_TITLE = "www.CodingGears.com"
Dim arrAnimals
Dim matched, not_matched, item

' Filter(Filter(inputstrings, value[,include[,compare]]))
' Compare:
'   0 - vbBinaryCompare (Binary comparison)
'   1 - vbTextCompare (Textual comparison)

arrAnimals = Array("spotted", "attacked", "loss", "read", "tracked")

matched = Filter(arrAnimals, "ed", True, 1)

MsgBox matched(0)
MsgBox matched(1)
MsgBox matched(2)

not_matched = Filter(arrAnimals, "ed", False, 1)

MsgBox not_matched(0)
MsgBox not_matched(1)

option explicit

Const SITE_TITLE = "www.CodingGears.com"
Dim lowerIndex, upperIndex
Dim arrNums(4)

arrNums(0) = 100
arrNums(1) = 200
arrNums(2) = 300
arrNums(3) = 400
arrNums(4) = 500

lowerIndex = LBound(arrNums)
upperIndex = UBound(arrNums)

MsgBox "Lower Index is " & lowerIndex, 0, SITE_TITLE
MsgBox "Upper Index is " & upperIndex, 0, SITE_TITLE



' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Dim num1, b, x, bnot
Dim arrayNums(2)
Dim arrAnimals

arrayNums(0) = 100
arrayNums(1) = 200

MsgBox "Checking if the arrayNums is an array: " _
        & IsArray(arrayNums), 0, SITE_TITLE
MsgBox "Checking if the num1 is an array: " _
        & IsArray(num1), 0, SITE_TITLE
        
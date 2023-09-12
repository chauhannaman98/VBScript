option explicit

Const SITE_TITLE = "www.CodingGears.com"

Dim arrTitle1
arrTitle1 = Split(SITE_TITLE, ".")

WScript.Echo arrTitle1(0)
WScript.Echo arrTitle1(1)
WScript.Echo arrTitle1(2)


' +++++++++++++++++++++++++++++++++++++++++++++++

Dim arrWebsite(2)

arrWebsite(0) = "www"
arrWebsite(1) = "CodingGears"
arrWebsite(2) = "io"

Dim siteUrl
siteUrl = Join(arrWebsite, ".")

WScript.Echo(siteUrl)

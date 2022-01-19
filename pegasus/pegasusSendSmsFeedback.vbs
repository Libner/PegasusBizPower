Option Explicit
On Error Resume Next

' Declare our vars
Dim objWinHttp, strURL

' Request URL from 1st Command Line Argument.  This is
' a nice option so you can use the same file to
' schedule any number of differnet scripts just by
' changing the command line parameter.
''strURL = WScript.Arguments(0)
''--020220---strURL = "http://pegasus.bizpower.co.il/pegasus/SendfeedBack_Form.aspx"
strURL = "http://pegasus.bizpower.co.il/pegasus/SendfeedBack_Form_IsTravelers.aspx"

' Could also hard code if you want:
'strURL = "http://localhost/ScheduleMe.asp"

' For more WinHTTP v5.0 info, including where to get
' the component, see our HTTP sample:
' http://www.asp101.com/samples/winhttp5.asp
Set objWinHttp = CreateObject("MSXML2.ServerXMLHTTP")

objWinHttp.Open "GET", strURL
objWinHttp.Send

' Get the Status and compare it to the expected 200
' which is the code for a successful HTTP request:
' http://www.asp101.com/resources/httpcodes.asp
If objWinHttp.Status <> 200 Then
	' If it's not 200 we throw an error... we'll
	' check for it and others later.
	Err.Raise 1, "HttpRequester", "Invalid HTTP Response Code"
End If

' Since in this example I could really care less about
' what's returned, I never even check it, but in
' general checking for some expected text or some sort
' of status result from the ASP script would be a good
' idea.  Use objWinHttp.ResponseText

Set objWinHttp = Nothing
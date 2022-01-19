Option Explicit
'On Error Resume Next

' Declare our vars
Dim objWinHttp, strURL

Set objWinHttp = CreateObject("MSXML2.ServerXMLHTTP")

'local==========
strURL = "http://192.168.0.92/pegasusnew_mila/biz_form/sendReservationForm_ScheduleTask.ashx"

'production================================
'strURL = "https://www.pegasusisrael.co.il/biz_form/sendReservationForm_ScheduleTask.ashx"

objWinHttp.Open "GET", strURL, False,  "", ""


objWinHttp.Send

If objWinHttp.Status <> 200 Then
	Err.Raise 1, "HttpRequester", "Invalid HTTP Response Code"
End If

Set objWinHttp = Nothing

If Err.Number <> 0  Then
	' Something has gone wrong... do whatever is
	' appropriate for your given situation... I'm
	' emailing someone:
	Dim objMessage
	Set objMessage = Server.CreateObject("CDO.Message")
	objMessage.From       = "mila@cyberserve.co.il"
	objMessage.To     = "mila@cyberserve.co.il"
	objMessage.Subject  = "An Error Has Occurred in a " _
		& "Scheduled Task PAIS loadLottoXmlJob"
	objMessage.TextBody = "Error #: " & Err.Number & vbCrLf _
		& "From: " & Err.Source & vbCrLf _
		& "Desc: " & Err.Description & vbCrLf _
		& "Time: " & Now()
							
	objMessage.Send
	Set objMessage = Nothing
End If

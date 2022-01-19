<%@Language=VBScript%>
<%For Each SessionItem in Session.Contents
	If InStr(SessionItem,"admin") = 0 Then
			Session.Contents.Remove(SessionItem)
	End If
		
 Next
''	 Response.Cookies("bizpegasus")=""
	 
''	 response.Write "1="& Request.Cookies("bizpegasus")("UserId")
	' response.end

''	Request.Cookies("bizpegasus").Expires = Now() - 100
    Response.Cookies("bizpegasus").Domain="bizpower.co.il"
	Response.Cookies("bizpegasus").Expires = Now()-100
	'Response.Cookies("bizpegasus") = ""
	''response.Cookies("bizpegasus")("UserId")=""
	'' response.Write "2="& Request.Cookies("bizpegasus")("UserId")

	Response.Redirect "default.asp" %>

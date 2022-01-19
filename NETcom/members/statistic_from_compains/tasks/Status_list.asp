<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->

<%
  If Request.QueryString("OrgId") = nil Then
		OrgId = trim(Request.Cookies("bizpegasus")("OrgId")) 
	Else
		OrgId = trim(Request.QueryString("OrgId"))
	End If	  
UrlStatus="default.asp?search_company="& Server.URLEncode(search_company) &"&amp;search_contact="& Server.URLEncode(search_contact)&"&amp;search_project="&Server.URLEncode(search_project)&"&amp;start_date="&start_date&"&amp;end_date="&end_date & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id & "&amp;tasktypeID=" & tasktypeID & "&amp;T=" & T
	 arr_Status = Array("","חדש","בטיפול","סגור")	
   
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
</head>
<body style="margin:0px;background-color:#e6e6e6" onload="window.focus()">
<table border=0 cellpadding=0 cellspacing=0 width=100%>
<%For i=1 To uBound(arr_Status)	%>

	<tr><TD width=100% align=right  onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; border-bottom:1px solid black; cursor:hand;"
	ONCLICK="self.close();opener.location.href='<%=UrlStatus%>&amp;task_status=<%=i%>'">&nbsp;<%=arr_Status(i)%>&nbsp;</td></tr>
	
	<%Next%>    

	<tr><TD width=100% align=right  onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand;"
	ONCLICK="self.close();opener.location.href='<%=UrlStatus%>&amp;task_status=1,2,3'">
    &nbsp;כל הרשימה&nbsp;</td></tr>

</table>
</body>

<%set con=nothing%>
</html>

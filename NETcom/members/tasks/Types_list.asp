<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->

<%
  If Request.QueryString("OrgId") = nil Then
		OrgId = trim(Request.Cookies("bizpegasus")("OrgId")) 
	Else
		OrgId = trim(Request.QueryString("OrgId"))
	End If	  
	urlType="default.asp?search_company="& Server.URLEncode(search_company) &"&amp;search_contact="& Server.URLEncode(search_contact) &"&amp;search_project="&Server.URLEncode(search_project)&"&amp;start_date="&start_date&"&amp;end_date="&end_date&"&amp;task_status="&task_status & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id  & "&amp;T=" & T

%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
</head>
<body style="margin:0px;background-color:#e6e6e6" onload="window.focus()">
<table border=0 cellpadding=0 cellspacing=0 width=100%>
<%  sqlstr = "Select activity_type_id, activity_type_name from activity_types WHERE ORGANIZATION_ID = " & OrgID & " Order By activity_type_id"
	set rstask_type = con.getRecordSet(sqlstr)
	while not rstask_type.eof %>
	<tr><TD width=100% align=right  onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; border-bottom:1px solid black; cursor:hand;"
	ONCLICK="self.close();opener.location.href='<%=urlType%>&amp;tasktypeID=<%=rstask_type(0)%>&amp;task_status=1,2,3'">&nbsp;<%=rstask_type(1)%>&nbsp;</td></tr>
	<%
		rstask_type.moveNext
		Wend
		set rstask_type=Nothing
	%>
	<tr><TD width=100% align=right  onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand;"
	ONCLICK="self.close();opener.location.href='<%=urlType%>'">
    &nbsp;כל הרשימה&nbsp;</td></tr>

</table>
</body>

<%set con=nothing%>
</html>

<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%	UserID=trim(Request.Cookies("bizpegasus")("UserID"))
	OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
	lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
	perSize = trim(Request.Cookies("bizpegasus")("perSize"))
	is_groups = trim(Request.Cookies("bizpegasus")("ISGROUPS"))
	
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"
		align_var = "left"
		dir_obj_var = "ltr"
	Else
		dir_var = "ltr"
		align_var = "right"
		dir_obj_var = "rtl"
	End If		  	
  
	pageID = trim(Request.Form("pageID"))
	categID = trim(Request.Form("categID"))	
	   
	sqlStr = "Select Page_Title From Pages Where Page_ID=" & pageID
	set prod = con.GetRecordSet(sqlStr)
	if not prod.eof then
		PageTitle=trim(prod("Page_Title"))
	end if
	set prod = nothing%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css" />
</head>
<body marginwidth="0" marginheight="0" hspace="0" vspace="0" topmargin="0" leftmargin="0" rightmargin="0">
<table border="0" width="750" cellspacing="0" cellpadding="0" align=center>
<tr>
<td nowrap align="center">
<!--#INCLUDE FILE="../reports/logo_top.asp"-->
</td>
<td width=480>
<table border="0" width="480" cellspacing="0" cellpadding="0" align=center>
	<tr>
		<td class="card5" style="font-size:16pt" align="right" <%=dir_obj_var%>><b><%If trim(lang_id) = "1" Then%>דוח כניסות
         <%Else%>Bar distribution report<%End If%></b></td>
	</tr>
	<tr>
		<td class="card5" style="font-size:13pt" dir="<%=dir_obj_var%>" align="right"><b><%=PageTitle%></b></td>
	</tr>	
</table>	
</td>
</tr> 
<tr><td align="center" colspan="2">
<table border="0" width="750" cellspacing="0" cellpadding="0" align=center>
	<tr><td height=10 nowrap></td></tr>
	<tr><td><img src="combine_graph.asp?pageID=<%=pageID%>&categID=<%=Server.HTMLEncode(categID)%>"></td></tr>
	<!--#INCLUDE FILE="../reports/bottom_inc.asp"-->	
</table>
</td></tr></table>
</body>
<%set con = nothing%>
</html>
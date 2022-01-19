<!--#include file="../include/connect.asp"-->
<!--#include file="../include/reverse.asp"-->
<%
ID=Request.QueryString("ID")

	if trim(ID)<>"" then
		sqlStr = "SELECT New_Title,Page_Width,New_Date,new_Time_on,new_Time_off,New_Content,New_Desc, Category_ID  FROM News  WHERE new_ID="& ID 
		set rs_news = con.execute(sqlStr)
		if not rs_news.eof then
			Page_Width = rs_news("Page_Width")
			New_Content = rs_news("New_Content")
			New_Title = rs_news("New_Title")	 
			Page_Width = rs_news("Page_Width")
			New_Desc = rs_news("New_Desc") 
			New_Date = rs_news("New_Date")
			Category_ID = rs_news("Category_ID")								
		end if
		
		set rs_news = nothing
	end if
%>
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="<%=Application("VirDir")%>/dynamic_style.css" rel="STYLESHEET" type="text/css">
<title>Bizpower - News - <%=New_Title%></title>
<!--#include file="../../include/title_meta_inc.asp"-->
</head>
<body marginwidth="0" marginheight="0" topmargin="5" leftmargin="0" rightmargin="0" bgcolor="white">
<!--#include file="../include/top.asp"-->
<table dir="rtl" border="1" style="border-collapse:collapse" bordercolor="#999999" cellspacing="0" cellpadding="0" width="780" align="center" ID="Table1">
	<tr><td colspan=2 bgcolor="#FFFFFF" align="center" width="780">
	<div align="center">
			<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"
				ID="Shockwaveflash1" VIEWASTEXT width=778>
				<PARAM NAME="movie" VALUE="../../flash/BP-flash-test.swf">
				<PARAM NAME="quality" VALUE="best">
				<PARAM NAME="wmode" VALUE="transparent">
				<PARAM NAME="bgcolor" VALUE="#EBEBEB">					
				<EMBED src="../../flash/BP-flash-test.swf" quality="best" wmode="transparent" bgcolor="#FFFFFF"
					TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"
					width="778">
				</EMBED>
			</OBJECT>
	</div>				
	</td></tr>							
   <tr>
		<td  align="left" class="title_page" bgcolor="#EBEBEB" dir="rtl"
			style="padding:5px;padding-right:10px">
		<span class="headDate"><%=New_Date%></span>
		<br>
		<span class="headTitle"><%=New_Title%></span>
	    </td>
		<td rowspan=2 dir="ltr" align="left" valign="top" width="255" nowrap bgcolor="#FFFFFF">
			<!--#include file="../include/ashafim_inc.asp"-->
		</td>	    
	</tr>							
							
	<tr>
		<td width="520" nowrap bgcolor="#F9F9F9" valign="top" align="left" style="padding-right:20px; padding-left:20px;">
        <table border="0" cellspacing="0" cellpadding="0" width="480" dir="ltr" ID="Table2">
			<%'//start of centent%>				
				<tr>
					<td width="100%" valign="top" align="left" height="15"></td>
				</tr>
				<tr>
					<td width="100%" bgcolor="#F9F9F9" height="300" valign="top" align="left">
					<table border="0" width="<%if trim(page_Width) <> "" then %><%=page_Width%><%else%>100%<%end if%>"
					cellspacing="0" cellpadding="0" ID="Table3">
					<tr>
					<td align="left" width="100%" >
						<table align="left" width="100%" border="0" cellspacing="0" cellpadding="0" ID="Table4">
							
							<tr><td width="100%" height="2"></td></tr>
							<tr><td align="left" dir=rtl width="100%" valign="top"><b><%=Replace(New_Desc,vbCrLf,"<br>")%></b></td></tr> 
							<tr><td width="100%" height="10"></td></tr>
							<tr>
							<td align="left" width="100%" valign="top"><%=New_Content%></td>
							</tr> 
							<tr><td width="100%" height="20"></td></tr>
							<tr><td align="left" width="100%" dir="ltr"><a href="" onclick="javascript:history.back(); return false;" class="btnLink" style="width:50px;">« Back</a></td></tr>
							<tr><td width="100%" height="10"></td></tr>
						</table>
				</td></tr></table></td></tr></table>

		</td>									
	</tr>					
</table>
<!--#include file="../include/bottom.asp"-->
</body>
</html>

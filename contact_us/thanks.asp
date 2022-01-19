<!--#include file="../include/connect.asp"-->
<!--#include file="../include/reverse.asp"-->
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="<%=Application("VirDir")%>/dynamic_style.css" rel="STYLESHEET" type="text/css">
<!--#include file="../include/title_meta_inc.asp"-->
<title>Bizpower - טופס הצטרפות</title>
</head>

<body marginwidth="0" marginheight="0" hspace="0" vspace="0" topmargin="5" leftmargin="0"
		rightmargin="0" bgcolor="white">
<!--#include file="../include/top.asp"-->
<table dir="rtl" border="1" style="border-collapse:collapse" bordercolor="#999999" cellspacing="0" cellpadding="0" width="780" align="center" ID="Table5">
	<tr><td colspan=2 bgcolor="#FFFFFF" align="center" width="780">
	<div align="center">
			<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"
				ID="Shockwaveflash1" VIEWASTEXT width=780 height="175">
				<PARAM NAME="movie" VALUE="../flash/BP-flash-test.swf">
				<PARAM NAME="quality" VALUE="best">
				<PARAM NAME="wmode" VALUE="transparent">
				<PARAM NAME="bgcolor" VALUE="#EBEBEB">					
				<EMBED src="../flash/BP-flash-test.swf" quality="best" wmode="transparent" bgcolor="#FFFFFF"
					TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"
					width="780" height="175">
				</EMBED>
			</OBJECT>
	</div>				
	</td></tr>
	<tr>
			<td  align="right" class="title_page" dir="rtl"	style="padding:5px;padding-right:10px">
			טופס הצטרפות													
			</td>
			<td rowspan=2 align="right" valign="top" width="180" nowrap bgcolor="#FFFFFF">
		    <!--#include file="../include/ashafim_inc.asp"-->
		    </td>														
		</tr>			
	<tr>
		<td width="600" height=280 nowrap bgcolor="#F9F9F9" valign="top" align="right" style="padding-right:20px; padding-left:20px;">
        <table border="0" cellspacing="0" cellpadding="0" width="480" dir="ltr">
        <tr><td height="20" nowrap></td></tr>				
		<%'//start of content%>			
		<%
		name = Request.Form("name")

		if request.form("tafkid")<>nil then
			tafkid = Request.Form("tafkid")
		else
			tafkid = "&nbsp;--&nbsp;"
		end if
		company   = Request.Form("company")
		phone     = Request.Form("phone")

		email     = Request.Form("email")
		if request.form("address")<>nil then
			address = Request.Form("address")
		else
			address = "&nbsp;--&nbsp;"
		end if
		if request.form("date")<>nil then
			Ddate = Request.Form("date")
		else
			Ddate = "&nbsp;--&nbsp;"
		end if
		comments = Request.Form("comments")

	BodyText=""

	BodyText = BodyText &  "<html><head><meta charset=windows-1255></head>"
	BodyText = BodyText &  "<style>"
	BodyText = BodyText &  "td.content_cl"
	BodyText = BodyText &  "{"
	BodyText = BodyText &  vbTab & "FONT-WEIGHT: bold;"
	BodyText = BodyText &  vbTab & "FONT-SIZE: 10pt;"
	BodyText = BodyText &  vbTab & "COLOR: #022c63;"
	BodyText = BodyText &  vbTab & "BACKGROUND-COLOR: #e0e0f4;"
	BodyText = BodyText &  vbTab & "FONT-FAMILY: Arial;"
	BodyText = BodyText &  "}"
	BodyText = BodyText &  "td.title_cl"
	BodyText = BodyText &  "{"
	BodyText = BodyText &  vbTab & "FONT-WEIGHT: 600;"
	BodyText = BodyText &  vbTab & "FONT-SIZE: 11pt;"
	BodyText = BodyText &  vbTab & "COLOR: #6f6da6;"
	BodyText = BodyText &  vbTab & "FONT-FAMILY: Arial;"
	BodyText = BodyText &  "}"
	BodyText = BodyText &  "td.content_b"
	BodyText = BodyText &  "{"
	BodyText = BodyText &  vbTab & "FONT-WEIGHT: 500;"
	BodyText = BodyText &  vbTab & "FONT-SIZE: 10pt;"
	BodyText = BodyText &  vbTab & "BACKGROUND-COLOR: #f7f7f7;"
	BodyText = BodyText &  vbTab & "COLOR: #022c63;"
	BodyText = BodyText &  vbTab & "FONT-FAMILY: Arial;"
	BodyText = BodyText &  "}"
	BodyText = BodyText &  "</style>"
	BodyText = BodyText &  "<body>"


	'*************************************************************************************

	BodyText = BodyText & "<table width=60% cellpadding=3 cellspacing=2 align='center' border='0'>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td colspan=2 class='title_cl' dir=rtl align='right' >" & "פנייה ציבורית לצוות BizPower" & "</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td class='content_b' dir=rtl align='right' >" & name & "&nbsp;</td>"
	BodyText = BodyText & "<td class='content_cl' align='right' nowrap>:שם פרטי ושם משפחה</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td class='content_b' dir=rtl align='right'>" & tafkid & "&nbsp;</td>"
	BodyText = BodyText & "<td class='content_cl' align='right'>:תפקיד</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td class='content_b' dir=rtl align='right'>" & company & "&nbsp;</td>"
	BodyText = BodyText & "<td class='content_cl' align='right'>:חברה</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td class='content_b' dir=rtl align='right'>" & address & "&nbsp;</td>"
	BodyText = BodyText & "<td class='content_cl' align=right>:כתובת</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td class='content_b' dir=rtl align='right'>" & phone & "&nbsp;</td>"
	BodyText = BodyText & "<td class='content_cl' align='right'>:טלפון</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td class='content_b' align='right'><a href=mailto:" & email & " target=_blank>" & email & "</a></td>"
	BodyText = BodyText & "<td class='content_cl' align='right' nowrap>:דואר אלקטרוני</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "<tr>"
	BodyText = BodyText & "<td align='right' dir=rtl class='content_b' nowrap width='100%'>" & comments & "</td>"
	BodyText = BodyText & "<td class='content_cl' align=right nowrap>:התייחסות חופשית</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "</table>"
	BodyText = BodyText & "</td>"
	BodyText = BodyText & "</tr>"
	BodyText = BodyText & "</table>"

	'*************************************************************************************

	BodyText = BodyText & "</body>"
	BodyText = BodyText & "</html>"

		Dim Msg
			Set Msg = Server.CreateObject("CDO.Message")
				Msg.BodyPart.Charset = "windows-1255"
			if trim(email)<>"" and instr(1,email,"@")>1 and instr(1,email,".")>3 and instr(1,email,"@") < instr(1,email,".") then 
			Msg.From = email
			else   
			Msg.From   =  "support@pegasusisrael.co.il"
		end if
				Msg.MimeFormatted = true
				Msg.Subject = "פנייה ציבורית לצוות פגסוס"
				Msg.To = "adina@eltam.com;"				
				'Msg.To = "olga@eltam.com;"
				Msg.HTMLBody = BodyText
			
				Msg.Send()						
			Set Msg = Nothing

%>	

		<tr>
		<td align="center">
		<b>
	     <%if err.number>0 then
				Response.write "הפניתך לא נשלחה, אנא נסה שוב מאוחר יותר"
				err.Clear
            else
				%>
				<font color=#93A59B><span style="font-size:20pt;" >תודה רבה</span><br>&nbsp;
				נציגנו ייצרו איתך קשר</font>
				<%
			end if
		%>
		</b>		
<%'//end of content%>																				
		</td></tr>														
		<%'//end of content%>
		</table>
		</td>										
	</tr>					
</table>
<!--#include file="../include/bottom.asp"-->
</body>
</html>

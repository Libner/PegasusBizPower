<%@Language=VBScript%>
<!--#include file="include/connect.asp"-->
<!--#include file="include/reverse.asp"-->
<%	'OrgID = trim(Request("OrgID"))
	'PEOPLE_ID = trim(Request("PEOPLE_ID"))
	
	OrgID = Decode(Request(Encode("OrgID")))
	PEOPLE_ID = Decode(Request(Encode("PEOPLE_ID")))
	
	If isNumeric(OrgID) And IsNumeric(PEOPLE_ID) Then
	
	If trim(OrgID) <> "" Then
		sqlstr = "Select ORGANIZATION_NAME, ORGANIZATION_LOGO, IsNULL(LANG_ID, 1) as LANG_ID From ORGANIZATIONS Where ORGANIZATION_ID = " & OrgID
		set rs_org = con.Execute(sqlstr)
		If not rs_org.eof Then
			OrgName = rs_org(0)
			sizeLogoOrg = rs_org(1).ActualSize
			langId = cint(rs_org("LANG_ID"))
		End If
		set rs_org = Nothing
		
		If trim(PEOPLE_ID) <> "" Then
			sqlstr = "Select PEOPLE_EMAIL From PEOPLES Where PEOPLE_ID = " & PEOPLE_ID
			set rs_email = con.Execute(sqlstr)
			If not rs_email.eof Then
				People_Email = rs_email(0)
			End If
			set rs_email = nothing	
		End If	
	Else
		Response.Redirect "/"	
	End If
	
	If Len(OrgName) > 0 And Len(People_Email) > 0 Then
	
	If Request.Form("People_Email") <> nil Then
		sqlstr = "Select blocked_email From blocked_emails Where blocked_email Like '" & sFix(Request.Form("People_Email")) &_
		"' And ORGANIZATION_ID = " & OrgID
		set rs_check = con.Execute(sqlstr)
		If rs_check.eof Then
			is_checked = false
			sqlstr = "Insert Into blocked_emails (ORGANIZATION_ID,blocked_email,date_add) Values "&_
			"(" & OrgID & ",'" & sFix(Request.Form("People_Email")) & "', getDate())"
			con.Execute(sqlstr)
		Else
			is_checked = true	
		End If
		set rs_check = nothing			
	End If
%>
<html>
<head>
<!-- #include file="netcom/title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="<%=Application("VirDir")%>/dynamic_style.css" rel="STYLESHEET" type="text/css">
</head>
<body marginwidth="10" marginheight="0" topmargin="10" leftmargin="10" rightmargin="10" bgcolor="white">
<%If sizeLogoOrg > 0 Then%>     
<p align=center><img src="<%=strLocal%>/netcom/GetImage.asp?DB=ORGANIZATION&amp;FIELD=ORGANIZATION_LOGO&amp;ID=<%=OrgId%>"></p>
<%End If%>
<table cellpadding=5 cellspacing=0 width=600 align=center dir="ltr">
<tr><td height=50 nowrap colspan=2></td></tr>
<%If Request.Form("People_Email") = nil Then%>
<tr>
<td class="titleB">
<%If langId = 1 Then%><span  align=right dir="rtl" style="width: 100%">האם ברצונך להסיר את כתובת הדואר האלקטרוני שלך מרשימת התפוצה של  <%=OrgName%>?</span><%Else%>
 <span align="left" dir="ltr">Are you sure wish to unsubscribe and receive no more messages from <%=OrgName%> ?</span>
<%End If%>
</td>
</tr>
<tr>
<td class="titleB"><%If langId = 1 Then%><span align="right" dir="rtl" style="width: 100%">כתובת דואר אלקטרוני : <%=People_Email%></span><%Else%> Email address:  <%=People_Email%> <%End If%>
</td>
</tr>
<form name=form1 id=form1 action="" method=post>
<tr><td align=right>
<input type=hidden name=OrgID id=OrgID value="<%=OrgID%>">
<input type=hidden name=PEOPLE_ID id="PEOPLE_ID" value="<%=PEOPLE_ID%>">
<input type=hidden name="People_Email" id="People_Email" value="<%=vFix(People_Email)%>">
</td></tr>
<tr><td align=center><input type=submit value="<%If langId = 1 Then%>הסר<%Else%>unsubscribe<%End If%>" class="button_small" style="width:100px;"></td>
</form>
</tr>
<%Else%>
<tr>
<td class="titleB"><%If langId = 1 Then%><span  align=right dir="rtl" style="width: 100%">כתובת דואר אלקטרוני <%=Request.Form("People_Email")%> הוסרה מרשימת התפוצה של <%=OrgName%>.</span><%Else%>You have been removed from <%=OrgName%> mailing list.<%End If%></td>
</tr>
<%End If%>
</table>
</body>
</html>
<%End If%>
<%End If%>
<%Set con = Nothing%>
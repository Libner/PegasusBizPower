<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<% 	

    if request("submit1")<>nil then 
		if trim(Request.QueryString("groupId")) <> "" then
			 sqlstr = "Update groups set groupName='"& sFix(Trim(Request.Form("groupName"))) &"' Where groupId=" &Request.QueryString("groupId")& " And ORGANIZATION_ID=" & OrgID
			 'Response.Write sqlstr
			 con.ExecuteQuery(sqlstr)
		else
			 sqlstr = "Insert into groups(groupName,GROUPDATE,ORGANIZATION_ID,User_ID) values('"& sfix(Trim(Request.Form("groupName"))) &"',getDate()," & OrgID & "," & UserId &")"
			 'Response.Write sqlstr
			 con.ExecuteQuery(sqlstr)	
			
			 sqlStr = "select groupId from groups where ORGANIZATION_ID="& OrgID &" order by groupId desc"
			 ''Response.Write	sqlStr
			 ''Response.End
			 set recLastGroup = con.GetRecordSet(sqlStr)
			 if not recLastGroup.eof then
			 	groupId=recLastGroup("groupId")
			 end if	
			 set recLastGroup = Nothing	
		end if
		
		adding_type = trim(Request.Form("adding_type"))
		If trim(adding_type) = "excel" Then
			strUrl = "excelUpload.asp?groupID=" & groupId
			h = 200
			w = 450
		Else
			strUrl = "getContacts.asp?groupID=" & groupId
			h = 460
			w = 450
		End If			
%>
	<SCRIPT  LANGUAGE=javascript>
	<!--		
		window.document.location.href = "groupAdd.asp?groupID=<%=groupID%>"
		window.open('<%=strUrl%>',null,'alwaysRaised,left=180,top=50,height=<%=h%>,width=<%=w%>,status=no,toolbar=no,menubar=no,location=no');		
		
	//-->
	</SCRIPT>
<%	end if %>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function button1_onclick() {
	this.close();
}
function checkfields() {
	if (document.form1.groupName.value == '')
	{	alert('אנא מלאו שם הקבוצה');
		document.form1.groupName.focus();
		return false;
	}
	var chk_flag = false;
	for (var i=0; i < document.form1.adding_type.length; i++)
    {
	  if (document.form1.adding_type[i].checked)
      {
         chk_flag = true;
         break;
      }
    }
    if(chk_flag == false)
    {
		window.alert("חובה לבחור אופן טעינת נמענים לקבוצה");
		return false;
    }
	return true;	
}
//-->
</SCRIPT>
</head>

<body>
<!-- #include file="../../logo_top.asp" -->
<%numOftab = 2%>
<!--#include file="../../top_in.asp"-->
<table border="0" bordercolor="navy" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF ID="Table1">
<tr><td colspan=2 background="../../images/Line_3D.gif" height=2 width="100%"></td></tr>
<tr>
<td width="100%" class="page_title" dir=rtl>&nbsp;<%If IsNumeric(wizard_id) Then%> <%=page_title_%> <%Else%>יצירת קבוצת דיוור חדשה<%End If%>&nbsp;</td></tr>         
<tr><td colspan=2 background="../../images/Line_3D.gif" height=2 width="100%"></td></tr>
<%If trim(wizard_id) <> "" Then%>
<tr><td width="100%">
<% wizard_page_id = 2 %>
<!--#include file="../../wizard_inc.asp"-->
</td></tr>
<tr>
<td width=100% align="right" bgcolor="#FFD011" style="padding:5px">
<table border=0 cellpadding=0 cellspacing=0 width=100% ID="Table3">
<tr><td class="explain_title">
יצירת קבוצת דיוור חדשה?
</td></tr>
<tr><td height=5 nowrap></td></tr>
<tr><td class=explain>
באפשרותך ליצור קבוצה חדשה באחד משני אופנים:
</td></tr>
<tr><td class=explain>
1.	טעינת נמענים מקובץ Excel. באפשרות זו תקבל טופס Excel מובנה, אליו תוכל להקליד או להעתיק כתובות דוא"ל מרשימת הדיוור, ולטעון את הקבוצה למערכת.
</td></tr>
<tr><td class=explain>
2.	טעינה ממועדון ארגונים. אפשרות זו תאפשר לך לשלוף נמענים לדיוור מתוך מועדון הארגונים שבנית במערכת BizPower, על פי קריטריונים מוגדרים מראש (תפקידים, מגזר ארגונים, וכיו"ב)
</td></tr>
<tr><td class=explain>
הקלד שם לקבוצת הדיוור שברצונך ליצור, סמן את אופן בניית הקבוצה הרצוי לך, ולחץ <b>[המשך >>]</b>.
</td></tr>
</table></td></tr>
<tr><td height=2 nowrap></td></tr>
<%End If%>
<tr><td width="100%">
<table cellpadding=0 cellspacing=0 width=100% bgcolor="#E6E6E6">
<tr><td height=10 nowrap></td></tr>
<tr><td align=right>
<%
	if trim(groupId) <> "" then
	
		sqlGroup="SELECT groupId, groupName FROM groups WHERE groupId ="&groupId& " and ORGANIZATION_ID=" & OrgID 
		set clientGroups = con.GetRecordSet(sqlGroup)
		if not clientGroups.eof then
			groupId=clientGroups("groupId")
			groupName=clientGroups("groupName")
		end if
	end if	
%>
<form action="groupAdd.asp?groupId=<%=groupId%>" method=POST id=form1 name=form1 onsubmit="return checkfields();">
<table align=center border="0" cellpadding="1" cellspacing="1" width="60%" ID="Table2">
	<tr>
		<td align="right" width=100%><input type="text" class="texts" dir="rtl" value="<%=vfix(groupName)%>" size=40 name="groupName" ID="Text1">&nbsp;</td>
		<td align="right" width="65" nowrap>&nbsp;שם קבוצה&nbsp;&nbsp;</td>		
	</tr>
	<tr><td colspan=2><hr noshade size=1 color="#000000"></td></tr>	
	<tr><td colspan=2 align=right class="title_table">אופן טעינת נמענים לקבוצה</td></tr>
	<tr>
		<td align="right" width=100%>Excel טעינה מקובץ</td>
		<td align="center" width="65" nowrap><input type="radio" value="excel" name="adding_type">&nbsp;</td>		
	</tr>	
	<tr>
		<td align="right" width=100%>טעינה ממועדון ארגונים</td>
		<td align="center" width="65" nowrap><input type="radio" value="contacts" name="adding_type">&nbsp;</td>		
	</tr>	
	<tr>
		<td height="30" nowrap colspan="2"></td>
	</tr>	
	<tr>
		<td colspan="2" width="100%">
			<table width="100%" border="0" cellpadding="5" dir=rtl>
				<tr>
					<input type="hidden" id=submit1 name="submit1" value="ok">
					<td align="left" width=45% nowrap><INPUT class="but_menu" type="button" style="width:90px"  value="<< חזרה" onclick="document.location.href='../wizard/wizard_<%=wizard_id%>_<%=wizard_page_id%>.asp'"></td>
					<td width="10%" nowrap></td>
					<td align="right" width=45%><input type="submit" value="המשך >>" class=but_menu style="width:90px" id=submit1 name=submit1></td>
				</tr>
			</table>
		</td>		
	</tr>											
</table>
</form>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>

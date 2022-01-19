<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%	UserID = cInt(trim(Request.Cookies("bizpegasus")("UserID")))
    OrgID = cInt(trim(trim(Request.Cookies("bizpegasus")("ORGID"))))
    If Request.QueryString("groupId") <> nil Then
		groupId = cInt(trim(Request.QueryString("groupId")))
	Else
		groupId = ""	
	End If
	
    If Request.Form.Count > 0 Then 
		if trim(Request.QueryString("groupId")) <> "" then
			 sqlstr = "Update groups set groupName='"& sfix(Trim(Request.Form("groupName"))) &"' Where groupId=" &Request.QueryString("groupId")& " And ORGANIZATION_ID=" & OrgID
			 'Response.Write sqlstr
			 con.ExecuteQuery(sqlstr)
		else
			 sqlstr = "SET NOCOUNT ON;INSERT Into GROUPS (groupName,GROUPDATE,ORGANIZATION_ID,USER_ID,GROUPTYPE) "&_
			 " Values ('" & sfix(Trim(Request.Form("groupName"))) & "',getDate()," & OrgID & "," & UserID & ", 1);"&_
			 " Select @@IDENTITY AS NewID"
			 'Response.Write sqlstr
			 Set rs_tmp = con.GetRecordSet(sqlstr)	
			 If not rs_tmp.EOF Then
				groupId = trim(rs_tmp(0))
			 End If
			 Set rs_tmp = Nothing			
		end if
%>
	<SCRIPT  LANGUAGE=javascript>
	<!--
		window.opener.document.location.href = "group.asp?groupId=<%=groupId%>"
		this.close();
	//-->
	</SCRIPT>
<%	end if %>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 32 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing
	  
	  sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	  set rsbuttons = con.getRecordSet(sqlstr)
	  If not rsbuttons.eof Then
		arrButtons = ";" & rsbuttons.getString(,,",",";")		
		arrButtons = Split(arrButtons, ";")
	  End If
	  set rsbuttons=nothing	 	  
%>
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
	{	<%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס שם קבוצה "
				Else
					str_alert = "Please insert the group name !!"
				End If   
		%>
		window.alert("<%=str_alert%>");
		document.form1.groupName.focus();
		return false;}
	return true;	
}
//-->
</SCRIPT>
</head>

<body style="margin:0px;background-color:#E5E5E5" onload="window.focus()">
<table border="0" bordercolor="navy" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" dir="<%=dir_var%>">
<tr><td class="page_title">&nbsp;<%if trim(groupId)<>"" then%><!--עדכון--><%=arrTitles(1)%><%Else%><!--הוספת--><%=arrTitles(2)%><%End If%>&nbsp;<!--קבוצה--><%=arrTitles(4)%>&nbsp;</td></tr>         	
<tr><td height=15 nowrap></td></tr>
<tr><td width="100%">
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
<table align=center border="0" cellpadding="2" cellspacing="1" width="70%">
<form action="groupEdit.asp?groupId=<%=groupId%>" method=POST id=form1 name=form1 onsubmit="return checkfields();">
	<tr>
		<td align="<%=align_var%>" width=100%><input type="text" class="form" dir="<%=dir_obj_var%>" value="<%=vfix(groupName)%>" size=50 maxlength=100 name="groupName">&nbsp;</td>
		<td align="<%=align_var%>" width="80" nowrap><span id=word3 name=word3><!--שם קבוצה--><%=arrTitles(3)%></span></td>		
	</tr>	
	<tr>
		<td align="<%=align_var%>" height="30" nowrap colspan="2"></td>
	</tr>	
	<tr>
		<td colspan="2" width="100%">
			<table width="100%" border="0" cellpadding="5">
				<tr>
					<input type="hidden" id=submit1 name="submit1" value="ok">
					<td align="<%=align_var%>" width=45% nowrap><INPUT class="but_menu" type="button" style="width:90px"  value="<%=arrButtons(2)%>" id=button2 name=button2 onclick="window.close()"></td>
					<td width="10%" nowrap></td>
					<td align="left" width=45%><input type="submit" value="<%=arrButtons(1)%>" class=but_menu style="width:90px" id=Button1 name=Button1></td>
				</tr>
			</table>
		</td>		
	</tr>											
</form>
</table>
</td></tr></table>
</body>
</html>
<%Set con=Nothing%>
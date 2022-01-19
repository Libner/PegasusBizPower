<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%groupId = cLng(trim(Request("groupId")))
  pageId = trim(Request("pageId"))  
  delGroup = trim(Request.QueryString("delGroup"))
  found_group = false
  If trim(groupId) <> "" Then
	sqlstr="Select GROUPNAME, GROUPTYPE FROM GROUPS WHERE GROUPID="&groupId&" And Organization_ID = " & OrgID
	set rsn = con.getRecordSet(sqlstr)
	If not rsn.eof Then
		groupname = trim(rsn(0))
		grouptype = trim(rsn(1))
		found_group = true
	Else
		found_group = false	
	End if
	set rsn = Nothing	
  Else
	  found_group = false	
  End If
  
  if trim(delGroup) = "True" then
	
		sqlStr = "Delete from PEOPLES where ORGANIZATION_ID="&OrgID&" and GROUPID="&groupId 
		con.ExecuteQuery(sqlStr)
	   
		slqStr = "delete from groups where groupId="&groupId&" and ORGANIZATION_ID = "&OrgID
		con.ExecuteQuery(slqStr)
		
		Response.Redirect "default.asp"	
		
	end if

	PEOPLE_ID = Request.QueryString("PEOPLE_ID")
	delUser = Request.QueryString("delUser")
	if trim(delUser) = "True" then		
		sqlStr = "Delete from PEOPLES Where PEOPLE_ID="&PEOPLE_ID&" and ORGANIZATION_ID=" & OrgID  
		con.ExecuteQuery(sqlStr) 
	end if

	if trim(Request.Form("PageCurr"))<>"" then
		PageCurr=cLng(Request.Form("PageCurr"))
	else
		PageCurr=1
	end if
	 
	if trim(Request.QueryString("numOfRow"))<>"" then
		numOfRow=Request.QueryString("numOfRow")
	else
		numOfRow = 1
	end if  		
  
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 47 Order By word_id"				
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
%> 	
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
function goForms(groupId)
{
   mywindCust = window.open("groupEdit.asp?groupId=" + groupId ,'groupEdit',"alwaysRaised,left=180,top=200,height=200,width=450,status=no,toolbar=no,menubar=no,location=no");
   return false;
}
function goForms1(groupId)
{
   mywindCust = window.open("groupEdit1.asp?groupId=" + groupId ,'groupEdit',"alwaysRaised,left=180,top=200,height=200,width=450,status=no,toolbar=no,menubar=no,location=no");
   return false;
}

function goFormsClient(PEOPLE_ID,groupId,page)
{
	mywindCust = window.open("addpeople.asp?PEOPLE_ID="+ PEOPLE_ID +"&groupId=" + groupId + "&page=" + page,'addclient',"alwaysRaised,left=100,top=100,height=300,width=450,status=no,toolbar=no,menubar=no,location=no");
	return false;
}

function FailureDelete(groupId) {
   <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את הקבוצה"
     Else
		str_confirm = "Are you sure want to delete the group ?"
     End If   
   %>		
	if (confirm("<%=str_confirm%>"))		
	{
		window.document.location.href = "group.asp?groupId=" + groupId + "&delGroup=True"
		return false;
	}
	return false;	
}

function UserDelete(PEOPLE_ID,groupId,page) {
   <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק"
     Else
		str_confirm = "Are you sure want to delete ?"
     End If   
   %>		
	if (confirm("<%=str_confirm%>"))		
	{
		window.document.location.href = "group.asp?PEOPLE_ID=" + PEOPLE_ID + "&groupId=" + groupId + "&page=" + page + "&delUser=True"
		return false;
	}
	return false;	
}
//-->
</SCRIPT>

</head>
<body>
<!-- #include file="../../logo_top.asp" -->
<%numOftab = 2%>
<%numOfLink = 1%>
<!--#include file="../../top_in.asp"-->
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF dir="<%=dir_var%>">
<tr><td width="100%" class="page_title" dir=rtl>&nbsp;<font color="#6E6DA6"><%=groupName%></font></td></tr>         
<%If trim(wizard_id) <> "" Then%>
<tr><td width="100%">
<% wizard_page_id = 2 %>
<!--#include file="../../wizard_inc.asp"-->
</td></tr>
<%End If%>
<%If found_group Then%>
<tr><td width="100%" colspan=2>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
 <%if trim(groupId) <> "" and IsNumeric(groupId)="True" and delGroup <> "True" then%>				
  <tr><td align="<%=align_var%>" valign=top>
	 <table border=0 width=100% cellpadding=0 cellspacing=1 bgcolor="#FFFFFF" ID="Table1">			
	 <tr>
		<td width="45" nowrap class="title_sort" align=center><!--מחק--><%=arrTitles(3)%></td>
		<td width="80" nowrap class="title_sort" align=center>כרטיס לקוח</td>			
		<td width="100" nowrap align="<%=align_var%>" class="title_sort">&nbsp;טלפון נייד&nbsp;</td>				
		<td width="180" nowrap align="<%=align_var%>" class="title_sort">&nbsp;<!--חברה--><%=arrTitles(4)%>&nbsp;</td>
		<td width="150" nowrap align="<%=align_var%>" class="title_sort">&nbsp;<!--שם מלא--><%=arrTitles(5)%>&nbsp;</td>																					
		<td width="100%" style="line-height:18px" align="<%=align_var%>" class="title_sort">&nbsp;Email&nbsp;</td>
	 </tr>					
	 <%					
		PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
		If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
     			PageSize = 10
		End If	
		sqlStr = "SELECT PEOPLE_ID,PEOPLE_EMAIL,PEOPLE_COMPANY,PEOPLE_NAME,CONTACT_ID, PEOPLE_CELL FROM " & _
		" PEOPLES WHERE ORGANIZATION_ID=" & OrgID & " and groupId="& groupId &" ORDER BY rtrim(ltrim(PEOPLE_EMAIL))" 
		''Response.Write sqlStr
		set rs_tmp = con.GetRecordSet(sqlStr)
		If Not rs_tmp.Eof Then
			recCount = rs_tmp.RecordCount
			TotalPages = Fix((recCount / PageSize) + 0.99)			
			rstop = PageSize * (PageCurr)
			rstart = rstop - (PageSize - 1) 
			rs_tmp.move(rstart-1) 						
			If rstop > recCount - 1 then PageSize = (recCount + 1) - rstart 
			arr_tmp = rs_tmp.getRows(PageSize)			
		Else
			recCount = 0	: TotalPages = 0
		End If
		Set rs_tmp = Nothing

		If isArray(arr_tmp) Then
		 For nn = 0 To Ubound(arr_tmp, 2)
			PEOPLE_ID = trim(arr_tmp(0,nn))
			PEOPLE_EMAIL = trim(arr_tmp(1,nn))
			PEOPLE_COMPANY = trim(arr_tmp(2,nn))
			PEOPLE_NAME = trim(arr_tmp(3,nn))
			CONTACT_ID = trim(arr_tmp(4,nn))	
			PEOPLE_CELL = trim(arr_tmp(5,nn))	%>
		<tr>
			<td class="card" align=center><a href="" ONCLICK="return UserDelete(<%=PEOPLE_ID%>,<%=groupId%>,<%=PageCurr%>);"><IMG class=img_class SRC=../../images/delete_icon.gif BORDER=0 ALT="מחיקה"></a></td>						
			<td class="card" align=center><%If Len(CONTACT_ID) > 0 Then%><a href="../companies/contact.asp?contactId=<%=CONTACT_ID%>" target="_self"><img src="../../images/edit_icon.gif" border="0"></a><%End If%></td>
			<td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=PEOPLE_CELL%>&nbsp;</td>
			<td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>"><%If trim(PEOPLE_COMPANY) <> "" Then%><a href="" ONCLICK="return goFormsClient(<%=PEOPLE_ID%>,<%=groupId%>,<%=PageCurr%>);" class="link_categ">&nbsp;<%=PEOPLE_COMPANY%>&nbsp;</a><%End If%></td>
			<td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>"><%If trim(PEOPLE_NAME) <> "" Then%><a href="" ONCLICK="return goFormsClient(<%=PEOPLE_ID%>,<%=groupId%>,<%=PageCurr%>);" class="link_categ">&nbsp;<%=PEOPLE_NAME%>&nbsp;</a><%End If%></td>
			<td class="card" align="<%=align_var%>" dir="ltr"><a href="" ONCLICK="return goFormsClient(<%=PEOPLE_ID%>,<%=groupId%>,<%=PageCurr%>);" class="link_categ">&nbsp;<%=PEOPLE_EMAIL%>&nbsp;</a></td>
		</tr>					
	<%Next%>
	<%If cLng(recCount) > cLng(PageSize) Then%>
	<tr bgcolor="#e6e6e6">
	<td align="center" colspan="6" width=100%>	
		<form id="paging_form" name="paging_form" action="group.asp?groupId=<%=groupId%>" method="Post" style="margin:0px">
		<table border="0" cellspacing="2" cellpadding="2" ID="Table2">
			<tr>			  
			<td align="center" class="card">
			<select id="PageCurr" name="PageCurr" style="color:#6E6DA6;font-size:9pt;font-weight:bold" onchange="window.document.paging_form.submit()">
			<%For pp=1 To TotalPages%>
			<option value="<%=pp%>" <%If trim(PageCurr) = trim(pp) Then%> selected <%End If%>><%=pp%></option>  
			<%Next%>	
			</select>
			</td>	
			<td align="center" class="card" dir="rtl"><b>&nbsp;עמוד&nbsp;<span style="color:#6E6DA6;font-size:9pt;font-weight:bold"><%=PageCurr%></span>&nbsp;מתוך&nbsp;<span style="color:#6E6DA6;font-size:9pt;"><%=TotalPages%></span></b></td>
			</tr>
		</table>
		</form>		
		</td>
	</tr>	
	<%End If%>
	<tr>
		<td colspan="6" height=18 bgcolor="#e6e6e6" align=center class="card" dir="ltr" style="color:#6E6DA6;font-weight:600"><span id=word6 name=word6><!--נמצאו--><%=arrTitles(6)%></span> <%=recCount%> <span id="word7" name=word7><!--נמענים--><%=arrTitles(7)%></span></td>
	</tr>
	<% Else %>
	<tr><td colspan="6"  bgcolor="#e6e6e6" align=center style="font-size=12px;font-weight=550;color=red;" dir="rtl"><span id="word8" name=word8><!--לא נמצאו נמענים--><%=arrTitles(8)%></span></td></tr>
	<% End If%>
	</table></td>
	<%End If%>		
		<td width=125 nowrap align="<%=align_var%>" valign=top class="td_menu">
		<table cellpadding=1 cellspacing=1 width=100% ID="Table3">		
		<tr><td align="right" colspan=2 height="17" nowrap></td></tr>
		<%If trim(grouptype) = "1" Then%>
		<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:115;"  href='#' ONCLICK="window.open('addpeople.asp?groupId=<%=groupId%>',null,'alwaysRaised,left=100,top=160,height=320,width=450,status=no,toolbar=no,menubar=no,location=no');"><!--הוסף נמען--><%=arrTitles(11)%></a></td></tr>
		<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:115;"  href='#' ONCLICK="window.open('excelUpload.asp?groupId=<%=groupId%>',null,'alwaysRaised,left=180,top=200,height=200,width=450,status=no,toolbar=no,menubar=no,location=no');"><!--(Excel) הוסף נמענים--><%=arrTitles(12)%></a></td></tr>
		<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:115;line-height:110%;padding:3px"  href='#'  ONCLICK="window.open('getContacts.asp?groupId=<%=groupId%>',null,'alwaysRaised,left=180,top=50,height=450,width=450,status=no,toolbar=no,menubar=no,location=no,scrollbars=yes');"><!--הוסף נמענים מקשרי לקוחות--><%=arrTitles(13)%></a></td></tr>
		<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:115;" href='#' ONCLICK="window.open('excelUploadBlocked.asp?groupId=<%=groupId%>',null,'alwaysRaised,left=100,top=200,height=250,width=500,status=no,toolbar=no,menubar=no,location=no');">(Excel) מחק נמענים</a></td></tr>
		<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:115;"  href='#' ONCLICK="window.open('excelDownload.asp?groupId=<%=groupId%>',null,'alwaysRaised,left=180,top=200,height=200,width=450,status=no,toolbar=no,menubar=no,location=no');"><!--(Excel) ייצוא נמענים--><%=arrTitles(14)%></a></td></tr>
		<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:115;"  href='#' ONCLICK="return goForms(<%=groupId%>);"><span id="word10" name=word10><!--עדכן פרטי קבוצה--><%=arrTitles(10)%></span></a></td></tr>
		<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:115;"  href='#' ONCLICK="return FailureDelete(<%=groupId%>);"><span id=word9 name=word9><!--מחק קבוצה--><%=arrTitles(9)%></span></a></td></tr>
		<%Else%>
		<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:115;line-height:110%;padding:3px"  href='#'  ONCLICK="window.open('getContacts1.asp?groupId=<%=groupId%>',null,'alwaysRaised,left=180,top=50,height=450,width=450,status=no,toolbar=no,menubar=no,location=no,scrollbars=yes');"><!--הוסף נמענים מקשרי לקוחות--><%=arrTitles(13)%></a></td></tr>
		<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:115;"  href='#' ONCLICK="window.open('excelDownload.asp?groupId=<%=groupId%>',null,'alwaysRaised,left=180,top=200,height=200,width=450,status=no,toolbar=no,menubar=no,location=no');"><!--(Excel) ייצוא נמענים--><%=arrTitles(14)%></a></td></tr>
		<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:115;"  href='#' ONCLICK="return goForms1(<%=groupId%>);"><!--עדכן פרטי קבוצה--><%=arrTitles(10)%></a></td></tr>
		<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:115;"  href='#' ONCLICK="return FailureDelete(<%=groupId%>);"><!--מחק קבוצה--><%=arrTitles(9)%></a></td></tr>		
		<%End If%>
		<tr><td align="right" colspan=2 height="5" nowrap></td></tr>
		</table></td></tr>
		</table></td></tr>
		<tr><td height=10 nowrap></td></tr>
		</table>
	</tr>		
</table>
</td></tr></table>
</td></tr>
<%End If%>
</table>
</body>
<%set con=nothing%>
</html>

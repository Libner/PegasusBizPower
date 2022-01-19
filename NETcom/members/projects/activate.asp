<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--
function CheckFields()
{	
	if ( document.frmMain.dateStart.value=='' )		   
		{
				alert("אנא הכנס תאריך פתיחה")
				return false;				
		}
	else
		{
			document.frmMain.submit();
			return true;
		}		
}

function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}

//-->
</script>
<%
UserID=trim(Request.Cookies("bizpegasus")("UserID"))
OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))

project_ID=request.querystring("project_ID")
company_id=Request.Form("company_id")
errName = false

if request.form("dateStart")<>nil then 'after form filling	
	
	   dateStart = "'" & Request.Form("dateStart") & "'"	

		sqlStr = "select project_name from projects where project_name = '" & sFix(project_name) & "' and project_id <> " & project_id & " AND ORGANIZATION_ID = " & OrgID
		set rs_name = con.GetRecordSet(sqlStr)
		if  not rs_name.EOF then
		    errName=true
		end if
		set rs_name = nothing			
		'Response.Write status
		'Response.End
		if errName=false then
			con.executeQuery("SET DATEFORMAT dmy")
			sqlStr = "Update projects set start_date = " & dateStart &_
			", status = '2', active = '1' where project_ID=" & project_ID
			con.GetRecordSet (sqlStr)
		If trim(wizard_id) = "" Then	
		%>
		<%if trim(company_id)="" or 	company_id=0 then%>
		 <SCRIPT LANGUAGE=javascript>
			<!--
			window.close();
		    window.opener.document.location.href = "default_action.asp" 
			//-->
		</SCRIPT>
		 <%else %>
		  <SCRIPT LANGUAGE=javascript>
			<!--
			window.close();
		    window.opener.document.location.href = "default.asp" 
			//-->
		</SCRIPT>
		 
		 <%end if%>
	   	<%Else%>
		<SCRIPT LANGUAGE=javascript>
			<!--
			window.close();
			window.opener.document.location.href = "../wizard/wizard_<%=wizard_id%>_4.asp?wPricing=true";
			//-->
		</SCRIPT>		
		<%	
		End If	     
		elseif errName=true then
		project_name = ""%>
		<SCRIPT LANGUAGE=javascript>
			<!--
			alert('שם פרויקט זה בשימוש, בחר שם פרויקט אחר')
			history.back();
			//-->
		</SCRIPT>				 			 				
	<%
    end if        
end if
%>
<%
if project_ID<>nil and project_ID<>"" then
  sqlStr = "select company_id, project_name, project_description, start_date, end_date from projects where project_ID=" & project_ID  
  ''Response.Write sqlStr
  set rs_projects = con.GetRecordSet(sqlStr)
	if not rs_projects.eof then
		company_id = rs_projects("company_id")
		project_name = rs_projects("project_name")			
	end if
	set rs_projects = nothing	
    dateStart = FormatDateTime(Date(),2)
end if
'Response.Write company_id
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 68 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(trim(arr_title(0))) = arr_title(1)
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
<body style="margin:0px" onload="window.focus();">
<table border="0" cellpadding="0" cellspacing="1" width="100%" dir="<%=dir_var%>">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0" dir="<%=dir_var%>">
	  <tr><td class="page_title" dir=rtl>&nbsp;<span name="word1" id="word1"><!--הפעלת--><%=arrTitles(1)%></span>&nbsp;<font color="#6F6DA6"><%=project_name%></font></td></tr>		   
    </table>
</td></tr> 
<tr><td width=100% bgcolor="#E6E6E6"> 
<table align=center border="0" cellpadding="3" cellspacing="1" width="70%">
<tr><td height=10></td></tr>          
<FORM name="frmMain" ACTION="activate.asp?project_ID=<%=project_ID%>" METHOD="post" onSubmit="return CheckFields()">
<input id=company_id name=company_id value=<%=company_id%> type=hidden>
<tr>
	<td align="<%=align_var%>">	
	<input type=image src='../../images/calend.gif' border=0  onclick='return popupcal(this.form.elements("dateStart"));'>&nbsp;<input dir="ltr" type="text" class="passw" style="width:85" id="dateStart" name="dateStart" value="<%=dateStart%>" maxlength=10 onclick="return popupcal(this);" readonly>
	</td>
	<td align="<%=align_var%>" nowrap width="100">&nbsp;<span name="word2" id="word2"><!--תאריך פתיחה--><%=arrTitles(2)%></span>&nbsp;</td>
</tr>
<tr><td colspan="2" height=10></td></tr>
<tr>
<td colspan=2>
<table cellpadding=0 cellspacing=0 align=center dir="<%=dir_var%>" width=90%>
<tr>
<td width=50% align=center><input type=button class="but_menu" style="width:90px" onclick="window.close();" value="<%=arrButtons(2)%>" id=Button2 name=Button2></td>
<td width=50% align=center><input type=button class="but_menu" style="width:90px" onclick="javascript:CheckFields(); return false;" value="<%=arrButtons(1)%>" id=Button1 name=Button1></td></tr>
</table></td></tr>
</form>
<tr><td colspan="2" height=10></td></tr>
</table>
</td></tr></table>
</body>
<%
set con = nothing
%>
</html>

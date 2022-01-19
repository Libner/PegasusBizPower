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
	if ( document.frmMain.project_name.value=='' ||
		 document.frmMain.project_code.value=='' ||
	     document.frmMain.company_id.value=='' && document.frmMain.project_type(1).checked == true
	     )		   
		{
				alert("'*' אנא מלאו כל השדות הצוינות בסימן")
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

project_ID=request.querystring("project_ID")
pricing=request.querystring("pricing") ' בנו להוסיף פרויקט עתידי
errName = false

if request.form("project_name")<>nil then 'after form filling
	 project_code = sFix(request.form("project_code"))
	 project_name = sFix(request.form("project_name"))
	 project_description  = sFix(trim(Request.Form("project_description")))
	 company_id = trim(Request.Form("company_id"))
	 If trim(company_id) = "" Then
		company_id = 0
	 End If
	 If Request.Form("dateStart") <> nil Then
		 dateStart = "'" & Request.Form("dateStart") & "'"
	 Else
		 dateStart = "NULL"
	 End If		 
	 If Request.Form("dateEnd") <> nil Then
		 dateEnd = "'" & Request.Form("dateEnd") & "'"
	 Else
		 dateEnd = "NULL"
	 End If
	 If IsDate(Request.Form("dateStart")) = false Or Request.Form("dateStart") = nil Then
		  status = "1"		 
	 Else
		  status = "2"			
	 End if	
	 
	 if project_ID=nil or project_ID="" then 'new record in DataBase
		   
		con.executeQuery("SET DATEFORMAT dmy")
		sqlStr = "insert into projects (company_id,project_code,project_name,project_description,start_date,end_date,status,organization_id,active) "
		sqlStr=sqlStr& " values (" & company_id &",'"& project_code & "','" & project_name &"','"& project_description &"'," & dateStart & "," & dateEnd & ",'" & status & "'," & OrgID & ",1)"
		'Response.Write sqlStr
		con.GetRecordSet(sqlStr)					
		If isNumeric(wizard_id) Then
			Response.Redirect ("../wizard/wizard_"&wizard_id&"_4.asp?wProjects=true") 				
		ElseIf not trim(pricing) = "1" Then
			Response.Redirect ("default.asp")          
		Else
			Response.Redirect ("prices.asp")          	
		End If			          
			 
		    
	 else			   
			
			If IsDate(Request.Form("dateEnd")) Then
				status = "3"
			Else
				status = "2"		 
			End if	
			If IsDate(Request.Form("dateStart")) = false Or Request.Form("dateStart") = nil Then
				status = "1"		 
			End if	
			'Response.Write status
			'Response.End			
			con.executeQuery("SET DATEFORMAT dmy")
			sqlStr = "Update projects set project_name = '" & project_name &_
			"', project_code = '" & project_code &_
			"', project_description = '" & project_description &_
			"', start_date = " & dateStart & ", end_date = " & dateEnd &_
			", status ='" & status & "' where project_ID=" & project_ID
			con.GetRecordSet (sqlStr)
			If isNumeric(wizard_id) Then
				Response.Redirect ("../wizard/wizard_"&wizard_id&"_4.asp?wProjects=true") 				
			ElseIf not trim(pricing) = "1" Then
				Response.Redirect ("default.asp")          
			Else
				Response.Redirect ("prices.asp")          	
			End If			
    end if        
end if
%>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 0%>
<%If not trim(pricing) = "1" Then%>
<%numOfLink = 5%>
<%Else%>
<%numOfLink = 6%>
<%End If%>
<%If trim(wizard_id) = "" Then%>
<!--#include file="../../top_in.asp"-->
<%End If%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0" ID="Table8">
	  <tr><td class="page_title"><%If IsNumeric(wizard_id) Then%> <%=page_title_%> <%Else%>&nbsp;<%if project_ID<>nil then%>עדכון<%else%>הוספת<%end if%>&nbsp;פרויקט&nbsp;<%End If%></td></tr>		   
	  		       	
   </table></td></tr>  
<% wizard_page_id = 3 %>
<tr><td width=100% align="right">
<!--#include file="../../wizard_inc.asp"-->
</td></tr>
<%If trim(wizard_id) <> "" Then%>
<tr>
<td width=100% align="right" bgcolor="#FFD011" style="padding:5px">
<table border=0 cellpadding=0 cellspacing=0 width=100% ID="Table1">
<tr><td class="explain_title">צור טופס רישום חדש!</td></tr>
<tr><td height=5 nowrap></td></tr>
<tr><td class=explain>
בשדה <b>שפת הטופס</b> , בחר את השפה בה ייכתב טופס ההרשמה.
</td></tr>
<tr><td class=explain>
בשדה <b>שם הטופס</b> הזן שם לטופס הרישום. שם זה לא יופיע בדיוור אלא ישמש לזיהוי הטופס במערכת ה-Bizpower.
</td></tr>
<tr><td class=explain>
בשדה <b>טקסט מקדים</b> , הקלד את הטקסט שיופיע בתחילת הטופס לפני שדות הרישום.
</td></tr>
<tr><td class=explain>
בשדה <b>טקסט תודה לאחר הרישום</b> , הקלד את הטקסט שברצונך שהנרשם יראה לאחר שליחת הטופס. <br>לדוגמה: "תודה! נציגינו יחזרו אליך ב-24 השעות הקרובות"
</td></tr>
</table>
</td>
</tr>
<tr>
    <td width="100%" height="2"></td>
</tr>

<%End If%>        
<tr><td width=100% bgcolor="#E6E6E6"> 
<table align=center border="0" cellpadding="3" cellspacing="1" width="500">    
<%
if project_ID<>nil and project_ID<>"" then
  sqlStr = "select company_id, project_code, project_name, project_description, start_date, end_date from projects where project_ID=" & project_ID  
  ''Response.Write sqlStr
  set rs_projects = con.GetRecordSet(sqlStr)
	if not rs_projects.eof then
		company_id = rs_projects("company_id")
		project_code = rs_projects("project_code")
		project_name = rs_projects("project_name")
		project_description = rs_projects("project_description")
		dateStart = rs_projects("start_date")
		dateEnd = rs_projects("end_date")	
	end if
	set rs_projects = nothing
	
	sqlstr = "Select TOP 1 project_id FROM hours WHERE organization_id = " & trim(OrgID)
	set rs_h = con.getRecordSet(sqlstr)
	if not rs_h.eof then
		is_worked = true
	else
		is_worked = false
	end if
	set rs_h = nothing
else
	dateStart = Day(date()) & "/" & Month(date()) & "/" & Year(date())
end if
'Response.Write company_id
%>
<FORM name="frmMain" ACTION="addProject.asp?project_ID=<%=project_ID%>&pricing=<%=pricing%>" METHOD="post" onSubmit="return CheckFields()">
<tr>
	<td align=right><input type=text dir="rtl" name="project_code" value="<%=vfix(project_code)%>" style="width:100px;font-family:arial" ID="project_code" maxlength=50></td>
	<td align="right" nowrap width="100">קוד פרויקט / הזמנה&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align=right><input type=text dir="rtl" name="project_name" value="<%=vfix(project_name)%>" style="width:300px;font-family:arial"></td>
	<td align="right" nowrap width="100">שם פרויקט&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align=right><textarea dir=rtl name="project_description" style="width:300px;font-family:arial" ID="Textarea1"><%=vfix(project_description)%></textarea></td>
	<td align="right" nowrap width="100">תיאור פרויקט&nbsp;</td>
</tr>
<tr>
	<td align="right" dir=rtl>פעילות כללית&nbsp;<font color="#736BA6">(פעילות לכלל הלקוחות)</font></td>
	<td align=left nowrap width="100"><input <%If is_worked Then%> disabled <%End if%> type=radio dir="rtl" name="project_type" <%If trim(company_id) = "" Or trim(company_id) = "0" Then%> checked <%End If%>  onclick="client_body.style.display='none';company_id.disabled=true;" value=0></td>
</tr>
<tr>
	<td align="right" dir=rtl>פרויקט לקוח&nbsp;<font color="#736BA6">(פרויקט שרלונטי ללקוח ספציפי)</font></td>
	<td align=left nowrap width="100"><input <%If is_worked Then%> disabled <%End if%> type=radio dir="rtl" name="project_type" <%If trim(company_id) <> "" And trim(company_id) <> "0" Then%> checked <%End If%> onclick="client_body.style.display='inline';company_id.disabled=false;" value=1></td>	
</tr>
<tbody name="client_body" id="client_body" <%If trim(company_id) = ""  Or trim(company_id) = "0" Then%> style="display:none" <%end if%>>
<tr>
	<td align=right>
	<select dir="rtl" name="company_id" style="width:300px;font-family:arial" <%If is_worked Then%> disabled <%End if%>>	
	<option value=""><%=String(20,"-")%> בחר לקוח <%=String(20,"-")%></option>
	<%  sqlstr = "Select COMPANY_ID, COMPANY_NAME FROM COMPANIES WHERE ORGANIZATION_ID = " & OrgID & " Order BY PRIVATE_FLAG DESC,COMPANY_NAME"
		set rs_comp = con.getRecordSet(sqlstr)
		While not rs_comp.eof
	%>
	<option value="<%=rs_comp(0)%>" <%If trim(company_id) = trim(rs_comp(0)) Then%> selected <%End If%>><%=rs_comp(1)%></option>
	<%
		rs_comp.moveNext
		Wend
		set rs_comp = Nothing
	%>
	</select></td>
	<td align="right" nowrap width="100">לקוח&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
</tbody>
<%If not trim(pricing) = "1" Then%>
<tr>
	<td align=right>	
	<input type=image src='../../images/calend.gif' border=0  onclick='return popupcal(this.form.elements("dateStart"));'>&nbsp;<input dir="ltr" type="text" class="passw" style="width:85" id="dateStart" name="dateStart" value="<%=dateStart%>" maxlength=10 onclick="return popupcal(this);" readonly>
	</td>
	<td align="right" nowrap width="100">&nbsp;תאריך פתיחה&nbsp;</td>
</tr>
<%End If%>
<tr><td colspan="2" height=10></td></tr>
<%If trim(wizard_id) <> "" Then%>	
<tr>
<td colspan=2>
<table cellpadding=0 cellspacing=0 align=center>
<tr>	
	
		<td width=50% align="right" nowrap>
		<input style="width:90px" class="but_menu" type=button value="<< המשך" onclick="javascript:CheckFields(); return false;">
		</td>
		<td width=50 align="right" nowrap></td>
		<td width=50% align="left" nowrap>
		<input type=button class="but_menu" value="חזור >>" onclick="document.location.href='../wizard/wizard_<%=wizard_id%>_<%=wizard_page_id%>.asp'" style="width:90px">
		</td>		
	</tr>
</table></td></tr>		
<%Else%>	
<tr>
<td colspan=2>
<table cellpadding=0 cellspacing=0 align=center>
<tr>
<td width=50% align=right><input type=button class="but_menu" style="width:90px" onclick="document.location.href='<%If trim(pricing) = "1" Then%>prices.asp<%Else%>default.asp<%End If%>';" value="ביטול"></td>
<td width=50 nowrap></td>
<td width=50% align=left><input type=button class="but_menu" style="width:90px" onclick="javascript:CheckFields(); return false;" value="אישור"></td></tr>
</table></td></tr>
<%End If%>
</form>
<tr><td colspan="2" height=10></td></tr>
</table>
</td></tr></table>
</body>
<%
set con = nothing
%>
</html>

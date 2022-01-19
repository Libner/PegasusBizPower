<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--INCLUDE FILE="../checkWorker.asp"-->
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
	if ( document.frmMain.project_name.value==''
	//||	     document.frmMain.company_id.value=='' && document.frmMain.project_type(1).checked == true
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

	
	function closeWin()
	{
		opener.focus();
		opener.window.location.reload();
		self.close();
	}
<!--End-->
</script>  

<%

UserID=trim(Request.Cookies("bizpegasus")("UserID"))
OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
project_ID=Request.querystring("project_ID")
company_id=Request.querystring("companyID")

sqlstr = "Select PRODUCT_ID FROM PRODUCTS WHERE PRODUCT_TYPE = 1 AND ORGANIZATION_ID=" & trim(OrgID) ' טופס של פרויקט
set rs_pr = con.getRecordSet(sqlstr)
If not rs_pr.eof Then
	prod_id = trim(rs_pr(0))
End If
set rs_pr = nothing

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
		sqlStr=sqlStr& " values (" & company_id &",'"& project_code & "','" & project_name &"','"& project_description &"'," & dateStart & "," & dateEnd & ",'" & status & "'," & OrgID & ",0)"
		'Response.Write sqlStr
		con.GetRecordSet(sqlStr)
		
		set tmp = con.GetRecordSet("Select TOP 1 project_id From projects order by project_id desc")
		if not tmp.eof then
			project_ID = tmp(0)			
		end if
		tmp.close()
		set tmp = nothing		
		
		sqlstr="SELECT Field_Id,FIELD_TYPE  FROM FORM_FIELD Where product_id=" & prod_id & " Order by Field_Order"
		'Response.Write sqlstr
		'Response.End
		set fields=con.GetRecordSet(sqlstr)
			do while not fields.EOF
				Field_Id = fields("Field_Id")				
				Field_Value = trim(Request("field" & Field_Id))		
				'Response.Write Field_Id & " " & Field_Value
				'Response.End						
				sqlstr="INSERT INTO FORM_PROJECT_VALUE (project_ID,Field_Id,Field_Value) VALUES (" & project_ID & "," & Field_Id & ",'"& sFix(Field_Value) &"')"
				'Response.Write(sqlstr) & "<br><br>"
				'Response.End
				con.ExecuteQuery sqlstr				
			fields.moveNext()
			loop
		set fields=nothing
			
		%>				
		<script language=javascript>
		<!--
			closeWin();
		//-->
		</script>		          
		<%	 
		    
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
			
			sqlstr = "DELETE FROM FORM_PROJECT_VALUE Where project_ID=" & project_ID
			con.executeQuery(sqlstr)
			
			sqlstr="SELECT Field_Id,FIELD_TYPE  FROM FORM_FIELD Where product_id=" & prod_id & " Order by Field_Order"
			set fields=con.GetRecordSet(sqlstr)
				do while not fields.EOF
					Field_Id = fields("Field_Id")				
					Field_Value = trim(Request("field" & Field_Id))						
					
				    sqlstr="INSERT INTO FORM_PROJECT_VALUE (project_ID,Field_Id,Field_Value) VALUES (" & project_ID & "," & Field_Id & ",'"& sFix(Field_Value) &"')"
					'Response.Write(sqlstr)
					con.ExecuteQuery sqlstr	
					
				fields.moveNext()
				loop
			set fields=nothing
			%>
			<script language=javascript>
			<!--
			closeWin();
			//-->
			</script>
			<%	
    end if        
end if
%>
<%
if project_ID<>nil and project_ID<>"" then
  sqlStr = "select company_id, project_code, project_name, project_description, start_date, end_date, status from projects where project_ID=" & project_ID  
  ''Response.Write sqlStr
  set rs_projects = con.GetRecordSet(sqlStr)
	if not rs_projects.eof then
		company_id = rs_projects("company_id")
		'If trim(rs_projects("company_id")) <> "" And trim(rs_projects("company_id")) <> "0" Then		
			project_title = trim(Request.Cookies("bizpegasus")("Projectone"))
		'Else
			'project_title = "פעילות כללית"	
		'End If
		project_code = rs_projects("project_code")
		project_name = rs_projects("project_name")
		project_description = rs_projects("project_description")
		dateStart = rs_projects("start_date")
		dateEnd = rs_projects("end_date")
		status = rs_projects("status")	
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
	project_title =trim(Request.Cookies("bizpegasus")("Projectone"))	
end if
'Response.Write company_id
%>
<body style="margin:0px;background-color:#E6E6E6" onload="self.focus()">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title">&nbsp;<%if project_ID<>nil then%>עדכון<%else%>הוספת<%end if%>&nbsp;<%=project_title%>&nbsp;</td></tr>		   	  		       	
   </table>
</td></tr>  
<tr><td width=100% bgcolor="#E6E6E6" style="padding-right:15px"> 
<table align=center border="0" cellpadding="3" cellspacing="1" width="100%">    
<FORM name="frmMain" ACTION="editProject.asp?project_ID=<%=project_ID%>" METHOD="post" onSubmit="return CheckFields()">
<tr>
	<td align=right><input type=text dir="rtl" name="project_name" value="<%=vfix(project_name)%>" style="width:300px;font-family:arial"></td>
	<td align="right" nowrap width="100">שם <%=project_title%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align=right><textarea dir=rtl name="project_description" style="width:300px;font-family:arial" ID="Textarea1"><%=vfix(project_description)%></textarea></td>
	<td align="right" nowrap width="100" valign=top>תיאור <%=project_title%>&nbsp;</td>
</tr>
<!-- start fields dynamics -->	
<%If project_ID <> nil Then%>	
<!--#INCLUDE FILE="upd_fields.asp"-->	
<%Else%>
<!--#INCLUDE FILE="chform_fields.asp"-->	
<%end if%>
<!-- end fields dynamics -->	
<% 'if project_ID=nil or project_ID="" then %>
<!--tr>
	<td align="right" dir=rtl>פעילות כללית&nbsp;<font color="#736BA6">(פעילות לכלל הלקוחות)</font></td>
	<td align=left nowrap width="100"><input <%If is_worked Then%> disabled <%End if%> type=radio dir="rtl" name="project_type" <%If trim(company_id) = "" Or trim(company_id) = "0" Then%> checked <%End If%>  onclick="client_body.style.display='none';company_id.disabled=true;" value=0 ID="Radio1"></td>
</tr>
<tr>
	<td align="right" dir=rtl>פרויקט לקוח&nbsp;<font color="#736BA6">(פרויקט שרלונטי ללקוח ספציפי)</font></td>
	<td align=left nowrap width="100"><input <%If is_worked Then%> disabled <%End if%> type=radio dir="rtl" name="project_type" <%If trim(company_id) <> "" And trim(company_id) <> "0" Then%> checked <%End If%> onclick="client_body.style.display='inline';company_id.disabled=false;" value=1 ID="Radio2"></td>	
</tr-->
<%'End If%>
<tbody name="client_body" id="client_body" ></tbody>
<tr>
	<td align=right>
	<select dir="rtl" name="company_id" style="width:300px;font-family:arial" <%If is_worked Then%> disabled <%End if%> ID="Select1">	
	<option value=""><%=String(20,"-")%> בחר <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%> <%=String(20,"-")%></option>
	<%  sqlstr = "Select COMPANY_ID, COMPANY_NAME FROM COMPANIES WHERE ORGANIZATION_ID = " & OrgID & " Order BY COMPANY_NAME"
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
	<td align="right" nowrap width="100"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
</tbody>
<tr>
	<td align=right><input type=text dir="rtl" name="project_code" value="<%=vfix(project_code)%>" style="width:100px;font-family:arial" ID="project_code" maxlength=50></td>
	<td align="right" nowrap width="100">קוד <%=project_title%> / הזמנה&nbsp;</td>
</tr>
<% if project_ID=nil or project_ID="" then %>
<tr>
	<td align="right" dir=rtl><%=project_title%> עתידי&nbsp;<font color="#736BA6"></font></td>
	<td align=left nowrap width="100"><input type=radio dir="rtl" name="project_status" <%If trim(project_ID) = "" Then%> checked <%End if%> onclick="date_body.style.display='none';dateStart.disabled=true;" value=0></td>
</tr>
<tr>
	<td align="right" dir=rtl><%=project_title%> בביצוע&nbsp;<font color="#736BA6"></font></td>
	<td align=left nowrap width="100"><input type=radio dir="rtl" name="project_status" onclick="date_body.style.display='inline';dateStart.disabled=false;" value=1></td>	
</tr>
<%End If%>
<tbody name="date_body" id="date_body" <%If trim(dateStart) = ""  Or cInt(status) = 1 Or trim(project_ID) = "" Then%> style="display:none" <%end if%>>
<tr>
	<td align=right>	
	<input type=image src='../../images/calend.gif' border=0  onclick='return popupcal(this.form.elements("dateStart"));'>&nbsp;<input dir="ltr" type="text" class="passw" style="width:85" id="dateStart" name="dateStart" value="<%=dateStart%>" <%If trim(project_ID) = "" Then%> disabled <%End If%>  maxlength=10 onclick="return popupcal(this);" readonly>
	</td>
	<td align="right" nowrap width="100">&nbsp;תאריך פתיחה&nbsp;</td>
</tr>
</tbody>
<%If trim(status) = "3" Then%>
<tr>
	<td align=right>	
	<input type=image src='../../images/calend.gif' border=0  onclick='return popupcal(this.form.elements("dateEnd"));'>&nbsp;<input dir="ltr" type="text" class="passw" style="width:85" id="dateEnd" name="dateEnd" value="<%=dateEnd%>" maxlength=10 onclick="return popupcal(this);" readonly>
	</td>
	<td align="right" nowrap width="100">&nbsp;תאריך סגירה&nbsp;</td>
</tr>
<%End If%>
<tr><td colspan="2" height=10></td></tr>
<tr>
<td colspan=2>
<table cellpadding=0 cellspacing=0 align=center>
<tr>
<td width=50% align=right><input type=button class="but_menu" style="width:90px" onclick="closeWin()" value="ביטול"></td>
<td width=50 nowrap></td>
<td width=50% align=left><input type=button class="but_menu" style="width:90px" onclick="javascript:CheckFields(); return false;" value="אישור"></td></tr>
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

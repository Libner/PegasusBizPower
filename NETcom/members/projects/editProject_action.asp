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
	if (window.document.all("project_name").value=='')
	{
 		<%
			If trim(lang_id) = "1" Then
				str_alert = "נא למלא שם"
			Else
				str_alert = "Please, insert the name"
			End If	
		%> 		
 		window.alert("<%=str_alert%>"); 		
 		window.document.all("project_name").focus();
		return false;
	}
	if (window.document.all("project_description").value=='')
	{
 		<%
			If trim(lang_id) = "1" Then
				str_alert = "נא למלא תאור"
			Else
				str_alert = "Please, insert the description"
			End If	
		%> 		
 		window.alert("<%=str_alert%>"); 
 		window.document.all("project_description").focus();
		return false;
	}
	if (click_fun() == false)
		return false;
	document.frmMain.submit();
	return true;
	
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
//-->
</script>  

<%
'project_type = Request.querystring("project_type")
'response.Write (project_type)
'Response.end
UserID=trim(Request.Cookies("bizpegasus")("UserID"))
OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
project_ID=Request.querystring("project_ID")
company_id=0

sqlstr = "Select PRODUCT_ID, LANGU FROM PRODUCTS WHERE PRODUCT_TYPE = 1 AND ORGANIZATION_ID=" & trim(OrgID) ' טופס של פרויקט
set rs_pr = con.getRecordSet(sqlstr)
If not rs_pr.eof Then
	prod_id = trim(rs_pr(0))
	langu = trim(rs_pr(1))
End If
set rs_pr = nothing

if request.form("project_name")<>nil then 'after form filling
	 project_code = sFix(request.form("project_code"))
	 project_name = sFix(request.form("project_name"))
	 project_description  = sFix(trim(Request.Form("project_description")))
	
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
		
		project_title = trim(Request.Cookies("bizpegasus")("ActivitiesOne"))
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
	project_title = trim(Request.Cookies("bizpegasus")("ActivitiesOne"))
end if
'Response.Write company_id
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 15 Order By word_id"				
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
<body style="margin:0px;background-color:#E6E6E6" onload="self.focus()">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title">&nbsp;<%if project_ID<>nil then%><span id=word1 name=word1><!--עדכן--><%=arrTitles(1)%></span><%else%><span id=word2 name=word2><!--הוסף--><%=arrTitles(2)%></span><%end if%>&nbsp;<%=project_title%>&nbsp;</td></tr>		   	  		       	
   </table>
</td></tr>  
<tr><td width=100% bgcolor="#E6E6E6"> 
<FORM name="frmMain" ACTION="editProject_action.asp?project_ID=<%=project_ID%>" METHOD="post" onSubmit="return CheckFields()">
<TABLE WIDTH=99% BORDER=0 CELLSPACING=1 CELLPADDING=3 bgcolor=white align=center dir="<%=dir_var%>">
<tr><td align="<%=align_var%>" class="form paddingr" bgcolor=#DADADA nowrap><b><!--שם--><%=arrTitles(4)%></b></td></tr>
<tr><td align="<%=align_var%>" class="form paddingr" bgcolor=#f0f0f0><input type=text dir="<%=dir_obj_var%>" name="project_name" value="<%=vFix(project_name)%>" style="width:300px;font-family:arial" ID="project_name"></td></tr>
<tr>
	<td align="<%=align_var%>" class="form paddingr" bgcolor=#DADADA nowrap valign=top><b><!--תיאור--><%=arrTitles(5)%></b></td>
</tr>
<tr>
	<td align="<%=align_var%>" class="form paddingr" bgcolor=#f0f0f0><textarea dir="<%=dir_obj_var%>" name="project_description" style="width:300px;font-family:arial" ID="Textarea1"><%=vFix(project_description)%></textarea></td>
</tr>
<tr>	
	<td align="<%=align_var%>" class="form paddingr" bgcolor=#DADADA nowrap><b><!--קוד--><%=arrTitles(6)%></b></td>
</tr>
<tr>
	<td align="<%=align_var%>" class="form paddingr" bgcolor=#f0f0f0><input type=text dir="<%=dir_obj_var%>" name="project_code" value="<%=vFix(project_code)%>" style="width:100px;font-family:arial" ID="project_code" maxlength=50></td>
</tr>
<% if project_ID=nil or project_ID="" then %>
<tr>
	<td align="<%=align_var%>" dir="<%=dir_var%>" bgcolor=#f0f0f0 dir="<%=dir_obj_var%>"><b><!--עתידי--><%=arrTitles(7)%></b><font color="#736BA6"></font>
	&nbsp;
	<input type=radio dir="<%=dir_obj_var%>" name="project_status" <%If trim(project_ID) = "" Then%> checked <%End if%> onclick="date_body.style.display='none';dateStart.disabled=true;" value=0 ID="Radio1"></td>
</tr>
<tr>
	<td align="<%=align_var%>" dir="<%=dir_var%>" width=100% bgcolor=#f0f0f0 dir="<%=dir_obj_var%>"><b><!--בביצוע--><%=arrTitles(8)%></b><font color="#736BA6"></font>
	&nbsp;
	<input type=radio dir="<%=dir_obj_var%>" name="project_status" onclick="date_body.style.display='inline';dateStart.disabled=false;" value=1 ID="Radio2"></td>	
</tr>
<%End If%>
<tbody name="date_body" id="date_body" <%If trim(dateStart) = ""  Or cInt(status) = 1 Or trim(project_ID) = "" Then%> style="display:none" <%end if%>>
<tr>	
	<td align="<%=align_var%>" class=form bgcolor=#DADADA style="padding-right:15px;padding-left:15px;" nowrap><b><!--תאריך פתיחה--><%=arrTitles(9)%></b></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=100% bgcolor=#f0f0f0>
	<input type=image src='../../images/calend.gif' border=0  onclick='return popupcal(this.form.elements("dateStart"));' ID="Image1" NAME="Image1">&nbsp;<input dir="ltr" type="text" class="passw" style="width:85" id="dateStart" name="dateStart" value="<%=dateStart%>" <%If trim(project_ID) = "" Then%> disabled <%End If%>  maxlength=10 onclick="return popupcal(this);" readonly>
	</td>
</tr>
</tbody>
<%If trim(status) = "3" Then%>
<tr>	
	<td align="<%=align_var%>" nowrap class=form style="padding-right:15px;padding-left:15px;" bgcolor=#DADADA><b><!--תאריך סגירה--><%=arrTitles(10)%></b></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=100% bgcolor=#f0f0f0>	
	<input type=image src='../../images/calend.gif' border=0  onclick='return popupcal(this.form.elements("dateEnd"));' ID="Image2" NAME="Image2">&nbsp;<input dir="ltr" type="text" class="passw" style="width:85" id="dateEnd" name="dateEnd" value="<%=dateEnd%>" maxlength=10 onclick="return popupcal(this);" readonly>
	</td>
</tr>
<%End If%> 
<!-- start fields dynamics -->	
<%If project_ID <> nil Then%>	
<!--#INCLUDE FILE="upd_fields.asp"-->	
<%Else%>
<!--#INCLUDE FILE="chform_fields.asp"-->	
<%end if%>
<!-- end fields dynamics -->
<tr bgcolor=#E6E6E6>
<td class="form paddingr">
<table cellpadding=0 cellspacing=0 align=center dir="<%=dir_var%>" width=100% ID="Table1">
<tr bgcolor=#E6E6E6><td colspan="2" height=10></td></tr>	
<tr>
<td width=50% align="center" nowrap><input type=button class="but_menu" style="width:90px" onclick="closeWin()" value="<%=arrButtons(2)%>" id=Button2 name=Button2></td>
<td width=50% align="center" nowrap><input type=button class="but_menu" style="width:90px" onclick="javascript:CheckFields(); return false;" value="<%=arrButtons(1)%>" id=Button1 name=Button1></td></tr>
<tr><td colspan="2" height=10></td></tr>
</table></form>
</td></tr></table>
 <SCRIPT LANGUAGE=javascript>
<!-- 
		 <%
			If trim(lang_id) = "1" Then
				str_alert = "!!!נא למלא את השדה"
			Else
				str_alert = "Please provide the answer!!!"
			End If	
		%> 		 
		function click_fun()
		{ 
        <% sqlStr = "SELECT Field_Id,Field_Must,Field_Type FROM FORM_FIELD Where product_id=" & prod_id & " Order by Field_Order"
		'Response.Write sqlStr
		'Response.End
		set fields=con.GetRecordSet(sqlStr)
		 do while not fields.EOF 
		  Field_Id = fields("Field_Id")				
		  Field_Must = trim(fields("Field_Must"))
		  Field_Type = trim(fields("Field_Type"))
		  'Response.Write  Field_Type
		  If Field_Must = "True"  Then		  
		 %>		
       field =  window.document.all("field<%=Field_Id%>");         
    <%If trim(Field_Type) = "8" Or trim(Field_Type) = "9" Or trim(Field_Type) = "11" Or trim(Field_Type) = "12" Then%>			
		if(field != null)
		{
			answered = false;
			for(j=0;j<field.length;j++)
			{									
				if(field(j).checked == true)
				{
					answered = true;					
				}	
			}		
			if(answered == false)
			{
				window.alert("<%=str_alert%>");			
				field(0).focus();			 
				return false;
			}
		}	  	 	   
	  <%Else%>
	    if((field != null) && document.all("field<%=Field_Id%>").value == '') 
	      { document.all("field<%=Field_Id%>").focus(); 
		    window.alert("<%=str_alert%>"); 
		    return false; } 
	    <%
	     End If	 
	   End If
       fields.moveNext()
	   loop
       set fields=nothing				
    %>    return true;	}
//-->
</SCRIPT>  
</body>
<%
set con = nothing
%>
</html>

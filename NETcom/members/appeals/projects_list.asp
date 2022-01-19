<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%CompanyId = trim(Request.QueryString("CompanyId"))
      If IsNumeric(CompanyId) Then
		Set rs = con.getRecordSet("SELECT Company_Name FROM dbo.Companies WHERE (Company_Id = " & cInt(CompanyId) & ")")
		If Not rs.EOF Then
			CompanyName = trim(rs(0))
		End If
		Set rs = Nothing
      End If
	
	if trim(Request.Form("project_name")) <> "" Or trim(Request.QueryString("project_name")) <> "" then
		project_name = trim(Request.Form("project_name"))
		if trim(Request.QueryString("project_name")) <> "" then
			project_name = trim(Request.QueryString("project_name"))
		end if
		is_search = true
		
		where_project_name = " and UPPER(projects.project_name) LIKE '%"& UCase(sFix(project_name)) &"%'"
	else
		where_project_name = ""
	end if
	
	if trim(Request.Form("project_code")) <> "" Or trim(Request.QueryString("project_code")) <> "" then
		project_code = trim(Request.Form("project_code"))
		if trim(Request.QueryString("project_code")) <> "" then
			project_code = trim(Request.QueryString("project_code"))
		end if
		is_search = true
		
		where_project_code = " and UPPER(projects.project_code) LIKE '"& UCase(sFix(project_code)) &"%'"
	else
		where_project_code = ""
	end if
	
	if Request("Page")<>"" then
	Page=request("Page")
	else
		Page=1
	end if
	if Request.QueryString("all") <> nil then
		all = 1
		show_all = true		
	else
		all = 0
		show_all = false
	end if	%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">	  
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html;">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE="javascript" type="text/javascript">
<!--
	function choose(ProjectId, ProjectName)
	{		
		if(window.opener.document.getElementById("project_id") != null)
		{
			window.opener.document.getElementById("project_id").value = ProjectId;
		}	
			
		if(window.opener.document.getElementById("projectName") != null)	
		{
			window.opener.document.getElementById("projectName").value = unescape(ProjectName);
		}	
			
		if(window.opener.document.getElementById("contacter_body") != null)	
		{
			window.opener.document.getElementById("contacter_body").style.display = "inline";
		}	
		
		if(window.opener.document.getElementById("mechanism_body") != null)	
		{
			window.opener.document.getElementById("mechanism_body").style.display = "inline";
			window.opener.chooseProject(ProjectId);
		}
		
		if(window.opener.document.getElementById("project_body") != null)	
		{				
			window.opener.document.getElementById("project_body").style.display = "inline";	 
		}	
		
		self.close();
		return false;
	}

	function submitenter(myfield,e)
	{
		var keycode;
		if (window.event) keycode = window.event.keyCode;
		else if (e) keycode = e.which;
		else return true;

		if (keycode == 13)
		{
			myfield.form.submit();
			return false;
		}
		else
		return true;
	}
//-->
</SCRIPT>
</HEAD>
<BODY style="margin:0px" onload="window.focus()"><%=trim(Request.Cookies("bizpegasus")("ProjectMulti"))%>
<table width="100%" align="center" cellpadding=0 cellspacing=0 border=0>
<tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
	<td align="<%=align_var%>" class="page_title" nowrap width="70" style="border-right: none"><A class="but_menu" href='projects_list.asp?all=1&amp;CompanyId=<%=CompanyId%>' target="_self" style="width: 70px;" >הצג הכול</a></td>
	<td align="<%=align_var%>" class="page_title" nowrap width="70" style="border-right: none"><A class="but_menu" href="javascript:void(0)" onclick="form_search.submit();"style="width:100%">חפש</a></td>
	<td align="<%=align_var%>" class="page_title" width="100%" style="border-left: none">&nbsp;<%=trim(Request.Cookies("bizpegasus")("ProjectMulti"))%>&nbsp;<%=CompanyName%></td>
</tr>
</table>
</td></tr>
<TR>
<td>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="100%" bgcolor="#E4E4E4" align="middle" valign="top"><!-- start code --> 
		<FORM action="projects_list.asp?CompanyId=<%=CompanyId%>" method="post" id="form_search" name="form_search" style="margin: 0px;">
		<table align="<%=align_var%>" border="0" cellpadding="1" cellspacing="1" bgcolor=white width="100%" dir="<%=dir_var%>">
		<tbody bgcolor="#E4E4E4">
		<tr>
		 <td>&nbsp;</td>
		 <td>&nbsp;<INPUT type=text class="search" dir="<%=dir_obj_var%>" id="project_code" name="project_code" value="<%=trim(vFix(project_code))%>" onKeyPress="return submitenter(this,event)" style="width:100%"></td>
         <td>&nbsp;<INPUT type=text class="search" dir="<%=dir_obj_var%>" id="project_name" name="project_name" value="<%=trim(vFix(project_name))%>" onKeyPress="return submitenter(this,event)" style="width:100%"></td>
		</tr>
		</tbody>
		<tr>						
			<td class="title_sort" align="center" width="5%">&nbsp;בחר&nbsp;</td>
			<td class="title_sort" align="center" width="20%">&nbsp;קוד&nbsp;</td>			
			<td class="title_sort" align="center" nowrap width="75%">&nbsp;<%=trim(Request.Cookies("bizpegasus")("ProjectsOne"))%>&nbsp;</td>			
		</tr>
<%sqlstr = "SELECT project_id, project_name, project_code FROM dbo.Projects WHERE ORGANIZATION_ID = " & cInt(OrgID) &_
	 " AND (Company_Id = " & cInt(CompanyId) & ") " & where_project_code & where_project_name & " ORDER BY project_NAME"	
	'Response.Write sqlstr
	'Response.End
	set pr=con.GetRecordSet(sqlstr)    
	if not pr.eof then 
		recnom = 1		
		pr.PageSize = 10			
		pr.AbsolutePage=Page		
		NumberOfPages = pr.PageCount
		page_size = pr.PageSize
		prArray = pr.GetRows(page_size)	
		recCount=pr.RecordCount 	
		pr.Close
		set pr=Nothing    
		i=1
		j=0
		do while (j<=ubound(prArray,2))			
			project_ID = trim(prArray(0,j))
			projectNAME = trim(prArray(1, j))					
			projectCode = trim(prArray(2, j))	%>	
		<tr bgcolor="#E6E6E6">			
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data"><input type="button" ID="btnSend" NAME="btnSend" value="בחר" onclick="return choose('<%=project_ID%>', escape('<%=altFix(projectNAME)%>'));" class="but_site"></td>
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data"><A class="link_categ" onclick="return choose('<%=project_ID%>', escape('<%=altFix(projectNAME)%>'));" href="javascript:void(0)">&nbsp;<%=projectCode%>&nbsp;</a></td>			
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data"><A class="link_categ" onclick="return choose('<%=project_ID%>', escape('<%=altFix(projectNAME)%>'));" href="javascript:void(0)">&nbsp;<%=projectNAME%>&nbsp;</a></td>
		</tr>
	<%	'pr.movenext
		recnom = recnom + 1
		j = j + 1
		loop%>			
	    <%if NumberOfPages > 1 then%>
		<tr>
		<td width="100%" align="middle" colspan="3" nowrap dir="ltr" class="card">
			<table border="0" cellspacing="0" cellpadding="2">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRow") <> nil Then
	               numOfRow = Request.QueryString("numOfRow")
	           Else numOfRow = 1
	           End If
	         %>
	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center"><A class=pageCounter title="לדפים הקודמים" href="projects_list.asp?all=<%=all%>&amp;project_name=<%=vFix(project_name)%>&amp;project_code=<%=vFix(project_code)%>&amp;page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>&CompanyId=<%=CompanyId%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td class="form" valign="top"><font color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle"><p style="PADDING-RIGHT: 2px; PADDING-LEFT: 2px; FONT-WEIGHT: bold; FONT-SIZE: 9pt; BACKGROUND: #74b3d5; CURSOR: default; COLOR: #ffffff" 
                 ><%=i+10*(numOfRow-1)%></p></td>
	                  <%else%>
	                     <td align="middle"><A class=pageCounter href="projects_list.asp?all=<%=all%>&amp;project_name=<%=vFix(project_name)%>&amp;project_code=<%=vFix(project_code)%>&amp;page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>&CompanyId=<%=CompanyId%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>	            
					<td class="form" valign="top"><font color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center"><A class=pageCounter title="לדפים הבאים" href="projects_list.asp?all=<%=all%>&amp;project_name=<%=vFix(project_name)%>&amp;project_code=<%=vFix(project_code)%>&amp;page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>&CompanyId=<%=CompanyId%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>					
<%End If%>	
	<tr><td class="card" align="center" colspan="3" style="color:#5F435A;font-weight:600" dir="ltr">נמצאו&nbsp;<%=recCount%>&nbsp;רשומות</td></tr>					
<%Else 'not pr.eof%>
	<tr><td align="center" colspan="3" bgcolor="#E4E4E4" style="font-size:12px;font-weight:600;color:red;">לא נמצאו רשומות</td></tr>	
<%End If 'not pr.eof %>
<!-- end code -->
</form></table>
</td></tr></table>
</td></tr></table>
</BODY>
</HTML>
<%set con =nothing%>
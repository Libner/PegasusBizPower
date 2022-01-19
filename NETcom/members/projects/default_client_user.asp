<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
	companyID = trim(Request("companyID"))
	curr_user_id = trim(Request("curr_user_id"))
	projectId = trim(Request("projectId"))
	mechanismId = trim(Request("mechanismId"))
	start_date = "1/" & Month(Date()) & "/" & Year(date())
	end_date = DateAdd("d",30,start_date)
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language=javascript>
<!--
	
	function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=150pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}

	function report_open(fname)
	{	
		formsearch.action=fname;
		formsearch.target = '_blank';
		formsearch.submit();	
    }
	
//-->
</script>
</head>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 58 Order By word_id"				
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
<body topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 bgcolor="#E5E5E5">
<!--#include file="../../logo_top.asp"-->
<div align="center"><center>
<table border="0" width="100%" cellspacing="0" cellpadding="0"  dir="<%=dir_var%>">
<%numOftab = 3%>
<%numOfLink = 7%>
<%topLevel2 = 30 'current bar ID in top submenu - added 03/10/2019%>
  <tr><td width=100% align="<%=align_var%>"><!--#include file="../../top_in.asp"--></td></tr>
</table>
</center></div>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<tr><td class="page_title" dir="<%=dir_obj_var%>"><span id=word1 name=word1><!--דוח עלות לפי עובדים לתקופה--><%=arrTitles(1)%></span></td></tr>

<tr><td height=20 nowrap></td></tr>
<tr>
   <td width="100%" align=center valign=top colspan=2>      
	<table border="0" cellspacing="1" width=450 align=center bgcolor="#646E77" cellpadding="4">
	<FORM action="" method=POST id=formsearch name=formsearch>
	<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">		
		<select name="curr_user_id" id="curr_user_id" class="norm" dir="<%=dir_obj_var%>" style="width:310"
		onchange="document.location.href='default_client_user.asp?curr_user_id='+this.value+'&projectId='+project_id.value+'&companyId='+company_id.value+'&mechanismId='+mechanism_id.value">		
		<option value="" id=word2 name=word2><%=arrTitles(2)%></option>
		<%  sqlstr = "Select USER_ID, FIRSTNAME + ' ' + LASTNAME FROM USERS WHERE ORGANIZATION_ID = " & OrgID			
		    sqlstr = sqlstr & " Order BY FIRSTNAME + ' ' + LASTNAME "
			set rs_comp = con.getRecordSet(sqlstr)
			While not rs_comp.eof
		%>
			<option value="<%=rs_comp(0)%>" <%If trim(curr_user_id) = trim(rs_comp(0)) Then%> selected <%End if%>><%=rs_comp(1)%></option>			
		<%
			rs_comp.moveNext
			Wend
			set rs_comp = Nothing
		%>
		</select>
		</TD>
		<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<span id="word3" name=word3><!--עובד--><%=arrTitles(3)%></span>&nbsp;</td>
	</TR>
	<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">
		<select name="company_id" id="company_id" class="norm" dir="<%=dir_obj_var%>" style="width:310"
		onchange="document.location.href='default_client_user.asp?curr_user_id='+curr_user_id.value+'&projectId='+project_id.value+'&companyId='+this.value+'&mechanismId='+mechanism_id.value">
		<%If trim(lang_id) = "1" Then%>
		<option value=""><%=String(25,"-")%> כל ה<%=trim(Request.Cookies("bizpegasus")("CompaniesMulti"))%> <%=String(25,"-")%></option>	
		<%Else%>
		<option value=""><%=String(25,"-")%> All <%=trim(Request.Cookies("bizpegasus")("CompaniesMulti"))%> <%=String(25,"-")%></option>	
		<%End If%>
		<%  sqlstr = "Select DISTINCT Company_ID, Company_NAME FROM HOURS_VIEW WHERE ORGANIZATION_ID = " & OrgID
		    If trim(projectId) <> "" Then
			sqlstr = sqlstr & " AND project_Id = " & projectId
			End if	
			If trim(curr_user_id) <> "" Then
				sqlstr = sqlstr & " AND User_ID = " & curr_user_id
			End if	    
		    sqlstr = sqlstr & " AND Company_ID IS NOT NULL Order BY Company_NAME "
			set rs_comp = con.getRecordSet(sqlstr)
			While not rs_comp.eof
		%>
			<option value="<%=rs_comp(0)%>" <%If trim(companyID) = trim(rs_comp(0)) Then%> selected <%End if%>><%=rs_comp(1)%></option>			
		<%
			rs_comp.moveNext
			Wend
			set rs_comp = Nothing
		%>		
		</select>
		</TD>
		<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;</td>
	</TR>
	<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">
		<select name="project_id" id="project_id" class="norm" dir="<%=dir_obj_var%>" style="width:310"
		onchange="document.location.href='default_client_user.asp?curr_user_id='+curr_user_id.value+'&companyId='+company_id.value+'&projectId='+this.value+'&mechanismId='+mechanism_id.value">
		<%If trim(lang_id) = "1" Then%>
		<option value=""><%=String(25,"-")%> כל ה<%=trim(Request.Cookies("bizpegasus")("ProjectMulti"))	%> <%=String(25,"-")%></option>		
		<%Else%>
		<option value=""><%=String(27,"-")%> All <%=trim(Request.Cookies("bizpegasus")("ProjectMulti"))	%> <%=String(27,"-")%></option>		
		<%End If%>
		<%  
			con.executeQuery("SET DATEFORMAT dmy")
			sqlstr = "Select DISTINCT PROJECT_ID, PROJECT_NAME, status FROM HOURS_VIEW WHERE ORGANIZATION_ID = " & OrgID
			If trim(companyID) <> "" Then
			sqlstr = sqlstr & " AND company_id IN (0," & companyID  & ")"
			End if	
			If trim(curr_user_id) <> "" Then
				sqlstr = sqlstr & " AND User_ID = " & curr_user_id
			End if			
			sqlstr = sqlstr & " AND PROJECT_ID IS NOT NULL Order BY PROJECT_NAME"
			'Response.Write sqlstr
			'Response.End
			set rs_comp = con.getRecordSet(sqlstr)
			While not rs_comp.eof
		%>
			<option value="<%=rs_comp(0)%>" <%If trim(rs_comp(2)) = "3" Then%> class="status_num<%=rs_comp(2)%>" <%End If%>  <%If trim(projectId) = trim(rs_comp(0)) Then%> selected <%End if%>><%=rs_comp(1)%></option>		
		<%
			rs_comp.moveNext
			Wend
			set rs_comp = Nothing
		%>		
		</select>
		</TD>
		<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<%=trim(Request.Cookies("bizpegasus")("Projectone"))	%>&nbsp;</td>
	</TR>
	<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">
		<select dir="<%=dir_obj_var%>" name="mechanism_id" style="width:310px;font-family:Arial" ID="mechanism_id">
		<option value=""><%=arrTitles(7)%></option>
        <%
		sqlstr = "Select mechanism_id, mechanism_name, project_name From mechanism Inner Join Projects On "&_
		" mechanism.project_id = projects.project_id Where mechanism.ORGANIZATION_ID = " & OrgID
	    If trim(projectID) <> "" Then
	       sqlstr = sqlstr & " AND mechanism.PROJECT_ID = " & projectID
	    End If
	    If trim(companyId) <> "" Then
			sqlstr = sqlstr & " And mechanism.COMPANY_ID = " & companyId 
	    End If
	    sqlstr = sqlstr & " Order BY project_name, mechanism_name"
	    'Response.Write sqlstr
	    'Response.End
		set rs_mech = con.getRecordSet(sqlstr)
		Do While not rs_mech.eof
	    %>
	    <option value="<%=rs_mech(0)%>" <%If trim(mechanismId) = trim(rs_mech(0)) Then%> selected <%End If%>><%=rs_mech(2)%> - <%=rs_mech(1)%></option>
	    <%
		rs_mech.moveNext
		Loop
		set rs_mech = Nothing
	%>
	</select>		
	</TD>
	<td width="20%" class="subject_form" align="<%=align_var%>"><b><!--מנגנון--><%=arrTitles(8)%></b>&nbsp;</td>
	</tr>			
	<TR>						
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>"><input type=image src='../../images/calend.gif'   border=0  onclick='return popupcal(dateStart);'>&nbsp;<input dir="ltr" type="text" class="passw" size=12 id="dateStart" name="dateStart" value="<%=start_date%>" maxlength=10 readonly>&nbsp;</TD>
		<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<span id=word4 name=word4><!--- תאריך מ--><%=arrTitles(4)%></span>&nbsp;</td>
	</TR>
	<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>"><input type=image src='../../images/calend.gif'   border=0  onclick='return popupcal(dateEnd);'>&nbsp;<input dir="ltr" type="text" class="passw" size=12 id="dateEnd" name="dateEnd" value="<%=end_date%>" maxlength=10 readonly>&nbsp;</TD>
			<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<span id="word5" name=word5><!--- תאריך עד--><%=arrTitles(5)%></span>&nbsp;</td>
	</TR>	
		<tr id=but_tr>
			<td bgcolor="#dbdbdb" align="center" colspan="2">
				<TABLE WIDTH=100% align=center BORDER=0 CELLSPACING=5 CELLPADDING=0>
				 	<TR>
						<td width=20%>&nbsp;</td>
						<TD align=center width=60%><A class="but_menu" href="#" onclick="report_open('report_client_user.asp');"><span id=word6 name=word6><!--הצג דוח--><%=arrTitles(6)%></span></a></TD>
						<td width=20%>&nbsp;</td>
					</TR>																						
					</table></td></tr>
			</form>		
  <!-- End main content -->
</td></tr></table>
</td></tr></table>
</center></div>
</body>
<%set con=nothing%>
</html>

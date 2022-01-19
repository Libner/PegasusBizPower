<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	projectID=Request.QueryString("projectID")
	companyID=Request.QueryString("companyID")
	mechID=Request.QueryString("mechID")

	If Request.Form("mechanism_name")<>nil Then 'after form filling

		mechanism_name = sFix(trim(Request.Form("mechanism_name")))
		budget_hours = sFix(trim(Request.Form("budget_hours")))
		proposal_hours = sFix(trim(Request.Form("proposal_hours")))
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
		 
		If mechID=nil Or mechID="" Then 'new record in DataBase		   
			sqlstr = "SET DATEFORMAT DMY; Insert Into mechanism (project_id,company_id,mechanism_name,budget_hours,"&_
			" start_date,end_date,Organization_ID,proposal_hours) Values (" & projectID & "," & companyID & ",'" & mechanism_name & "','" &_
			budget_hours & "'," & dateStart & "," & dateEnd & "," & orgID & ",'" & proposal_hours & "')"
			'Response.Write sqlStr
			con.GetRecordSet(sqlStr)						    
		Else			   
			sqlStr = "SET DATEFORMAT DMY; Update mechanism Set mechanism_name = '" & mechanism_name &_
			"', budget_hours = '" & budget_hours & "', proposal_hours = '" & proposal_hours &_
			"', start_date = " & dateStart & ", end_date = " & dateEnd &_
			" Where mechanism_ID=" & mechID
			con.GetRecordSet(sqlStr)
		End If        
		
		%>
		<script language=javascript>
		<!--
			window.opener.document.location.reload(true);
			window.close();	
		//-->
		</script>
		<%
	end if
	
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 10 Order By word_id"				
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
	
	sqlstr = "Select value From buttons Where lang_id = " & lang_id & " Order By button_id"
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
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
<!--
function CheckFields()
{	
	if (document.frmMain.mechanism_name.value=='')		   
	{
		<%
			If trim(lang_id) = "1" Then
				str_alert = "אנא הכנס שם של המנגנון"
			Else
				str_alert = "Please enter the mechanism name"
			End If   
		%>			
		window.alert("<%=str_alert%>");		
		document.frmMain.mechanism_name.focus();
		return false;				
	}
	document.frmMain.submit();
	return true;
	
}

function GetNumbers ()
{
	var ch=event.keyCode;
	event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
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
</head> 
<body style="margin:0px;background-color:#e6e6e6" onload="window.focus()">
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title"><%if mechID<>nil then%><!--עדכון--><%=arrTitles(39)%><%else%><!--הוספת--><%=arrTitles(38)%><%end if%>&nbsp;<!--מנגנון--><%=arrTitles(65)%>&nbsp;</td></tr>		   
     </table></td></tr>  
     <tr><td height=15 nowrap></td></tr>
	<tr><td width=100% bgcolor="#E6E6E6"> 
	<FORM name="frmMain" ACTION="addmechanism.asp?mechID=<%=mechID%>&projectID=<%=projectID%>&companyID=<%=companyID%>" METHOD="post" onSubmit="return CheckFields()" ID="frmMain">
	<table align=center border="0" cellpadding="3" cellspacing="1" width="100%" dir="<%=dir_var%>">    
	<%
	if mechID<>nil and mechID<>"" then
	sqlStr = "Select mechanism_name, budget_hours, proposal_hours, start_date, end_date From mechanism Where mechanism_ID=" & mechID  
	''Response.Write sqlStr
	set rs_mechanisms = con.GetRecordSet(sqlStr)
		if not rs_mechanisms.eof then
			mechanism_name = trim(rs_mechanisms("mechanism_name"))
			budget_hours = trim(rs_mechanisms("budget_hours"))
			proposal_hours = trim(rs_mechanisms("proposal_hours"))
			dateStart = trim(rs_mechanisms("start_date"))
			dateEnd = trim(rs_mechanisms("end_date"))
		end if
		set rs_mechanisms = nothing
	end if
	'Response.Write company_id
	%>
	<tr>
		<td align="<%=align_var%>"><input type=text dir="<%=dir_obj_var%>" name="mechanism_name" value="<%=vfix(mechanism_name)%>" style="width:300px;font-family:arial"></td>
		<td align="<%=align_var%>" nowrap width=100><!--שם מנגנון--><%=arrTitles(65)%>&nbsp;<span style="color:#FF0000">*</span></td>
	</tr>
	<tr>
		<td align=right>	
		<input type=image src='../../images/calend.gif' border=0  onclick='return popupcal(this.form.elements("dateStart"));' ID="Image1" NAME="Image1">&nbsp;<input dir="ltr" type="text" class="passw" style="width:85" id="dateStart" name="dateStart" value="<%=dateStart%>" maxlength=10 onclick="return popupcal(this);" readonly>
		</td>
		<td align="right" nowrap width="100">&nbsp;תאריך פתיחה&nbsp;</td>
	</tr>
	<tr>
		<td align=right>	
		<input type=image src='../../images/calend.gif' border=0  onclick='return popupcal(this.form.elements("dateEnd"));' ID="Image2" NAME="Image2">&nbsp;<input dir="ltr" type="text" class="passw" style="width:85" id="dateEnd" name="dateEnd" value="<%=dateEnd%>" maxlength=10 onclick="return popupcal(this);" readonly>
		</td>
		<td align="right" nowrap width="100">&nbsp;תאריך סיום&nbsp;</td>
	</tr>	
	<tr>
		<td align="<%=align_var%>"><input type=text dir="<%=dir_obj_var%>" name="budget_hours" value="<%=vfix(budget_hours)%>" style="width:100px;font-family:arial" ID="budget_hours" onKeypress="GetNumbers();" maxlength=10></td>
		<td align="<%=align_var%>" nowrap width=100><!--תקצוב שעות--><%=arrTitles(67)%></td>
	</tr>
	<tr>
		<td align="<%=align_var%>"><input type=text dir="<%=dir_obj_var%>" name="proposal_hours" value="<%=vfix(proposal_hours)%>" style="width:100px;font-family:arial" ID="proposal_hours" onKeypress="GetNumbers();" maxlength=10></td>
		<td align="<%=align_var%>" nowrap width=100>הצעה</td>
	</tr>		
	<tr><td colspan="2" height=10></td></tr>	
	<tr>
	<td colspan=2>
	<table cellpadding=0 cellspacing=0 width=80% align=center dir="<%=dir_var%>">
	<tr>
	<td width=50% align="center"><input type=button class="but_menu" style="width:90px" onclick="javascript:window.close();" value="<%=arrButtons(2)%>"></td>
	<td width=50 nowrap></td>
	<td width=50% align="center"><input type=button class="but_menu" style="width:90px" onclick="javascript:CheckFields(); return false;" value="<%=arrButtons(1)%>"></td></tr>
</table></td></tr>
</form></table>
</body>
<%set con = nothing%>
</html>
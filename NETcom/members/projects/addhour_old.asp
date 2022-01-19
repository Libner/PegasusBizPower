<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
<!--
   	
	function GetNumbers ()
	{
		var ch=event.keyCode;
		event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
	} 	
	
	function checkFields()
	{
		if(window.document.all("company_id").value == "")
		{
			window.alert("אנא בחר <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%> !!");
			window.document.all("company_id").focus();
			return false;
		}
		if(window.document.all("project_id").value == "")
		{
			window.alert("!! אנא בחר <%=trim(Request.Cookies("bizpegasus")("Projectone"))%>");
			window.document.all("project_id").focus();
			return false;
		}
		if(window.document.all("hours").value == "")
		{
			window.alert("!! אנא מלא את שעות העבודה על ה<%=trim(Request.Cookies("bizpegasus")("Projectone"))	%>");
			window.document.all("hours").focus();
			return false;
		}
	}
<!--End-->
</script> 
</head> 
<%
UserID=trim(Request.Cookies("bizpegasus")("UserID"))
OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
date_ = trim(Request("date"))
companyId = trim(Request("companyId"))
end_date = Request("end_date")
start_date = Request("start_date")
If trim(Request("USER_ID")) <> "" Then
	UserID = trim(Request("USER_ID"))
End If
if Request("add")<>nil then		
    company_id = trim(Request.Form("company_id"))   
    project_id = trim(Request.Form("project_id"))   
    hours = trim(Request.Form("hours"))   
    If trim(hours) <> "" Then
		hours = cDbl(hours)
		minuts = hours * 60		
	    con.executeQuery("SET DATEFORMAT dmy")
	    sqlstr = "Insert Into hours (user_id,organization_id,date,project_id,company_id,minuts) VALUES (" &_
	    userID & "," & orgID & ",'" & date_ & "'," & project_id & "," & company_id & "," & minuts & ")"
	    'Response.Write sqlstr
	    'Response.End
	    con.executeQuery(sqlstr)
	    
	%>
	<script language=javascript>
	<!--
		window.opener.document.location.href = window.opener.document.location.pathname + "?date_=<%=date_%>&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=UserID%>";
		window.close();	
	//-->
    </script>
  
	<%
	End If
end if

%>
<body style="margin:0px;background:#E6E6E6" onload="window.focus();">
<table border="0" cellpadding="0" cellspacing="1" width="100%" ID="Table7">
<tr>
   <td bgcolor="#e6e6e6" align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0" ID="Table8">
	  <tr><td class="page_title">הוספת שעות עבודה על <%=trim(Request.Cookies("bizpegasus")("Projectone"))	%>&nbsp;</td></tr>		   
	  		       	
   </table></td></tr>         
   <tr>    
    <td width="100%" valign="top" align="center">    
	<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1">	
	<form name="add_hour" id="add_hour" method=post action="addhour.asp?date=<%=date_%>&add=1&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=UserID%>">	
	<tr>
	<td align=right width="310"><select dir="rtl" name="company_id" style="width:300px;font-family:arial" onchange="document.location.href='addhour.asp?date=<%=date_%>&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=UserID%>&companyId='+this.value;">
	<option value=""><%=String(20,"-")%> בחר <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%> <%=String(20,"-")%></option>
	<%  sqlstr = "Select COMPANY_ID, COMPANY_NAME FROM COMPANIES Where ORGANIZATION_ID = "& OrgID & " Order BY COMPANY_NAME"
		set rs_comp = con.getRecordSet(sqlstr)
		While not rs_comp.eof
	%>
	<option value="<%=rs_comp(0)%>" <%If trim(companyId) = trim(rs_comp(0)) Then%> selected <%End If%>><%=rs_comp(1)%></option>
	<%
		rs_comp.moveNext
		Wend
		set rs_comp = Nothing
	%>
	</select></td>
	<td align=right><b><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></b>&nbsp;</td>
	</tr>
	<%If trim(companyId) <> "" Then%>
	<tr>		
	<td align=right width="310"><select dir="rtl" name="project_id" style="width:300px;font-family:arial" ID="Select2">
	<option value=""><%=String(20,"-")%> בחר <%=trim(Request.Cookies("bizpegasus")("Projectone"))	%> <%=String(20,"-")%></option>
	<%  
		con.executeQuery("SET DATEFORMAT dmy")
		sqlstr = "Select PROJECT_ID, PROJECT_NAME FROM PROJECTS WHERE company_id IN (0," & companyID &_
	    ") AND ORGANIZATION_ID = " & OrgID & " AND PROJECT_ID NOT IN (Select DISTINCT PROJECT_ID FROM hours WHERE ORGANIZATION_ID = " & orgID &_
	    " And USER_ID = " & userID & " And COMPANY_ID = " & companyId & " AND date = '" & date_ & "') AND active = 1 AND status = '2' Order BY PROJECT_NAME"
	    'Response.Write sqlstr
		set rs_comp = con.getRecordSet(sqlstr)
		While not rs_comp.eof
	%>
	<option value="<%=rs_comp(0)%>" <%If trim(project_id) = trim(rs_comp(0)) Then%> selected <%End If%>><%=rs_comp(1)%></option>
	<%
		rs_comp.moveNext
		Wend
		set rs_comp = Nothing
	%>
	</select></td>
	<td align=right><b><%=trim(Request.Cookies("bizpegasus")("Projectone"))	%></b>&nbsp;</td>
	</tr>
	<tr>
	<td align=right><input type=text class="texts" name="hours" id="hours" value="<%=hours_pr%>" onkeypress="GetNumbers();" maxlength=4 style="width:45"></td>		
	<td align=right><b>שעות</b>&nbsp;</td>
	</tr>
	<%End If%>	
	<tr><td colspan="2" height=20 nowrap></td></tr>
	<tr><td colspan="2">
	<table width=100% border=0 cellspacing=0 ID="Table1">
	<tr><td width=48% align="right">
	<INPUT class="but_menu" style="width:90" type="button" value="ביטול" id=button2 name=button2 onclick="window.close();"></td>
	<td width=4% nowrap></td>
	<td width=48% align="left">
	<input class="but_menu" style="width:90" type="submit" value="אישור" id=submit1 name=submit1 onclick="return  checkFields()"></td>
	</tr>
	</table>
	</td></tr>
	</form>		
</table>
</td>
</tr></table>
</body>
</html>
<%
set con = nothing
%>
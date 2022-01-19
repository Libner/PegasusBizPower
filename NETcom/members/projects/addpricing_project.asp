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
			window.alert("!! אנא בחר <%=trim(Request.Cookies("bizpegasus")("Projectone"))	%>");
			window.document.all("project_id").focus();
			return false;
		}
		
	}
<!--End-->
</script> 
</head> 
<%

UserID=trim(Request.Cookies("bizpegasus")("UserID"))
OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))

pricing_id = trim(Request.QueryString("pricing_id"))
if pricing_id<>nil and pricing_id<>"" then
  sqlStr = "select project_name from projects where project_id=" & pricing_id  
  ''Response.Write sqlStr
  set rs_projects = con.GetRecordSet(sqlStr)
	if not rs_projects.eof then
		project_name = rs_projects("project_name")		
	end if
	set rs_projects = nothing	
end if

companyId = trim(Request("companyId"))
projectID = trim(Request("projectID"))

if Request("add")<>nil then		
    
    company_id = trim(Request.Form("company_id"))   
    project_id = trim(Request.Form("project_id")) 
    mechanism_id = trim(Request.Form("mechanism_id"))   
    
	sqlstr = "Insert Into pricing_to_projects (mechanism_id,project_id,company_id,pricing_id)"&_
	" VALUES (" & mechanism_id & "," & project_id & "," & company_id & "," & pricing_id & ")"
	'Response.Write sqlstr
	'Response.End
	con.executeQuery(sqlstr)
	
%>
	<script language=javascript>
	<!--
		window.opener.document.location.href = "addPricing.asp?pricing_id=<%=pricing_id%>";
		window.close();	
	//-->
    </script>
  
<%End If%>
<body style="margin:0px;background:#E6E6E6" onload="window.focus();">
<table border="0" cellpadding="0" cellspacing="1" width="100%" ID="Table7">
<tr>
   <td bgcolor="#e6e6e6" align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0" ID="Table8">
	  <tr><td class="page_title">הוספת מנגנון להשוואה&nbsp;<font color="#6F6DA6"><%=pricing_name%></font></td></tr>		   
	  		       	
   </table></td></tr>         
   <tr>    
    <td width="100%" valign="top" align="center">    
	<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1">	
	<form name="add_hour" id="add_hour" method=post action="addpricing_project.asp?add=1&pricing_id=<%=pricing_id%>">	
	<tr>
	<td align=right width="310"><select dir="rtl" name="company_id" style="width:300px;font-family:arial"
	onchange="document.location.href='addpricing_project.asp?pricing_id=<%=pricing_id%>&companyId='+this.value;">
	<option value=""><%=String(20,"-")%> בחר <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%> <%=String(20,"-")%></option>
	<%  sqlstr = "Select DISTINCT COMPANY_ID, COMPANY_NAME FROM hours_view Where STATUS = '3' AND ORGANIZATION_ID = "& OrgID & " Order BY COMPANY_NAME"
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
	<td align=right width="310">
	<select dir="rtl" name="project_id" style="width:300px;font-family:arial" ID="project_id"
	onchange="document.location.href='addpricing_project.asp?pricing_id=<%=pricing_id%>&companyId='+company_id.value+'&projectID='+this.value;">
	<option value=""><%=String(20,"-")%> בחר <%=trim(Request.Cookies("bizpegasus")("Projectone"))	%> <%=String(20,"-")%></option>
	<%  	
		sqlstr = "Select DISTINCT PROJECT_ID, PROJECT_NAME FROM hours_view WHERE " &_
	    " ORGANIZATION_ID = " & OrgID & " And COMPANY_ID = " & companyId &_
		" AND STATUS = '3' Order BY PROJECT_NAME"
	    'Response.Write sqlstr
		set rs_comp = con.getRecordSet(sqlstr)
		While not rs_comp.eof
	%>
	<option value="<%=rs_comp(0)%>" <%If trim(projectId) = trim(rs_comp(0)) Then%> selected <%End If%>><%=rs_comp(1)%></option>
	<%
		rs_comp.moveNext
		Wend
		set rs_comp = Nothing
	%>
	</select></td>
	<td align=right><b><%=trim(Request.Cookies("bizpegasus")("Projectone"))	%></b>&nbsp;</td>
	</tr>	
	<% End If %>
	<% If trim(companyID) <> "" And trim(projectID) <> "" Then%>
	<tr>		
	<td align="right" width="310">
	<select dir="rtl" name="mechanism_id" style="width:300px;font-family:Arial" ID="mechanism_id">
	<option value=""><%=String(20,"-")%> בחר מנגנון <%=String(20,"-")%></option>
    <%
		sqlstr = "Select mechanism_id, mechanism_name From mechanism Where company_id = " & companyID &_
	    " AND ORGANIZATION_ID = " & OrgID & " AND PROJECT_ID = " & projectID & " AND mechanism_id NOT IN ("&_
	    " Select DISTINCT mechanism_id FROM pricing_to_projects WHERE COMPANY_ID = " & companyId &_
	    " AND pricing_id = " & pricing_id & " And Project_ID = " & projectID & ") Order BY mechanism_id"
	    'Response.Write sqlstr
	    'Response.End
		set rs_mech = con.getRecordSet(sqlstr)
		Do While not rs_mech.eof
	%>
	<option value="<%=rs_mech(0)%>"><%=rs_mech(1)%></option>
	<%
		rs_mech.moveNext
		Loop
		set rs_mech = Nothing
	%>
	</select></td>
	<td align="right"><b>מנגנון</b>&nbsp;</td>
	</tr>	
	<tr>
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
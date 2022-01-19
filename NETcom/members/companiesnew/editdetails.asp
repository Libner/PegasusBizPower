<% server.ScriptTimeout=10000 %>
<% Response.Buffer = True %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
 
  companyID = trim(Request("companyID"))
  
  If trim(companyID)<>"" then   
  sqlStr = "SELECT company_name,company_desc FROM companies WHERE company_id="& companyID 
  'Response.Write "test"
  'Response.End
  set pr=con.GetRecordSet(sqlStr)
  If not pr.EOF then	
	company_name = trim(pr(0))
  	company_desc = trim(pr(1))
  End if 
  
  If Request.Form("company_desc") <> nil Then
	  sqlstr = "Update companies SET company_desc = '" & sFix(Left(trim(Request.Form("company_desc")),500)) & "' WHERE company_id="& companyID 
	  con.executeQuery(sqlstr)
	  %>
	<SCRIPT LANGUAGE=javascript>
	<!--	
		
		window.opener.document.location.reload(true);
		self.close();
	//-->
	</SCRIPT>
	  <%
  End If  
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body style="margin:0;background:'#e6e6e6'">
<table border="0" bordercolor="navy" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top">
<tr><td class="page_title" dir=rtl>&nbsp;<%if trim(companyId)<>"" then%><font style="color:#000000;"><b>עדכון פרטים נוספים - <%=trim(company_name)%></b></font><%elseIf trim(companyId) = "" Then%><font style="color:#000000;"><b>הוספת <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></b><%end if%>&nbsp;</td></tr>         	
<tr><td height=10 nowrap></td></tr>
<tr><td width="100%">
<table cellpadding=1 bgcolor="#e6e6e6" cellspacing=0 width=100% border=0 style="border-collpase:collapse;" ID="Table5">
<FORM name="form_company" ACTION="editdetails.asp?companyID=<%=companyID%>" METHOD="post" ID="form_company">   
  <input type="hidden" name="COMPANY_ID" id="COMPANY_ID" value="<%=companyID%>">  
      <tr>
         <td width="10%" align=center>
         <textarea dir=rtl class=Form name="company_desc" cols=40 rows=6 ID="company_desc"><%=trim(company_desc)%></textarea>
         </td>
      </tr>                  
    <tr><td colspan=2 height="15" nowrap></td></tr>
    <tr>
		 <td colspan=2 align="center" nowrap>
		 <input type=button class="but_menu" value="ביטול" onClick="window.close();" style="width:80" ID="Button1" NAME="Button1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		 <input type=button class="but_menu" value="אישור" onClick="form_company.submit();" style="width:80" ID="Button2" NAME="Button2">
		 </td>		 
	</tr>
    </FORM>                        
</table>
</td>
</tr>
</table>
</body>
<%End If%>
<%set con=Nothing%>
</html>


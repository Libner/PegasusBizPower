<%@ LANGUAGE=VBScript %>
<%Response.Buffer = True%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--	
	function refreshCompany(companyObj)
	{
		if ((event.keyCode==13) || (event.keyCode == 9))
			return false;
		compName = new String(companyObj.value);
		if(compName.length > 0)
		{		
			window.document.all("frameCustomers").src = "companies.asp?search_company=" + compName;	
		}
		else
		{
			window.document.all("frameCustomers").src = "companies.asp";
		}	
		window.focus();
		companyObj.focus();	
		return true;	
				
	}

	function refreshContacts(contactObj)
	{
		if ((event.keyCode==13) || (event.keyCode == 9))
			return false;
		contName = new String(contactObj.value);
		if(contName.length > 0)
		{
			window.document.all("frameCustomers").src = "contacts.asp?" + contactObj.id + "=" + contName;	
		}
		else
		{
			window.document.all("frameCustomers").src = "contacts.asp";
		}	
				
	}	
//-->
</script>
</head>
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table1">
<tr><td width="100%" align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width="100%" align="<%=align_var%>">
  <%numOftab = 71%>
  <%numOfLink =0%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title">&nbsp;<!--אנא הקלד את שם החברה מימין או את שם איש הקשר לאיתור במאגר הלקוחות--><%'=arrTitles(25)%></td></tr>
<tr>    
    <td width="100%" valign="top" align="center">
    <table dir="<%=dir_var%>" border="0" cellpadding="0" cellspacing="1" width="100%" ID="Table2">
    <tr>
    <td align="left" width="100%" valign=top >   
    <iframe name="frameScrreen" id="frameScrreen"  src="screenSetting.aspx" 	ALLOWTRANSPARENCY=true height=1900  width="100%" marginwidth="0" marginheight="0" hspace="0" vspace="0" scrolling="0" frameborder="0"></iframe>
    </td>
</tr></table>
	</td></tr></table>

</body>
</html>
<%set con=Nothing%>
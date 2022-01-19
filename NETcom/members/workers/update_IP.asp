<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%If Request.QueryString("add") <> nil Then
'response.Write "11"
        IPadress = Request.Form("IPadress")
     
    
       If trim(IPadress)<>"" Then
			sqlstr="UPDATE Ip_Address SET IP =' " & (IPadress) &"'"
			con.executeQuery(sqlstr) 
		End If
	'   response.Write "22="& IPadress
     '   response.end	

     
     %>
		<SCRIPT LANGUAGE="javascript">
		<!--
			window.opener.document.location.href = window.opener.document.location.href;
			window.close();
		//-->
		</SCRIPT>	
<%End If%>
<%
	  sqlstr = "Select top 1 IP from Ip_Address"				
	  set rsIp= con.getRecordSet(sqlstr)		
	
	  If not rsIp.eof Then
	IPadress=rsIp("IP")
	  End If
	  set rstitle=nothing	 	  %>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
		function checkForm()
		{		
				
			return true;			   
		}
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<table border="0" width="380" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>" ID="Table1">
<tr >
   <td  align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="1" cellspacing="0" ID="Table2">
	   <tr>		 
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%">&nbsp;IP-עדכון כתובת ה</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width="100%">
<form name="form1" id="form1" action="update_IP.asp?add=1" target="_self" method="post">
<table width="380" cellspacing="1" cellpadding="2" align="center" border="0" ID="Table3">
<tr>
	<td align="center" width="230" nowrap><input type="text" name="IPadress" id="IPadress" value="<%=IPadress%>"  style="width: 250px"></td>
</tr>

<tr><td height="35"  nowrap></td></tr>
<tr><td align="center" >
<input type=button value="ביטול" class="but_menu" style="width:90" onclick="window.close();" id=Button2 name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=submit value="אישור" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1></td></tr>
</table>
</form>
</td></tr></table>
</BODY>
</html>
<%Set con = Nothing%>
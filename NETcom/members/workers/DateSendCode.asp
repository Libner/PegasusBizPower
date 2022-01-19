<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->


<%If Request.QueryString("add") <> nil Then
'response.Write "11"
if IsNumeric(Request.Form("datesend")) then
        datesend =Request.Form("datesend")
       end if
        
     
    
       If trim(datesend)<>"" Then
			sqlstr="UPDATE DateSendCode SET QDays = " & (datesend) &" ,QDaysUpdate=GetDate()"
			con.executeQuery(sqlstr) 
		End If
	'   response.Write "22="& IPadress
     '   response.end	

     
     %>
		
<%End If%>
<%
	  sqlstr = "Select top 1 QDays,QDaysUpdate,LastDateSend from DateSendCode"				
	  set rs= con.getRecordSet(sqlstr)		
	
	  If not rs.eof Then
	datesend=rs("QDays")
	QDaysUpdate=rs("QDaysUpdate")
	LastDateSend=rs("LastDateSend")
	  End If
	  set rstitle=nothing	 	  %>

<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
		function checkForm()
		{		
				
			return true;			   
		}
//-->
</SCRIPT>
</head>
<body>
<form name="form1" id="form1" action="DateSendCode.asp?add=1" target="_self" method="post">

<table bgcolor="#FFFFFF" border="0" width="100%" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>" ID="Table1">
<tr><td width="100%">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<%numOftab = 4%>
<%numOfLink = 17%>
<%topLevel2 = 130 'current bar ID in top submenu - added 03/10/2019%>
<tr><td width="100%">
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title" align="<%=align_var%>" valign="middle" width="100%">&nbsp;</td></tr>
<tr>
<td align=center valign=top>
<table cellpadding=0 cellspacing=0 width="100%" border="0" ID="Table2">
<tr>
<td width="100%" valign=top>
<table width="650" cellspacing="1" cellpadding="2" align="center" border="0" bgcolor="#ffffff" ID="Table3">
<tr><td height=50 nowrap colspan=2></td></tr>
<tr><td align="center" colspan=2>תאריך עדכון <%=QDaysUpdate%></td> </tr>

<tr><td align="center" colspan=2>תאריך שליחה אחרון <%=LastDateSend%></td></tr>
<tr><td height=50 nowrap colspan=2></td></tr>

<tr><td align="center" colspan=2 height="10" nowrap > (לאחר תאריך שליחה אחרון) שליחת קוד אימות אחת לכמות הימים</td></tr>

 <tr><td align="center" colspan=2><input type="text" name="datesend" id="datesend" value="<%=datesend%>"  style="width:50px"></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
<tr><td align="center" >
<input type=submit value="אישור" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1></td></tr>
</table>
</td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu" height="550">
<table cellpadding=1 cellspacing=1 width="100%" ID="Table4" border=0>
<tr><td height=50></td></tr>
<%if false then%>
<tr><td align="<%=align_var%>" colspan=2 height="550" nowrap valign=top>
 <a class="button_edit_1" style="width: 105px;"   href="javascript:void window.open('/pegasus/SendCodebySMS.aspx', 'DateSendCode' , 'scrollbars=1,toolbar=0,top=150,left=50,width=600,height=200,align=center,resizable=1');"><span id=word6 name=word6>שליחת קוד אימות</span></a>
 </td></tr>
<%end if%>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table>
</td></tr></table>
</form>
</BODY>
</HTML>

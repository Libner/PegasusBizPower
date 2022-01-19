<!--#include file="../include/reverse.asp"-->
<!--#include file="../include/connect.inc"-->

<%'on error resume next
if Request.Form("username")<>nil then
	username=Request.Form("username")
	password=Request.Form("password")
else if session("admin.username") <> nil then	
	username=session("admin.username")
	password=session("admin.password")
else 
	Response.Redirect("default.asp")	
end if
end if

sqlstring="SELECT * from workers where loginName='" & userName & "' and password='"& password &"' "
set worker=con.Execute(sqlstring)
if worker.EOF then%>
<HTML>
<p><center><font color="red" size="4"><b>You are not authorized to enter the SITE staff zone !!!</b></font></center>
<%else
  session("admin.username")=username
  session("admin.password")=password
%>
<% Set publ=con.Execute("SELECT * FROM Publication_Categories")
%>
<HEAD>
<META charset="visual">
<TITLE>Administration</TITLE>
<link rel="stylesheet" type="text/css" href="../leumicard.css">

</HEAD>
<BODY>
<center>

<table border="1" bordercolor="#afafaf" width="70%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" bgcolor="white"><p align="center"><strong><big><big>LeumiCard רתא לוהינ ןונגנמ</big></big></strong></td>
  </tr>
  <tr>
    <td width="100%" bgcolor="#ffffff">


<div align="center">
<table width="100%" cellspacing="4" cellpadding="1" >
  <!--
  <tr >
    <td align="center" valign="middle" nowrap><font size="3"><strong><a class="admin" href="news/admin.asp">תושדחו תועדוה</a></strong></font></td>
  </tr>
-->
  <tr >
    <td align="center" valign="middle"><font size="3"><strong><a class="admin" href="choose_e.asp"><span>LeumiCard - English</span></a></strong></font></td>
  </tr>

<%set maincat=con.execute("SELECT Main_Category_ID,Main_Category_Name FROM Main_Categories ORDER BY Category_Order")
  do while not maincat.eof%>
  <tr >
    <td align="center" valign="middle" nowrap><font size="3"><strong><a class="admin" href="pages/admin.asp?maincat=<%=maincat("Main_Category_ID")%>"><%=reverse(maincat("Main_Category_Name"))%>&nbsp;ןונגנמ לוהינ</a></strong></font></td>
  </tr>
  <%maincat.movenext
  loop%>
<!--
    <tr >
    <td align="center" valign="middle" nowrap><font size="3"><strong><a class="admin" href="sales/admin.asp">םיעצבמ</a></strong></font></td>
  </tr>
-->
</table>
</td>
</tr>
</table>
</div>
</center>
</BODY>
<%end if%>
</HTML>

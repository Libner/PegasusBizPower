<%@ Page Language="vb" AutoEventWireup="false" EnableViewState="false" Codebehind="Read_Excel.aspx.vb" Inherits="bizpower_pegasus.Read_Excel" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>Administration</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
     <LINK href="../../IE4.css" type="text/css" rel="STYLESHEET">
</head>
<body style="margin:0px;background-color:#E5E5E5" onload="window.focus();" style="margin: 0px;">
<table cellpadding="0" cellspacing="0" width="100%" border="0">
<tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td class="form_header" dir="rtl">&nbsp;<%=object_name%> יבוא&nbsp;</td></tr>
<tr><td>
<table cellpadding=5 cellspacing="0" width="100%" align="center">
<tr><td width="100%" colspan=2 height=30></td></tr>
<TR>
	<TD ALIGN="center" class="card"><span style="color:red"><%=sMessage%></span></TD>
</TR>
<tr><td align="center" width="100%" class="card" nowrap><a class="but_menu" style="width:120px;font-size:9pt;" href="#" onclick="window.opener.document.location.reload(); window.close();"><b>סגור חלון</b></a></td></tr>     
</table></td></tr>
</table>
</body>
</html>
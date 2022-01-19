<!--#INCLUDE file="..\..\netcom/reverse.asp"-->
<html>
<head>
<title>תחתון LOGO הוספת</title>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
</head>

<script language="JavaScript"><!--
function keepFocused() {
    self.focus();
    setTimeout('keepFocused()',30000);
}

keepFocused();

function assval(val)
{
	opener.document.all("logo_bottom").value = val;
}
//--></script>

<body bgcolor="#FFFFFF">
<div align="right">


<table width="80%" align=center border="0" cellpadding="0" cellspacing="2"  bordercolor="#E6E6E6">
   <tr>
      <td width="100%" align=center valign="top" bgcolor="#CCCCCC"><font size="3">ןותחת Logo תפסוה</font></td>
   </tr>
   <tr>
      <td width="100%" align=center valign="top" bgcolor="#E6E6E6">
		<table cellpadding="0" cellspacing="1" border="1" style="border-collapse:collapse;" bordercolor="white">
       <% set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
          foldString = Server.MapPath("../../templates/") 'server.MapPath is the current path we're in
          set fol=fs.GetFolder(foldString)
          for each thing in fol.Files
			If instr(1,thing.name,"asp") Then
			If instr(1,thing.name,"bottom") Then
			%>
			<tr><td align="left"><input onclick="assval('<%=thing.name%>')" type="radio" name="att" value="<%=thing.name%>">&nbsp;&nbsp;<a title="Show file" href="../../templates/<%=thing.name%>" target="_blank"><font color="navy"><b><%=thing.name%></b></font></a></td></tr>
			<%
			End If
			End If
	      next
	    %>
			<tr><td align="left"><input onclick="assval('')" type="radio" name="att" value="">&nbsp;&nbsp;<a title="Show file"><font color="navy"><b>Logo אלל</b></font></a></td></tr>	    
	    </table>
	</td>
   </tr>
   <tr>
      <td width="100%" align="center" bgcolor="#CCCCCC"><font size="3" face="Arial (Hebrew)"><input type="button" class="but"  value="אישור" onclick="javascript:window.close();" id=button1 name=button1></font></td>
   </tr>
</table>

</div>
</body>

</html>
 
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
if Request.Form("dateStart")<>"" then
start_date=Request.Form("dateStart")
else
start_date=date()-7
end if
if Request.Form("dateEnd")<>"" then
end_date=Request.Form("dateEnd")
else
end_date=Date()

end if



%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=157pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}

//-->
</SCRIPT>

</head>
<body>
<form id="FORM1"  name="FORM1"  method="post">
<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table1">
<tr><td width="100%" align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width="100%" align="<%=align_var%>">
  <%numOftab = 4%>
  <%numOfLink = 9%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title">&nbsp;</td></tr>
<tr>    
    <td width="100%" valign="top" align="center">
    <table dir="<%=dir_var%>" border="0" cellpadding="0" cellspacing="1" width="100%" ID="Table2">
    <tr>
    <td align="left" width="100%" valign=top >
    <table border=0 cellpadding=0 cellspacing=0 width=100%>
    <TR><td align=center><table border=0 cellpadding=0 cellspacing=0>
    	<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>"><input type=image src='../../images/calend.gif'   border=0  onclick='return popupcal(dateStart);' ID="Image1" NAME="Image1">&nbsp;<input dir="ltr" type="text" class="passw" size=12 id="dateStart" name="dateStart" value="<%=start_date%>" maxlength=10 readonly>&nbsp;</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id=word7 name=word7>תאריך מ</span>&nbsp;</TD>
		</TR>
		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>"><input type=image src='../../images/calend.gif'   border=0  onclick='return popupcal(dateEnd);' ID="Image2" NAME="Image2">&nbsp;<input dir="ltr" type="text" class="passw" size=12 id="dateEnd" name="dateEnd" value="<%=end_date%>" maxlength=10 readonly>&nbsp;</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id="word8" name=word8>תאריך עד</span>&nbsp;</TD>
		</TR>	
		
			<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><input type=submit value="דוח ביטוחים" ></TD>
		</TR>						

		</table></td></tr>
		<tr><td height=20></td></tr>
    
    <tr><TD align=center>
      <iframe name="frameGraph" id="frameGraph" src="ifrmInsureChart.aspx?sdate=<%=start_date%>&enddate=<%=end_date%>" 	ALLOWTRANSPARENCY=true height=600 width="100%" marginwidth="0" marginheight="0" hspace="0" vspace="0" scrolling="0" frameborder="0"></iframe>
    </TD></tr>
    </table>
      </td>
    </tr></table>
	</td></tr></table>
	</form>
</body>
</html>
<%set con=Nothing%>
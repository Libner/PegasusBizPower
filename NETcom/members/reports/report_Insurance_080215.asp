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
function report_open(fname){

		formsearch.action=fname;
		formsearch.target = '_blank';
		formsearch.submit();
	
}
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
	<FORM action="" method=POST id=formsearch name=formsearch target=_self>
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
    <TR><td align=center>	<table border="0" cellspacing="1" width=450 align=center bgcolor="#646E77" cellpadding="4" dir="<%=dir_var%>" ID="Table3">

    	<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>"><input type=image src='../../images/calend.gif'   border=0  onclick='return popupcal(dateStart);' ID="Image1" NAME="Image1">&nbsp;<input dir="ltr" type="text" class="passw" size=12 id="dateStart" name="dateStart" value="<%=start_date%>" maxlength=10 readonly>&nbsp;</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id=word7 name=word7>תאריך מ</span>&nbsp;</TD>
		</TR>
		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>"><input type=image src='../../images/calend.gif'   border=0  onclick='return popupcal(dateEnd);' ID="Image2" NAME="Image2">&nbsp;<input dir="ltr" type="text" class="passw" size=12 id="dateEnd" name="dateEnd" value="<%=end_date%>" maxlength=10 readonly>&nbsp;</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id="word8" name=word8>תאריך עד</span>&nbsp;</TD>
		</TR>	
		
			<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('ifrmInsureChart.aspx');"  style="width:250">דוח ביטוחים גרפי</a></TD>
		</TR>						
	<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('report_InsuranceExcel.asp');"  style="width:250">דוח מסכם באקסל</a></TD>
		</TR>						

		</table></td></tr>
		<tr><td height=20></td></tr>
  
    </table>
      </td>
    </tr></table>
	</td></tr></table>
	</form>
</body>
</html>
<%set con=Nothing%>
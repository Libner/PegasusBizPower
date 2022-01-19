<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%If trim(lang_id) = "1" Then
		title = " דוח לקוחות "  :  title2 = " דוח אנשי קשר "  :  	page_title = "דוחות לקוחות"
	 Else
		title = " Customers Report "  :  title2 = " Contacts Report "  :  	page_title = "Customers Reports"
	 End If	 %>
<%	 sqlPer="SELECT  bar_id  FROM bar_users WHERE (user_id =" & UserID &") and bar_id=47 and is_visible=1"
 set rs_Per = con.getRecordSet(sqlPer)
			if not rs_Per.eof then	
			SMS="1"
			else
			SMS="0"
		end if
		set rs_Per= nothing   %>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
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

function report_open(fname)
{
	form_report.action=fname;
	form_report.target = '_blank';
	form_report.submit();		
}
//-->
</SCRIPT>
</head>
<body topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 bgcolor="#E5E5E5">
<!--#include file="../../logo_top.asp"-->
<table cellpadding=0 cellspacing=0 width=100%>
<tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="0" >
<%numOftab = 0%>
<%numOfLink = 3%>
<%topLevel2 = 13 'current bar ID in top submenu - added 03/10/2019%>

<tr><td width=100% align="right"><!--#include file="../../top_in.asp"--></td></tr>
</table>
</td></tr>
<tr><td>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<tr><td class="page_title" dir=rtl>&nbsp;&nbsp;</td></tr>
<tr><td height=15 nowrap></td></tr>
 <tr>    
    <td width="100%" valign="top" align="center">
    <table width="100%" cellspacing="0" cellpadding="3" align=center border="0" bgcolor="#ffffff">
      <tr>
        <td width="100%" align=center valign=top >
	    <!-- start code -->
		<FORM action="" method=POST id=form_report name=form_report>
		<table border="0" cellspacing="1" width=450 align=center bgcolor="#646E77" cellpadding="4">
			<tr id=but_tr>
				<td bgcolor="#dbdbdb" align="center" colspan="2">
					<TABLE WIDTH=100% align=center BORDER=0 CELLSPACING=5 CELLPADDING=0>
					<TR>
						<TD width=20%>&nbsp;</TD>
						<TD align=center width=60%><A class="but_menu" href="#" onclick="report_open('export_excel_csv.asp');" dir="<%=dir_var%>"><%=title%></a></TD>
						<TD width=20%>&nbsp;</TD>
					</TR>
					<TR>
						<TD width=20%>&nbsp;</TD>
						<TD align=center width=60%><A class="but_menu" href="#" onclick="report_open('export_excel_contacts.asp');" dir="<%=dir_var%>"><%=title2%></a></TD>
						<TD width=20%>&nbsp;</TD>
					</TR>	
					<TR>
						<TD width=20%>&nbsp;</TD>
						<TD align="center" width="60%"><A class="but_menu" href="#" onclick="report_open('export_excel_dockets.asp');" dir="<%=dir_var%>">דוח ניצול נקודות</a></TD>
						<TD width=20%>&nbsp;</TD>
					</TR>	
					<tr><td colspan=3 height=10></td></tr>
						</TABLE>
				</td>
			</tr>
		</table>
	</FORM>
  <!-- End main content -->
</td></tr>
</table>   
</body>
<%set con=nothing%>
</html>

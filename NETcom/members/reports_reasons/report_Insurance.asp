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
<script language="javascript" type="text/javascript" src="../CalendarPopup.js"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function callCalendar(pf,pid)
	{

	cal1xx.select(pf,pid,'dd/MM/yyyy')
	}
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
<%topLevel2 = 59 'current bar ID in top submenu - added 03/10/2019%>
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
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>">
									<a href='' onclick='callCalendar(document.formsearch.dateStart,"AsdateStart");return false;' id='AsdateStart'>
					<image src='../../images/calend.gif' border=0 ></a>&nbsp;<input dir='ltr' type='text' class='passw' size=10 id="dateStart" name='dateStart' value='<%=start_date%>' maxlength=10 readonly>&nbsp;
			</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id=word7 name=word7>תאריך מ</span>&nbsp;</TD>
		</TR>
		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>">
										<a href='' onclick='callCalendar(document.formsearch.dateEnd,"AsdateEnd");return false;' id="AsdateEnd">
					<image src='../../images/calend.gif' border=0 ></a>&nbsp;<input dir='ltr' type='text' class='passw' size=10 id="dateEnd" name='dateEnd' value='<%=end_date%>' maxlength=10 readonly>&nbsp;
			</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id="word8" name=word8>תאריך עד</span>&nbsp;</TD>
		</TR>	
		
			<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('ifrmInsureChart.aspx');"  style="width:250">דוח ביטוחים גרפי</a></TD>
		</TR>						
	<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('report_InsuranceExcel.asp');"  style="width:250">דוח מסכם באקסל</a></TD>
		</TR>		
		
			<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('report_OnlyInsuranceExcel.asp');"  style="width:250">דוח ביטוחים שנשלחו באקסל</a></TD>
		</TR>						
				

		</table></td></tr>
		<tr><td height=20></td></tr>
  
    </table>
      </td>
    </tr></table>
	</td></tr></table>
	</form>
			 	<SCRIPT LANGUAGE="JavaScript">document.write(getCalendarStyles());</SCRIPT>
			<SCRIPT LANGUAGE="JavaScript" type="text/javascript">
            <!--
            var cal1xx = new CalendarPopup('CalendarDiv');
                cal1xx.showNavigationDropdowns();
                cal1xx.yearSelectStart
                cal1xx.offsetX = -50;
                cal1xx.offsetY = 0;
            //-->
			</SCRIPT>
					<DIV ID='CalendarDiv' name='CalendarDiv' STYLE='VISIBILITY:hidden;POSITION:absolute;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>


</body>
</html>
<%set con=Nothing%>
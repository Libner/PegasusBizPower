<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
 <%numOftab = 82
 numOfLink = 2%>
<%topLevel2 = 86 'current bar ID in top submenu - added 03/10/2019%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript" src="../CalendarPopup.js"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function report_open(fname){
	formsearch.action=fname;
		formsearch.target = '_blank';
		formsearch.submit();
	
}

//-->
</SCRIPT>
<script>
var oPopup = window.createPopup();
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
function callCalendar(pf,pid)
	{

	cal1xx.select(pf,pid,'dd/MM/yyyy')
	}

</script>
</head>
<body topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 bgcolor="#E5E5E5">
	<FORM action="" method=POST id="Form1" name=formsearch target=_self>

<!--#include file="../../logo_top.asp"-->
<table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>" ID="Table1">
<tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table2">
<tr><td width=100% align="<%=align_var%>"><!--#include file="../../top_in.asp"--></td></tr>
</table>
</td></tr>
<tr><td>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" ID="Table3">
<tr><td class="page_title" dir="<%=dir_obj_var%>">&nbsp;</td></tr>

<tr><td height=15 nowrap></td></tr>
 <tr>    
    <td width="100%" valign="top" align="center">
    <table width="100%" cellspacing="0" cellpadding="3" align=center border="0" bgcolor="#ffffff" ID="Table4">
      <tr>
        <td width="100%" align=center valign=top >
<!-- start code -->
	<FORM action="" method=POST id=formsearch name=formsearch target=_self>
	<table border="0" cellspacing="1" width=450 align=center bgcolor="#646E77" cellpadding="4" dir="<%=dir_var%>" ID="Table5">
	

	
	
									
		<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('reportExcel.asp?type_reports=1');" style="width:250"><span id="word12" name=word12>טיולים שעומדים לצאת בשבועיים הקרובים וטרם נקבע להם בריף</span></a></TD>
		</TR>						
		<TR>						
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('reportExcel.asp?type_reports=2');" style="width:250"><span id="word13" name=word13>טיולים שנמצאים כרגע בחו"ל ולא בוצעו שיחות למדריך</span></a></TD>
		</TR>	
		<TR>						
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('reportExcel.asp?type_reports=3');" style="width:250"><span id="Span1" name=word13>טיולים שחזרו וטרם תואם מפגש חזרה עם מדריך</span></a></TD>
		</TR>	
	<TR>						
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('reportExcel.asp?type_reports=4');" style="width:250"><span id="Span2" name=word13>טיולים שעומדים לצאת בשבועיים הקרובים וחסרה הזמנה מסימולטני</span></a></TD>
		</TR>	

		<TR>						
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('reportExcel.asp?type_reports=5');" style="width:250"><span id="Span3" name=word13>בריף שמתקיים ביומיים הקרובים ולא הופקו שוברים</span></a></TD>
		</TR>	
	
	</TABLE>			
	<br clear=all>
		<table border="0" cellspacing="1" width=450 align=center bgcolor="#646E77" cellpadding="4" dir="<%=dir_var%>" ID="Table7">

	<TR>						
			<TD align=center  colspan=2 height=20 bgcolor="#ffffff"></td></tr>
	<TR>						
			<TD align=center  colspan=2 bgcolor="#ffffff"><span style="font-size:14pt">דו"ח תיעודים שיחות</span></TD>
		</TR>	
			<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">		
		<select name="user_id" id="user_id" class="norm" dir="<%=dir_obj_var%>" style="width:310">		
		<option value="" id=word4 name=word4>כל העובדים</option>		
		<%  sqlstr = "Select USER_ID, FIRSTNAME + ' ' + LASTNAME FROM USERS WHERE ORGANIZATION_ID = " & OrgID			
		    sqlstr = sqlstr & " Order BY FIRSTNAME + ' ' + LASTNAME "
			set rs_comp = con.getRecordSet(sqlstr)
			While not rs_comp.eof
		%>
			<option value="<%=rs_comp(0)%>" <%If trim(userID) = trim(rs_comp(0)) Then%> selected <%End if%>><%=rs_comp(1)%></option>			
		<%
			rs_comp.moveNext
			Wend
			set rs_comp = Nothing
		%>
		</select>
		</TD>
		<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<span id="word5" name=word5>עובד</span>&nbsp;</td>
	</TR>
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
						<image src='../../images/calend.gif' border=0 ></a>&nbsp;<input dir="ltr" type="text" class="passw" size=10 id="dateEnd" name="dateEnd" value="<%=end_date%>" maxlength=10 readonly>&nbsp;
</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id="word8" name=word8>תאריך עד</span>&nbsp;</TD>
		</TR>	
		<tr id=but_tr>
			<td bgcolor="#dbdbdb" align="center" colspan="2">
				<TABLE WIDTH=100% align=center BORDER=0 CELLSPACING=5 CELLPADDING=0 ID="Table6">
				 	<TR>
						<td width=20%>&nbsp;</td>
						<TD align=center width=60%><A class="but_menu" href="#" onclick="report_open('reportMessages.asp');"><span id=word9 name=word9>הצג דוח</span></a></TD>
						<td width=20%>&nbsp;</td>
					</TR>																								
					</table></td></tr>
</table>
	
</FORM>
  <!-- End main content -->
</td></tr>
</table>	 
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
<%set con=nothing%>
</html>

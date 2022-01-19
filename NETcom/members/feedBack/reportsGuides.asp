<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
 <%numOftab = 77
 numOfLink = 2
 topLevel2 = 94 'current bar ID in top submenu - added 03/10/2019
 dtCurrentDate=Now()
 %>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript" src="../CalendarPopup.js"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function report_open(fname){
if (document.getElementById("guide_id").value=='0')
{
  alert("אנא בחר מדריך");
	return false;
	}
	formsearch.action=fname;
		formsearch.target = '_blank';
		formsearch.submit();
	
}
function report_openPdf(fname)
{
if (document.getElementById("guide_id").value=='0')
{
  alert("אנא בחר מדריך");
	return false;
	}
	
var wFile
wFile="https://pegasusisrael.co.il/biz_form/" + fname+"?guide_id="+document.getElementById("guide_id").value +"&y="+ document.getElementById("currentYear").value
//wFile="http://192.168.0.92/pegasusNew/biz_form/" + fname+"?guide_id="+document.getElementById("guide_id").value +"&y="+ document.getElementById("currentYear").value

//alert(wFile)
window.open(wFile,"pdfG")
}

//-->
</SCRIPT>
<script>
	function report_mail()
	{
	if (document.getElementById("guide_id").value=='0')
		{
		 alert("אנא בחר מדריך");
		 return false;
		}
	//alert (document.getElementById("guide_id").value)
	//alert (document.getElementById("currentYear").value)
	//window.open("SEndGuidePdfReportMail.aspx?gId="+document.getElementById("guide_id").value+"&gyear="+document.getElementById("currentYear").value,"_send","width=700,height=500")
	//window.open("http://192.168.0.92/pegasusNew/guides/SendGuidePdfReportMail.aspx?uId="+<%=Request.Cookies("bizpegasus")("UserId")%> +"&gId="+document.getElementById("guide_id").value+"&gyear="+document.getElementById("currentYear").value,"_send","width=700,height=500")
 	window.open("https://pegasusisrael.co.il/guides/SendGuidePdfReportMail.aspx?uId="+<%=Request.Cookies("bizpegasus")("UserId")%> +"&gId="+document.getElementById("guide_id").value+"&gyear="+document.getElementById("currentYear").value,"_send","width=700,height=500")

}
</script>
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
<table border="0" width="100%" cellspacing="0" cellpadding="0"  ID="Table2">
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
	
	<br clear=all>
		<table border="0" cellspacing="1" width=450 align=center bgcolor="#646E77" cellpadding="4" dir="<%=dir_var%>" ID="Table7">

		<tr>
			<td width="80%" bgcolor="#DBDBDB" align="right" dir="rtl">				
			<select dir="rtl" name="guide_id" class="app" id="guide_id" style="width:320">
		<option value="0"><%=String(28,"-")%> כל המדריכים <%=String(28,"-")%></option>	
			<%sqlstrGuide = "SELECT Guide_Id, (Guide_FName + ' - ' + Guide_LName) as Guide_Name  FROM Guides  where Guide_Vis=1 ORDER BY Guide_FName,Guide_LName"

	set rs_Guide= conPegasus.Execute(sqlstrGuide)
   do while not rs_Guide.EOF 
		Guide_Id =rs_Guide("Guide_Id")
		Guide_Name =rs_Guide("Guide_Name")%>
	<OPTION value="<%=Guide_Id%>" <%If trim(pGuideId) = trim(Guide_Id) Then%> selected <%End If%>><%=Guide_Name%> </OPTION>
			<%rs_Guide.moveNext
		loop
		rs_Guide.close
		set rs_Guide=Nothing	%>
    		</select>
			</td>
			<td width="20%" class="subject_form" nowrap="" align="right">&nbsp;<span id="Span1" name="word4"><!--טופס-->מדריך</span>&nbsp;</td>
		</tr>
			<tr>
			<td width="80%" bgcolor="#DBDBDB" align="right" dir="rtl">	
			<SELECT NAME="currentYear" CLASS="norm" ID="currentYear" dir="<%=dir_var%>" style="width:60" >
	      <% For counter = -10 to 0 %>
	        <OPTION VALUE="<%=Year(dtCurrentDate)+counter%>" <% If (DatePart("yyyy", dtCurrentDate) = Year(dtCurrentDate)) Then Response.Write "SELECTED"%>><%=Year(dtCurrentDate)+counter%></OPTION>
	      <% Next %>
	      </SELECT>
			</td>
			<td width="20%" class="subject_form" nowrap="" align="right">&nbsp;<span id="Span4" name="word4">שנת קלנדרי</span>&nbsp;</td>
	

		<tr id=but_tr>
			<td bgcolor="#dbdbdb" align="center" colspan="2">
				<TABLE WIDTH=100% align=center BORDER=0 CELLSPACING=5 CELLPADDING=0 ID="Table6">
				 	<TR>
						<td width=20%>&nbsp;</td>
						<TD align=center width=60%><A class="but_menu" href="#" onclick="report_open('graphByGuide.aspx');"><span id=word9 name=word9>הצג דוח גרפי שנתי למדריך</span></a></TD>
						<td width=20%>&nbsp;</td>
					</TR>	
						<TR>
						<td width=20%>&nbsp;</td>
						<TD align=center width=60%><A class="but_menu" href="#" onclick="report_open('reportExcelGuides.asp');"><span id="Span3" name=word9>Excel הצג דוח  שנתי למדריך</span></a></TD>
						<td width=20%>&nbsp;</td>
					</TR>						
					<TR>
						<td width=20%>&nbsp;</td>
						<TD align=center width=60%><A class="but_menu" href="#" onclick="report_openPdf('reportScreenGuidesPdf.aspx');"><span id="Span2" name=word9>PDF הצג דוח  שנתי למדריך</span></a></TD>
						<td width=20%>&nbsp;</td>
					</TR>	
					<%if false then%><TR>
						<td width=20%>&nbsp;</td>
						<TD align=center width=60%><A class="but_menu" href="#" onclick="report_openPdf('https://www.pegasusisrael.co.il/biz_form/reportScreenGuidesPdfEx.aspx');"><span id="Span6" name=word9>PDF -excel הצג דוח  שנתי למדריך</span></a></TD>
						<td width=20%>&nbsp;</td>
					</TR>	
<%end if%>
					<TR>
						<td width=20% colspan=2 height=20>&nbsp;</td></tr>
						<TR>
						<td width=20%>&nbsp;</td>
						<TD align=center width=60%><A class="but_menu" style="background-color:#38A638" href="#" onclick="report_mail();"><span id="Span5" name=word9>שלח במייל דוח  שנתי למדריך</span></a></TD>
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

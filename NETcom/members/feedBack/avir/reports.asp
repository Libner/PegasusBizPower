<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
 <%numOftab = 77
 numOfLink = 1
 start_date="01/11/2016"
 end_date=FormatDateTime(Now(),2)
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
<table border="0" width="100%" cellspacing="0" cellpadding="0" bgcolor="#060165" ID="Table2">
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
			<select dir="rtl" name="guide_id" class="app" id="guide_id" style="width:320;height:120" multiple >
		<option value="0"><%=String(28,"-")%> כל המדריכים <%=String(28,"-")%></option>	
			<%sqlstrGuide = "SELECT Guide_Id, (Guide_FName + ' - ' + Guide_LName) as Guide_Name  FROM Guides  where Guide_Vis=1 ORDER BY  Guide_FName,Guide_LName"

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
			<td width="20%" class="subject_form" nowrap align="right">&nbsp;<span id="Span1" name="word4"><!--טופס-->מדריך</span>&nbsp;</td>
		</tr>
		<tr>
			<td width="80%" bgcolor="#DBDBDB" align="right" dir="rtl">				
			<select dir="rtl" name="dep_Code" class="app" id="dep_Code" style="width:320;height:120" multiple>
		<option value="0"><%=String(28,"-")%> כל הטיולים <%=String(28,"-")%></option>	
			<%'sqlstrDeparture = "SELECT  Departure_Id, Departure_Code  FROM Tours_Departures where datediff(yy,Departure_Date,GETDATE())<=1	 ORDER BY Departure_Date"
			sqlstrDeparture = "SELECT  distinct Departure_Code  FROM Tours_Departures where datediff(yy,Departure_Date,GETDATE())<=1	 ORDER BY Departure_Code"

	set rs_Dep= conPegasus.Execute(sqlstrDeparture)
   do while not rs_Dep.EOF 
		'Departure_Id =rs_Dep("Departure_Id")
		Departure_Code =rs_Dep("Departure_Code")%>
	<OPTION value="<%=Departure_Code%>" <%If trim(pDepId) = trim(Departure_Code) Then%> selected <%End If%>><%=Departure_Code%> </OPTION>
			<%rs_Dep.moveNext
		loop
		rs_Dep.close
		set rs_Dep=Nothing	%>
    		</select>
			</td>
			<td width="20%" class="subject_form" nowrap align="right">&nbsp;<span id="Span2" name="word4">קוד טיול</span>&nbsp;</td>
		</tr>
	
		
<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>">
	<a href='' onclick='callCalendar(document.formsearch.dateStart,"AsdateStart");return false;' id='AsdateStart'>
					<image src='../../images/calend.gif' border=0 ></a>&nbsp;<input dir='ltr' type='text' class='passw' size=10 id="dateStart" name='dateStart' value='<%=start_date%>' maxlength=10 readonly>&nbsp;
			
			</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id=word7 name=word7>- תאריך יציאה מ</span>&nbsp;</TD>
		</TR>
		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>">	
				<a href='' onclick='callCalendar(document.formsearch.dateEnd,"AsdateEnd");return false;' id="AsdateEnd">
						<image src='../../images/calend.gif' border=0 ></a>&nbsp;<input dir="ltr" type="text" class="passw" size=10 id="dateEnd" name="dateEnd" value="<%=end_date%>" maxlength=10 readonly>&nbsp;
</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id="word8" name=word8>- תאריך יציאה עד</span>&nbsp;</TD>
		</TR>	
					<tr>
			<td width="80%" bgcolor="#DBDBDB" align="right" dir="rtl">				
			<select dir="rtl" name="ser_id" class="app" id="ser_id" style="width:320;height:120" multiple>
		<option value="0"><%=String(28,"-")%> כל הסדרות <%=String(28,"-")%></option>	
			<%sqlstrSer = "SELECT  Series_Id, Series_Name  FROM Series  ORDER BY Series_Name"

	set rs_Ser= con.getRecordSet(sqlstrSer)
   do while not rs_Ser.EOF 
		Series_Id =rs_Ser("Series_Id")
		Series_Name =rs_Ser("Series_Name")%>
	<OPTION value="<%=Series_Id%>" <%If trim(pSerId) = trim(Series_Id) Then%> selected <%End If%>><%=Series_Name%> </OPTION>
			<%rs_Ser.moveNext
		loop
		rs_Ser.close
		set rs_Ser=Nothing	%>
    		</select>
			</td>
			<td width="20%" class="subject_form" nowrap align="right">&nbsp;<span id="Span3" name="word4">סדרה אליה משויך הטיול</span>&nbsp;</td>
		</tr>
		
		<tr>
			<td width="80%" bgcolor="#DBDBDB" align="right" dir="rtl">				
			<select dir="rtl" name="country_id" class="app" id="country_id" style="width:320;height:120" multiple>
		<option value="0"><%=String(28,"-")%> כל המדינות <%=String(28,"-")%></option>	
			<%sqlstrCountry = "SELECT Country_Id, Country_Name from Countries  order by Country_Name"

	set rs_Country= conPegasus.Execute(sqlstrCountry)
   do while not rs_Country.EOF 
		Country_Id =rs_Country("Country_Id")
		Country_Name =rs_Country("Country_Name")%>
	<OPTION value="<%=Country_Id%>" <%If trim(pCountryId) = trim(Country_Id) Then%> selected <%End If%>><%=Country_Name%> </OPTION>
			<%rs_Country.moveNext
		loop
		rs_Country.close
		set rs_Country=Nothing	%>
    		</select>
			</td>
			<td width="20%" class="subject_form" nowrap align="right">&nbsp;<span id="Span4" name="word4">מדינה אליה משויך הטיול</span>&nbsp;</td>
		</tr>	
<tr>
			<td width="80%" bgcolor="#DBDBDB" align="right" dir="rtl">	
			<table border=0 cellpadding=2 cellspacing=2>
			<tr><td>מ- &nbsp;&nbsp;<input dir="ltr" type="text" class="passw" size=3 id="FromGrade" name="FromGrade" value="0" maxlength=10></td></tr>
			<tr><td>עד-&nbsp; <input dir="ltr" type="text" class="passw" size=3 id="ToGrade" name="ToGrade" value="100" maxlength=10></td></tr>
		
			</table>
				</td>
			<td width="20%" class="subject_form" nowrap align="right">&nbsp;<span id="Span5" name="word4">ציון של טיול מינימלי</span>&nbsp;</td>
		</tr>			
		<tr id=but_tr>
			<td bgcolor="#dbdbdb" align="center" colspan="2">
				<TABLE WIDTH=100% align=center BORDER=0 CELLSPACING=5 CELLPADDING=0 ID="Table6">
				 	<TR>
						<td width=20%>&nbsp;</td>
						<TD align=center width=60%><A class="but_menu" href="#" onclick="report_open('reportExcel.asp');"><span id=word9 name=word9>הצג דוח אקסל</span></a></TD>
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

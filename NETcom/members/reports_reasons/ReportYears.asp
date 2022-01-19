<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
	type_reports = trim(Request("type"))
	 numOftab = 1
     numOfLink =8
	topLevel2 = 92 'current bar ID in top submenu - added 03/10/2019
	
	start_date = "1/" & Month(Date()) & "/" & Year(date())
	end_date = getDaysInMonth(month(start_date),year(start_date)) & "/" &  Month(Date()) & "/" & Year(date())
	companyID = trim(Request("company_ID"))
	User_ID = trim(Request("user_id"))
	quest_id = trim(Request("quest_id"))
	projectId = trim(Request("project_Id"))
	If Request.Form("dateStart") <> nil Then
		start_date = Request.Form("dateStart")
	End If
	If Request.Form("dateEnd") <> nil Then
		end_date = Request.Form("dateEnd")
	End If
%>
<%   
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 25 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing	
%> 
<%
function getDaysInMonth(strMonth,strYear)
dim strDays	 
    Select Case cint(strMonth)
        Case 1,3,5,7,8,10,12:
			strDays = 31
        Case 4,6,9,11:
		strDays = 30
        Case 2:
		if  ((cint(strYear) mod 4 = 0  and  _
                 cint(strYear) mod 100 <> 0)  _
                 or ( cint(strYear) mod 400 = 0) ) then
		  strDays = 29
		else
		  strDays = 28 
		end if	
    End Select 
    
    getDaysInMonth = strDays

end function%>
	
<html>
<head>
<!--#include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript" src="../CalendarPopup.js"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--


function report_open(fname){
//var dateold =New Date(document.getElementById("dateStart").value)
//alert(dateold)
//document.getElementById("dateStart").value
  //var datenew=document.getElementById("dateEnd").value
//alert(dateold.getMonth())
//DateDiff()
//alert(pyear)
//alert(document.getElementById("dateStart").value)
//alert(document.getElementById("dateEnd").value)
	if(isNaN(eval(window.document.all("country_id").value)) == false)
	{
		formsearch.action=fname;
		formsearch.target = '_blank';
		formsearch.submit();
		
		start_tr.style.display='none';

	}else{
   <%
     If trim(lang_id) = "1" Then
        str_alert = "! נא לבחור  יעד"
     Else
		str_alert = "Please choose a form !"
     End If   
    %>
	window.alert("<%=str_alert%>");
	window.document.all("country_id").focus();
	
	}
}
function callCalendar(pf,pid)
	{

	cal1xx.select(pf,pid,'dd/MM/yyyy')
	}


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


//-->
</SCRIPT>
</head>
<body topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 bgcolor="#E5E5E5">
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
	<input type=hidden name="type" id="type" value="<%=type_reports%>">
	<table border="0" cellspacing="1" width=450 align=center bgcolor="#646E77" cellpadding="4" dir="<%=dir_var%>" ID="Table5">
		<tr>
			<td bgcolor="#dbdbdb" align="center" colspan="2" height="15" nowrap>
				<input type="hidden" name="report" value="yes" ID="Hidden1">
			</td>
		</tr>
	
	<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">
		<select name="country_id" id="country_id" class="norm" dir="<%=dir_obj_var%>" multiple="multiple" style="width:320" size="15">
		<%If trim(lang_id) = "1" Then%>
		<option value="0"><%=String(28,"-")%> כל המדינות <%=String(28,"-")%></option>	
		<%Else%>
		<option value="0"><%=String(26,"-")%> All Countries <%=String(26,"-")%></option>	
		<%End If%>
		<%sqlstr = "Select Country_Id, Country_Name FROM Countries "
	      sqlstr = sqlstr & " Order BY Country_Name "
		set rs_comp=con.GetRecordSet(sqlstr)
		If Not rs_comp.eof Then
			arr_comp = rs_comp.getRows()
		End If
		Set rs_comp = Nothing
		If isArray(arr_comp) Then
		For cc=0 To Ubound(arr_comp,2)%>		
		<option value="<%=arr_comp(0,cc)%>"<%If trim(country_id) = trim(arr_comp(0,cc)) Then%> selected <%End if%>><%=arr_comp(1,cc)%></option>			
		<%  Next
		   End If%>		
		</select>
		</TD>
		<td width="20%" class="subject_form" align="<%=align_var%>" valign=top>&nbsp;יעד&nbsp;</td>
	</TR>

		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>">
			<a href='' onclick='callCalendar(document.formsearch.dateStart,"AsdateStart");return false;' id='AsdateStart'>
					<image src='../../images/calend.gif' border=0 ></a>&nbsp;<input dir='ltr' type='text' class='passw' size=10 id="dateStart" name='dateStart' value='<%=start_date%>' maxlength=10 readonly>&nbsp;
</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id=word7 name=word7><!--- תאריך מ--><%=arrTitles(7)%></span>&nbsp;</TD>
		</TR>
		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>">
		<a href='' onclick='callCalendar(document.formsearch.dateEnd,"AsdateEnd");return false;' id="AsdateEnd">
						<image src='../../images/calend.gif' border=0 ></a>&nbsp;<input dir="ltr" type="text" class="passw" size=10 id="dateEnd" name="dateEnd" value="<%=end_date%>" maxlength=10 readonly>&nbsp;
			</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id="word8" name=word8><!--- תאריך עד--><%=arrTitles(8)%></span>&nbsp;</TD>
		</TR>	
			
		<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('reportYearEx.asp');" style="width:250"><span id="word12" name=word12>EXCEL דוח</span></a></TD>
		</TR>						
		<TR>						
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('reportYearGraph.aspx');" style="width:250"><span id="word13" name=word13>דוח גרפי</span></a></TD>
		</TR>
	

	</TABLE>			

</FORM>
  <!-- End main content -->
</td></tr>
</table>	 
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

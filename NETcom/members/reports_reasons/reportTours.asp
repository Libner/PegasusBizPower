<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
	type_reports = trim(Request("type"))
	 numOftab = 1
     numOfLink =  6
	topLevel2 = 71 'current bar ID in top submenu - added 03/10/2019

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
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	<script src="//code.jquery.com/jquery-1.10.2.js"></script>
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
<link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">	
 <script src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>
 <script type='text/javascript' src="https://jquery-ui.googlecode.com/svn-history/r3982/trunk/ui/i18n/jquery.ui.datepicker-he.js"></script>



<script>

   
  $(function() {
      $.datepicker.setDefaults($.datepicker.regional["he"]);
      $.datepicker.setDefaults({ dateFormat: 'dd/mm/yy' });
	$('#dateStart').datepicker();
	$('#dateEnd').datepicker();
	
	$('#icDatePicker').on('click', function() {
      $('#dateStart').datepicker('show');
      return false
   });
   	$('#icDatePickerEnd').on('click', function() {
      $('#dateEnd').datepicker('show');
      return false
   });
});

  </script>

<SCRIPT  LANGUAGE=javascript>
<!--
function report_open(fname){
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
	<tr>
			<td width="80%" bgcolor="#DBDBDB" align="right" dir="rtl">				
			<select dir="rtl" name="quest_id" class="app" id="quest_id" style="width:320">
			<option value="0" >----------------------------- בחר טופס -----------------------------</option>
					<option value="17857"> טופס מכתב תודה </option>
			    	<option value="17001"> טופס פניות הציבור </option>
					<option value="17057"> טופס תלונת לקוח </option>
    		</select>
			</td>
			<td width="20%" class="subject_form" nowrap="" align="right">&nbsp;<span id="word4" name="word4"><!--טופס-->טופס</span>&nbsp;</td>
		</tr>

	<tr>
			<td width="80%" bgcolor="#DBDBDB" align="right" dir="rtl">				
			<select dir="rtl" name="guide_id" class="app" id="guide_id" style="width:320">
		<option value="0"><%=String(28,"-")%> כל המדריכים <%=String(28,"-")%></option>	
			<%sqlstrGuide = "SELECT Guide_Id, (Guide_FName + ' - ' + Guide_LName) as Guide_Name  FROM Guides  where Guide_Vis=1 ORDER BY Guide_LName, Guide_FName"

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
		set rs_comp=conPegasus.Execute(sqlstr)
		   do while not rs_comp.EOF 
		Country_Id =rs_comp("Country_Id")
		Country_Name =rs_comp("Country_Name")%>
	<OPTION value="<%=Country_Id%>"><%=Country_Name%> </OPTION>
			<%rs_comp.moveNext
		loop
		rs_comp.close
		set rs_comp=Nothing	%>
		</select>
		</TD>
		<td width="20%" class="subject_form" align="<%=align_var%>" valign=top>&nbsp;יעד&nbsp;</td>
	</TR>
		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>">
			<a  id="icDatePicker" href=""><img class="calendar_icon" src="../../images/calend.gif" class="iconButton" border=0></a>
			  <input  id="dateStart" type="text"  name="dateStart" dir="ltr" class="passw" size=10 value="<%=start_date%>">
		
	</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id=word7 name=word7><!--- תאריך מ--><%=arrTitles(7)%></span>&nbsp;</TD>
		</TR>
		<TR>		
				<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>">
			<a  id="icDatePickerEnd" href=""><img class="calendar_icon" src="../../images/calend.gif" class="iconButton" border=0></a>
			  <input  id="dateEnd" type="text"  name="dateEnd" dir="ltr" class="passw" size=10 value="<%=end_date%>">
		
	</TD>				
				<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id="word8" name=word8><!--- תאריך עד--><%=arrTitles(8)%></span>&nbsp;</TD>
		</TR>	
		<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('reportToursEx.asp');" style="width:250"><span id="word12" name=word12>EXCEL דוח</span></a></TD>
		</TR>						
	

	</TABLE>			

</FORM>
  <!-- End main content -->
</td></tr>
</table>	 
 </body>
<%set con=nothing%>
</html>

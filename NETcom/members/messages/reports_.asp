<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
	type_reports = trim(Request("type"))

	numOftab=49
	numOfLink=3
	start_date = "1/" & Month(Date()) & "/" & Year(date())
	end_date = DateAdd("d",30,start_date)
	User_ID = trim(Request("user_id"))
	quest_id = trim(Request("quest_id"))
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
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function report_open(fname){
	if(isNaN(eval(window.document.all("quest_id").value)) == false)
	{
		formsearch.action=fname;
		formsearch.target = '_blank';
		formsearch.submit();
		
		start_tr.style.display='none';

	}else{
   <%
     If trim(lang_id) = "1" Then
        str_alert = "! נא לבחור טופס"
     Else
		str_alert = "Please choose a form !"
     End If   
    %>
	window.alert("<%=str_alert%>");
	window.document.all("quest_id").focus();
	}
}
function report_openedText(fname){

		formsearch.action=fname;
		formsearch.target = '_self';

}
function report_openfr(fname){
	formsearch.action=fname;
	formsearch.target = '_blank';
	formsearch.submit();
}
function report_open_self(fname){
	if (window.document.all("quest_id").value != "0")
	{
		formsearch.action=fname;
		formsearch.target = '_self';
		formsearch.submit();
	}
}

var oPopup = window.createPopup();
function ProductDropDown(obj)
{
	oPopup.document.body.innerHTML = Product_Popup.innerHTML;
	var popupBodyObjProduct = oPopup.document.body;
	popupBodyObjProduct.style.border = "1px black solid"; 	    	 
	var dHeight = popupBodyObjProduct.scrollHeight;
	var dWidth = popupBodyObjProduct.scrollWidth;			
	oPopup.show(0, 21, 320, 108, obj); 
	popupBodyObjProduct = null;
	return true;   
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
		<select name="user_id" id="user_id" class="norm" dir="<%=dir_obj_var%>" style="width:320" onchange="document.formsearch.action='reports.asp';document.formsearch.target='_self';document.formsearch.submit();">		
		<option value=""><%=arrTitles(5)%></option>	
		<%  sqlstr = "Select USER_ID, FIRSTNAME + ' ' + LASTNAME FROM USERS WHERE ACTIVE=1 and ORGANIZATION_ID = " & OrgID			
		    sqlstr = sqlstr & " Order BY FIRSTNAME + ' ' + LASTNAME "
			set rs_user=con.GetRecordSet(sqlstr)
			If Not rs_user.eof Then
				arr_user = rs_user.getRows()
			End If
			Set rs_user = Nothing
			If isArray(arr_user) Then
			For uu=0 To Ubound(arr_user,2)%>
			<option value="<%=arr_user(0,uu)%>" <%If trim(User_ID) = trim(arr_user(0,uu)) Then%> selected <%End if%>><%=arr_user(1,uu)%></option>			
		<%Next
		End If%>
		</select>
		</TD>
		<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<span id=word6 name=word6><!--עובד--><%=arrTitles(6)%></span>&nbsp;</td>
	</TR>
	
		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>"><input type=image src='../../images/calend.gif'   border=0  onclick='return popupcal(dateStart);' ID="Image1" NAME="Image1">&nbsp;<input dir="ltr" type="text" class="passw" size=12 id="dateStart" name="dateStart" value="<%=start_date%>" maxlength=10 readonly>&nbsp;</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id=word7 name=word7><!--- תאריך מ--><%=arrTitles(7)%></span>&nbsp;</TD>
		</TR>
		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>"><input type=image src='../../images/calend.gif'   border=0  onclick='return popupcal(dateEnd);' ID="Image2" NAME="Image2">&nbsp;<input dir="ltr" type="text" class="passw" size=12 id="dateEnd" name="dateEnd" value="<%=end_date%>" maxlength=10 readonly>&nbsp;</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id="word8" name=word8><!--- תאריך עד--><%=arrTitles(8)%></span>&nbsp;</TD>
		</TR>	
		<TR>						
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_openedText('reports.asp');" style="width:250">דוח שליחת הודעות</a></TD>						
		</TR>					
	
		<TR>						
			<TD align=center bgcolor="#DBDBDB" colspan=2 dir=rtl><A class="but_menu" href="#" onclick="report_openedText('reports.asp');" style="width:250">דוח שליחת הודעות ב-Excel</a></TD>						
		</TR>		
			<TR>						
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_openedText('reports.asp');" style="width:250">דוח הודעות שהתקבלו</a></TD>						
		</TR>					
			<TR>						
			<TD align=center bgcolor="#DBDBDB" colspan=2 dir=rtl><A class="but_menu" href="#" onclick="report_openedText('reports.asp');" style="width:250">דוח הודעות שהתקבלו ב-Excel</a></TD>						
		</TR>		

			
	</TABLE>			
</FORM>
  <!-- End main content -->
</td></tr>
		<TR>						
			<TD align=center  dir=rtl><A class="but_menu_red" href="#" onclick="report_openfr('reportNewMess.asp');" style="width:250">דוח נציגים שלא קראו 5 הודעות ומעלה</a></TD>						
		</TR>		
	
</table>	 
</body>
<%set con=nothing%>
</html>

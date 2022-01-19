<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
	companyID = trim(Request("companyID"))
	userId = trim(Request("userId"))
	start_date = "1/" & Month(Date()) & "/" & Year(date())
	end_date = DateAdd("d",30,start_date)
	if lang_id = "1" then
  arr_Status = Array("","חדש","נקרא","תגובה")	
 	else
   arr_Status = Array("","new","read","replay")	
	end if
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 69 Order By word_id"				
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
<script language=javascript>
<!--
	
	function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
   }

	
	function report_open(fname)
	{	
		formsearch.action=fname;
		formsearch.target = '_blank';
		formsearch.submit();	
    }
	function report_openSelf(fname)
	{
	    formsearch.action=fname;
		formsearch.target = '_self';
		formsearch.submit();	
	}
//-->
</script>

</head>
<body topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 bgcolor="#E5E5E5">
<!--#include file="../../logo_top.asp"-->
<div align="center"><center>
<table border="0" width="100%" cellspacing="0" cellpadding="0"  dir="<%=dir_var%>">
<%numOftab = 49%>
<%numOfLink = 3%>
<%topLevel2 = 56 'current bar ID in top submenu - added 03/10/2019%>
  <tr><td width=100% align="<%=align_var%>"><!--#include file="../../top_in.asp"--></td></tr>
</table>
</center></div>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<tr><td class="page_title" dir="<%=dir_obj_var%>">&nbsp;</td></tr>
<tr><td height=20 nowrap></td></tr>
<tr>
   <td width="100%" align=center valign=top colspan=2>      
	<table border="0" cellspacing="1" width=450 align=center bgcolor="#646E77" cellpadding="4">
	<FORM action="" method=POST id=formsearch name=formsearch>
	<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">		
		<select name="T" id="T" class="norm" dir="<%=dir_obj_var%>" style="width:310">		
		<option value="OUT" id=word1 name=word1>הודעות שנשלחו <%'=arrTitles(1)%></option>		
		<option value="IN" id=word2 name=word2>הודעות שהתקבלו <%'=arrTitles(2)%></option>		
		</select>
		</TD>
		<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<span id="word3" name="word3"><!--סוג--><%=arrTitles(3)%></span>&nbsp;</td>
	</TR>
	<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">		
		<select name="user_id" id="user_id" class="norm" dir="<%=dir_obj_var%>" style="width:310">		
		<option value="" id=word4 name=word4><%=arrTitles(4)%></option>		
		<%  sqlstr = "Select USER_ID, FIRSTNAME + ' ' + LASTNAME FROM USERS WHERE active=1 and ORGANIZATION_ID = " & OrgID			
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
		<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<span id="word5" name=word5><!--עובד--><%=arrTitles(5)%></span>&nbsp;</td>
	</TR>

	<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">
		<select name="Message_status" id="Message_status" class="norm" dir="<%=dir_obj_var%>" style="width:310">
		<%If trim(lang_id) = "1" Then%>
		<option value=""><%=String(25,"-")%> כל הסטטוסים <%=String(25,"-")%></option>		
		<%Else%>
		<option value=""><%=String(25,"-")%> All statuses <%=String(25,"-")%></option>		
		<%End If%>
		<%For i=1 To uBound(arr_Status)	%>
		<option value="<%=i%>"><%=arr_Status(i)%></option>	
		<%Next%> 
		</select></td>
		<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<span id="word6" name=word6><!--סטטוס--><%=arrTitles(6)%></span>&nbsp;</td>
	</TR>   
	<TR>						
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>"><input type=image src='../../images/calend.gif'   border=0  onclick='return popupcal(date_start);'>&nbsp;<input dir="ltr" type="text" class="passw" size=12 id="date_start" name="date_start" value="<%=start_date%>" maxlength=10 readonly>&nbsp;</TD>
		<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<span id=word7 name=word7>מתאריך</span>&nbsp;</td>
	</TR>
	<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>"><input type=image src='../../images/calend.gif'   border=0  onclick='return popupcal(date_end);'>&nbsp;<input dir="ltr" type="text" class="passw" size=12 id="date_end" name="date_end" value="<%=end_date%>" maxlength=10 readonly>&nbsp;</TD>
			<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<span id="word8" name=word8>עד תאריך</span>&nbsp;</td>
	</TR>	
		<tr id=but_tr>
			<td bgcolor="#dbdbdb" align="center" colspan="2">
				<TABLE WIDTH=100% align=center BORDER=0 CELLSPACING=5 CELLPADDING=0>
				 	<TR>
						<td width=20%>&nbsp;</td>
						<TD align=center width=60% dir=rtl><A class="but_menu" href="#" onclick="report_open('report_MessagesEx.asp');"><span id=word9 name=word9>הצג דוח ב- Excel<%'=arrTitles(9)%></span></a></TD>
						<td width=20%>&nbsp;</td>
					</TR>	
					 	<TR>
						<td width=20%>&nbsp;</td>
						<TD align=center width=60% dir=rtl><A class="but_menu" href="#" onclick="report_openSelf('report_Messages.asp');"><span id="Span2" name=word9>הצג דוח <%'=arrTitles(9)%></span></a></TD>
						<td width=20%>&nbsp;</td>
					</TR>																									
					</table></td></tr>
				</form>		
  <!-- End main content -->
</td></tr>

</table>
</td></tr>
<TR><td height=15></td></TR>
		<TR>						
			<TD align=center  dir=rtl><A class="but_menu_red" href="#" onclick="report_open('reportNewMess.asp');" style="width:250">דוח נציגים שלא קראו 5 הודעות ומעלה</a></TD>						
		</TR>		

</table>
</center></div>
</body>
<%set con=nothing%>
</html>

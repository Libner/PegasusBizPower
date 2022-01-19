<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
<!--
function CheckField()
{
  var EMAILStr =  document.dataForm.peopleEMAIL.value;  
  if (window.document.dataForm.peopleEMAIL.value=='')
  {
	<%
	  If trim(lang_id) = "1" Then
		  str_alert = "!חובה להכניס כתובת אימייל"
	  Else
		  str_alert = "Please insert the email address !!"
	  End If   
	%>
	  window.alert("<%=str_alert%>");  	
	  document.dataForm.peopleEMAIL.focus();
	  return false;
  }
  else if (checkEmail(EMAILStr) == false)
  {
  	<%
	  If trim(lang_id) = "1" Then
		  str_alert = "!כתובת האימייל אינה חוקית"
	  Else
		  str_alert = "The email address is not valid!!"
	  End If   
	%>
	  window.alert("<%=str_alert%>");  	
	  document.dataForm.peopleEMAIL.focus();
	  return false;		
  }  
  return true;
}
function checkEmail(addr)
{
	addr = addr.replace(/^\s*|\s*$/g,"");
	if (addr == '') {
	return false;
	}
	var invalidChars = '\/\'\\ ";:?!()[]\{\}^|';
	for (i=0; i<invalidChars.length; i++) {
	if (addr.indexOf(invalidChars.charAt(i),0) > -1) {
		return false;
	}
	}
	for (i=0; i<addr.length; i++) {
	if (addr.charCodeAt(i)>127) {     
		return false;
	}
	}

	var atPos = addr.indexOf('@',0);
	if (atPos == -1) {
	return false;
	}
	if (atPos == 0) {
	return false;
	}
	if (addr.indexOf('@', atPos + 1) > - 1) {
	return false;
	}
	if (addr.indexOf('.', atPos) == -1) {
	return false;
	}
	if (addr.indexOf('@.',0) != -1) {
	return false;
	}
	if (addr.indexOf('.@',0) != -1){
	return false;
	}
	if (addr.indexOf('..',0) != -1) {
	return false;
	}
	var suffix = addr.substring(addr.lastIndexOf('.')+1);
	if (suffix.length != 2 && suffix != 'com' && suffix != 'net' && suffix != 'org' && suffix != 'edu' && suffix != 'int' && suffix != 'mil' && suffix != 'gov' & suffix != 'arpa' && suffix != 'biz' && suffix != 'aero' && suffix != 'name' && suffix != 'coop' && suffix != 'info' && suffix != 'pro' && suffix != 'museum') {
	return false;
	}
return true;
}
//-->
</script>
</head>
<body style="margin:0px;background-color:#E5E5E5">
<%	
	
	PEOPLE_ID=trim(Request("PEOPLE_ID"))	
	groupId=trim(Request("groupId"))
	if Request.Form("editFlag")<>nil then
		cl_mail = sFix(request.form("peopleEMAIL"))
		if PEOPLE_ID <> "" then
		sqlstr="Select PEOPLE_ID From PEOPLES Where PEOPLE_EMAIL='"& cl_mail &"' and PEOPLE_ID<>" & PEOPLE_ID & " And GROUPID = "&groupId&" And ORGANIZATION_ID = " & OrgID 		
		else
		sqlstr="Select PEOPLE_ID From PEOPLES Where PEOPLE_EMAIL='"& cl_mail &"'" & " AND GROUPID = "&groupId&" And ORGANIZATION_ID = " & OrgID 			
		end if
		'Response.Write sqlstr
		'Response.End
		set checks = con.GetRecordSet(sqlstr)	
			if not checks.eof then %>
				<SCRIPT LANGUAGE=javascript>
				<!--
					alert('! המשתמש כבר קיים במערכת');
					window.history.back();
				//-->
				</SCRIPT>
	<%			ischecked = false
			else
				ischecked = true
			end if
			set checks = nothing
	if 	ischecked = true then
		if PEOPLE_ID<>"" then			
			sqlstr=	"UPDATE PEOPLES SET PEOPLE_NAME='"&sFix(trim(request.form("peopleName")))&_			
			"',PEOPLE_COMPANY='" & sFix(trim(request.form("peopleCOMPANY")))&_
			"',PEOPLE_OFFICE='" & sFix(trim(request.form("peopleOFFICE")))&_			
			"',PEOPLE_PHONE='" & sFix(trim(request.form("peoplePHONE")))&_			
			"',PEOPLE_FAX='" & sFix(trim(request.form("peopleFAX")))&_
			"',PEOPLE_CELL='" & sFix(trim(request.form("peopleCELL")))&_
			"',PEOPLE_EMAIL='" & sFix(trim(request.form("peopleEMAIL")))&_
			"' where PEOPLE_ID=" & PEOPLE_ID
			'Response.Write("upd" & sqlstr)
			con.ExecuteQuery (sqlstr)
		elseif trim(Request("groupId")) <> "" then						
			sqlstring="SET NOCOUNT ON; INSERT INTO PEOPLES (User_ID,ORGANIZATION_ID,PEOPLE_NAME,PEOPLE_PHONE,"&_
			"PEOPLE_FAX,PEOPLE_CELL,PEOPLE_EMAIL,PEOPLE_OFFICE,PEOPLE_COMPANY,groupId) values ("&_
			UserID & "," & OrgID & ",'" & sFix(trim(request.form("peopleName"))) &"','" &_
			sFix(trim(request.form("peoplePHONE")))&"','"&sFix(trim(request.form("peopleFAX")))&_
            "','"&sFix(trim(request.form("peopleCELL")))&"','"&sFix(trim(request.form("peopleEMAIL")))&"','"&_
			sFix(trim(request.form("peopleOFFICE")))&"','"&_
			sFix(trim(request.form("peopleCOMPANY")))&"',"& Request("groupId") &"); SELECT @@IDENTITY AS NewID"
			
			set rs_tmp = con.getRecordSet(sqlstring)
				PEOPLE_ID = rs_tmp.Fields("NewID").value
			set rs_tmp = Nothing	  		
			
		end if 
	   end if 'ischecked = true 
	%>
	<SCRIPT LANGUAGE=javascript>
	<!--	
		opener.focus();
		opener.window.location.reload();
		self.close();
	//-->
	</SCRIPT>
  <%		
	end if 'if Request.Form("editFlag")<>nil
	
	if PEOPLE_ID<>"" then
		set people=con.GetRecordSet("select PEOPLE_NAME, PEOPLE_PHONE, PEOPLE_FAX, PEOPLE_CELL, PEOPLE_EMAIL, PEOPLE_COMPANY, PEOPLE_OFFICE  from PEOPLES where PEOPLE_ID=" & PEOPLE_ID)
		if not people.eof then
		peopleName=trim(people("PEOPLE_NAME"))				
		peoplePHONE=trim(people("PEOPLE_PHONE"))
		peopleFAX=trim(people("PEOPLE_FAX"))
		peopleCELL = trim(people("PEOPLE_CELL"))
		peopleEMAIL=trim(people("PEOPLE_EMAIL"))
		peopleCOMPANY=trim(people("PEOPLE_COMPANY"))
		peopleOFFICE=vFix(trim(people("PEOPLE_OFFICE")))
		end if		
		people.close
		set people=nothing
	end if
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 48 Order By word_id"				
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
	  
	  sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	  set rsbuttons = con.getRecordSet(sqlstr)
	  If not rsbuttons.eof Then
		arrButtons = ";" & rsbuttons.getString(,,",",";")		
		arrButtons = Split(arrButtons, ";")
	  End If
	  set rsbuttons=nothing	 	  
%>
<table border="0" dir="<%=dir_var%>" width="100%" cellspacing="0" cellpadding="0" align="<%=align_var%>" valign="top">
<tr>
<td width="100%" class="page_title" dir=rtl>&nbsp;<%if PEOPLE_ID<>"" then%><span id=word1 name=word1><!--עדכון--><%=arrTitles(1)%></span><%Else%><span id="word2" name=word2><!--הוספת--><%=arrTitles(2)%></span><%End If%>&nbsp;<span id=word3 name=word3><!--נמען--><%=arrTitles(3)%></span></td></tr>         
<tr><td height=15 nowrap></td></tr>
<tr><td width="100%">
<table align=center border="0" cellpadding="2" cellspacing="1" width="70%">
<FORM name="dataForm" ACTION="addpeople.asp?groupId=<%=Request("groupId")%>&page=<%=Request("page")%>" METHOD="post" onSubmit="return CheckField();">
	<tr>
		<td align="<%=align_var%>"><input dir="ltr" type="text" class="passw" size=40 name="peopleEMAIL" value="<%=vFix(peopleEMAIL)%>" maxlength=100 ID="peopleEMAIL"></td>
		<td align="<%=align_var%>"><font color=red>&nbsp;E-mail&nbsp;&nbsp;*</font>&nbsp;</td>
	</tr>
	<tr>
		<td align="<%=align_var%>"><input dir="<%=dir_obj_var%>" type="text" class="passw" size=40 maxlength=100 name="peopleName" value="<%=vFix(peopleName)%>"></td>
		<td align="<%=align_var%>" style="padding-right:5px">&nbsp;<span id="word4" name=word4><!--שם מלא--><%=arrTitles(4)%></span>&nbsp;</td>
	</tr>	
	<tr>
		<td align="<%=align_var%>"><input dir="<%=dir_obj_var%>" type="text" class="passw" size=40 maxlength=100 name="peopleCOMPANY" value="<%=vFix(peopleCOMPANY)%>"></td>
		<td align="<%=align_var%>" style="padding-right:5px">&nbsp;<span id="word5" name=word5><!--חברה--><%=arrTitles(5)%></span>&nbsp;</td>
	</tr>
	<tr>
		<td align="<%=align_var%>"><input dir="<%=dir_obj_var%>" type="text" class="passw" size=20 name="peopleOFFICE" value="<%=vFix(peopleOFFICE)%>"></td>
		<td align="<%=align_var%>" style="padding-right:5px">&nbsp;<span id="word6" name=word6><!--תפקיד--><%=arrTitles(6)%></span>&nbsp;</td>
	</tr>
	<tr>
		<td align="<%=align_var%>"><input dir="<%=dir_obj_var%>" type="text" class="passw" size=20 name="peopleCell" value="<%=vFix(peopleCELL)%>"></td>
		<td align="<%=align_var%>" style="padding-right:5px">&nbsp;<span id="word7" name=word7><!--טלפון נייד--><%=arrTitles(7)%></span>&nbsp;</td>
	</tr>	
	<tr>
		<td align="<%=align_var%>"><input dir="<%=dir_obj_var%>" type="text" class="passw" size=20 name="peoplePHONE" value="<%=vFix(peoplePHONE)%>"></td>
		<td align="<%=align_var%>" style="padding-right:5px">&nbsp;<span id="word8" name=word8><!--טלפון--><%=arrTitles(8)%></span>&nbsp;</td>
	</tr>	
	<tr>
		<td align="<%=align_var%>"><input dir="<%=dir_obj_var%>" type="text" class="passw" size=20 name="peopleFAX" value="<%=vfix(peopleFAX)%>"></td>
		<td align="<%=align_var%>" style="padding-right:5px">&nbsp;<span id="word9" name=word9><!--פקס--><%=arrTitles(9)%></span>&nbsp;</td>
	</tr>		
	<tr><td colspan=2 height="25" nowrap></td></tr>
	<tr><td colspan=2>
	<table cellpadding=0 cellspacing=0 width=100%>
	<tr>
	<td width=45% nowrap align="<%=align_var%>"><INPUT class="but_menu" type="button" style="width:90" value="<%=arrButtons(2)%>" id=button2 name=button2 onclick="window.close()"></td>
	<td width=5% nowrap></td>
	<td width=45% nowrap align="left">
			<input class="but_menu" style="width:90" type="submit" value="<%=arrButtons(1)%>" id=Button1 name=Button1>
			<input type="hidden" name="PEOPLE_ID" value="<%=PEOPLE_ID%>">
			<input type="hidden" name="editFlag" value="yes">
		</td>
	</tr>
	</table></td></tr>	
</FORM>
</table></td></tr></table>
<%set con=nothing%>
</body>
</html>

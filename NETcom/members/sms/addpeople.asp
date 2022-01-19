<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkworker.asp"-->
<%If isNumeric(trim(Request("groupId"))) And trim(Request("groupId")) <> "" Then   
		groupId = cInt(trim(Request("groupId")))
	Else
		groupId = 0
	End If
	
	If isNumeric(trim(Request("PeopleID"))) And trim(Request("PeopleID")) <> "" Then   
		PeopleID = clng(trim(Request("PeopleID")))
	Else
		PeopleID = 0
	End If	
	
	if Request.Form("editFlag")<>nil then
		new_cell = sFix(request.form("peopleCell"))
		if PeopleID > 0 then
		sqlstr="SELECT PEOPLE_ID From sms_peoples WHERE (PEOPLE_CELL='"& new_cell &"') AND (PEOPLE_ID<>" & PeopleID & ") And GROUP_ID = "&groupId&" And ORGANIZATION_ID = " & OrgID 		
		else
		sqlstr="SELECT PEOPLE_ID From sms_peoples WHERE (PEOPLE_CELL='"& new_cell &"') AND (GROUP_ID = "&groupId&") And ORGANIZATION_ID = " & OrgID 			
		end if
		'Response.Write sqlstr
		'Response.End
		set checks = con.GetRecordSet(sqlstr)	
			if not checks.eof then %>
				<script language=javascript>
				<!--
					window.alert('! טלפון נייד כבר קיים בקבוצה');
					window.history.back();
				//-->
				</script>
	<%			ischecked = false
			else
				ischecked = true
			end if
			set checks = nothing
	if 	ischecked = true then
		if PeopleID > 0 then			
			sqlstr=	"UPDATE sms_peoples SET PEOPLE_NAME='"&sFix(trim(request.form("peopleName")))&_			
			"',PEOPLE_COMPANY='" & sFix(trim(request.form("peopleCOMPANY")))&_
			"',PEOPLE_OFFICE='" & sFix(trim(request.form("peopleOFFICE")))&_			
			"',PEOPLE_PHONE='" & sFix(trim(request.form("peoplePHONE")))&_			
			"',PEOPLE_FAX='" & sFix(trim(request.form("peopleFAX")))&_
			"',PEOPLE_CELL='" & sFix(trim(request.form("peopleCELL")))&_
			"' WHERE PEOPLE_ID=" & PeopleID
			'Response.Write("upd" & sqlstr)
			con.ExecuteQuery (sqlstr)
		elseif IsNumeric(groupId) then						
			sqlstring="SET NOCOUNT ON; INSERT INTO sms_peoples (User_ID,ORGANIZATION_ID,PEOPLE_NAME,PEOPLE_PHONE,"&_
			"PEOPLE_FAX,PEOPLE_CELL,PEOPLE_OFFICE,PEOPLE_COMPANY,GROUP_ID) values ("&_
			UserID & "," & OrgID & ",'" & sFix(trim(request.form("peopleName"))) &"','" &_
			sFix(trim(request.form("peoplePHONE"))) & "','" & sFix(trim(request.form("peopleFAX"))) & "','" & _
			sFix(trim(request.form("peopleCELL"))) & "','" & sFix(trim(request.form("peopleOFFICE"))) & "','" &_
			sFix(trim(request.form("peopleCOMPANY"))) & "'," & groupId & "); SELECT @@IDENTITY AS NewID"			
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
	
	If PeopleID > 0 Then
		Set people=con.GetRecordSet("SELECT PEOPLE_NAME, PEOPLE_PHONE, PEOPLE_FAX, PEOPLE_CELL, PEOPLE_COMPANY, PEOPLE_OFFICE  FROM sms_peoples WHERE PEOPLE_ID=" & PeopleID)
		If not people.eof Then
			peopleName = trim(people("PEOPLE_NAME"))				
			peoplePHONE = trim(people("PEOPLE_PHONE"))
			peopleFAX = trim(people("PEOPLE_FAX"))
			peopleCELL = trim(people("PEOPLE_CELL"))
			peopleCOMPANY = trim(people("PEOPLE_COMPANY"))
			peopleOFFICE = vFix(trim(people("PEOPLE_OFFICE")))
		End If		
		people.close
		set people=nothing
	end if
	
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
	set rsbuttons=nothing	 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css" />
<script LANGUAGE="JavaScript">
<!--
function CheckField()
{
  var CellStr = new String(document.dataForm.peopleCell.value);
  if (CellStr.length < 10)
  {
	<%
	  If trim(lang_id) = "1" Then
		  str_alert = "!חובה להכניס טלפון נייד, 10 ספרות ללא רווחים"
	  Else
		  str_alert = "Please insert the cell phone !!"
	  End If   
	%>
	  window.alert("<%=str_alert%>");  	
	  document.dataForm.peopleCell.focus();
	  return false;
  }
  
  return true;
 }
 
 function GetNumbers(){
	var ch=event.keyCode;
	event.returnValue =(ch >= 48 && ch <= 57);
}
//-->
</script>
</head>
<body style="margin:0px;background-color:#E5E5E5">
<table border="0" dir="<%=dir_var%>" width="100%" cellspacing="0" cellpadding="0" align="<%=align_var%>" valign="top">
<tr>
<td width="100%" class="page_title" dir=rtl>&nbsp;<%if PeopleID > 0 then%><!--עדכון--><%=arrTitles(1)%><%Else%><!--הוספת--><%=arrTitles(2)%><%End If%>&nbsp<!--נמען--><%=arrTitles(3)%></td></tr>         
<tr><td height=15 nowrap></td></tr>
<tr><td width="100%">
<FORM name="dataForm" ACTION="addpeople.asp?groupId=<%=Request("groupId")%>" METHOD="post" onSubmit="return CheckField();" ID="dataForm">
<table align=center border="0" cellpadding="2" cellspacing="1" width="80%">
	<tr>
		<td align="<%=align_var%>" dir="<%=dir_obj_var%>"><input dir="ltr" type="text" class="passw" size="20" maxlength="10" name="peopleCell" value="<%=vFix(peopleCELL)%>" ID="peopleCell" onkeypress="GetNumbers();">
		&nbsp;(10 ספרות ללא רווחים)</td>
		<td align="<%=align_var%>" style="padding-right:5px"><font color=red>&nbsp;<!--טלפון נייד--><%=arrTitles(7)%>&nbsp;*</font></td>
	</tr>	
	<tr>
		<td align="<%=align_var%>"><input dir="<%=dir_obj_var%>" type="text" class="passw" size=40 maxlength=100 name="peopleName" value="<%=vFix(peopleName)%>"></td>
		<td align="<%=align_var%>" style="padding-right:5px">&nbsp;<!--שם מלא--><%=arrTitles(4)%>&nbsp;</td>
	</tr>	
	<tr>
		<td align="<%=align_var%>"><input dir="<%=dir_obj_var%>" type="text" class="passw" size=40 maxlength=100 name="peopleCOMPANY" value="<%=vFix(peopleCOMPANY)%>"></td>
		<td align="<%=align_var%>" style="padding-right:5px">&nbsp;<!--חברה--><%=arrTitles(5)%>&nbsp;</td>
	</tr>
	<tr>
		<td align="<%=align_var%>"><input dir="<%=dir_obj_var%>" type="text" class="passw" size=20 name="peopleOFFICE" value="<%=vFix(peopleOFFICE)%>"></td>
		<td align="<%=align_var%>" style="padding-right:5px">&nbsp;<!--תפקיד--><%=arrTitles(6)%>&nbsp;</td>
	</tr>
	<tr>
		<td align="<%=align_var%>"><input dir="<%=dir_obj_var%>" type="text" class="passw" size=20 name="peoplePHONE" value="<%=vFix(peoplePHONE)%>"></td>
		<td align="<%=align_var%>" style="padding-right:5px">&nbsp;<!--טלפון--><%=arrTitles(8)%>&nbsp;</td>
	</tr>	
	<tr>
		<td align="<%=align_var%>"><input dir="<%=dir_obj_var%>" type="text" class="passw" size=20 name="peopleFAX" value="<%=vfix(peopleFAX)%>"></td>
		<td align="<%=align_var%>" style="padding-right:5px">&nbsp;<!--פקס--><%=arrTitles(9)%>&nbsp;</td>
	</tr>		
	<tr><td colspan="2" height="25" nowrap></td></tr>
	<tr><td colspan="2">
	<table cellpadding="0" cellspacing="0" width="100%">
	<tr>
	<td width=45% nowrap align="<%=align_var%>"><INPUT class="but_menu" type="button" style="width:90" value="<%=arrButtons(2)%>" id=button2 name=button2 onclick="window.close()"></td>
	<td width=5% nowrap></td>
	<td width=45% nowrap align="left">
			<input class="but_menu" style="width:90" type="submit" value="<%=arrButtons(1)%>" id=Button1 name=Button1>
			<input type="hidden" name="PeopleID" value="<%=PeopleID%>">
			<input type="hidden" name="editFlag" value="yes">
		</td>
	</tr>
	</table>
 </FORM>
</td></tr></table>
</td></tr></table>
</body>
</html>
<%set con=nothing%>
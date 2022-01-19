<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
    If Request.QueryString("add") <> nil Then
		If trim(Request.Form("Series_Id")) = "" Then ' add type
			Series_Id = trim(Request.Form("Series_Id"))
			sqlstr = "Insert into Series (Series_Name,user_id) values ('" & sFix(Request.Form("Series_Name")) & "',"& sFix(Request.Form("user_id")) &")"
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.href = window.opener.document.location.href;
				window.close();
			//-->
			</SCRIPT>
		
	<%	Else ' update type
			Series_Id = trim(Request.Form("Series_Id"))
			sqlstr="Update Series set Series_Name = '" & sFix(Request.Form("Series_Name")) & "',User_id=" &  sFix(Request.Form("user_id"))  &" Where Series_Id = " & Series_Id
			con.executeQuery(sqlstr) %>
			
	<%	End If
	
		sqlstr = "DELETE FROM UserToSeries WHERE Series_Id = " & Request.Form("Series_Id")
		'Response.Write sqlstr
		'Response.End
		con.executeQuery(sqlstr)
	
		sqlstr="SELECT User_id FROM Users where ImportanceId>0"
		SET rs_U = con.getRecordSet(sqlstr)
		WHILE NOT rs_U.eof
			UID = trim(rs_U(0))
			UVisible = Request.Form("U_"&UID)
			response.Write UID &":"& UVisible &"<BR>"
			If trim(UVisible) = "on" Then
				sqlstr = "INSERT INTO UserToSeries  (User_Id,Series_Id) VALUES (" & UID & "," & Series_Id & ")"
		'	response.Write sqlstr
		'	response.end
			con.executeQuery(sqlstr)
				'UVisible = 1
			'
			End If	
		
		
		rs_U.moveNext
		WEND
		SET rs_U = Nothing  %>
	<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.href = window.opener.document.location.href;
				window.close();
			//-->
			</SCRIPT>	
	<%End If
	
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 33 Order By word_id"				
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
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
		function checkForm()
		{
			if(window.document.form1.Series_Name.value == "")
			{
			   <%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס סדרה "
				Else
					str_alert = "Please insert the position name !!"
				End If   
				%>
				window.alert("<%=str_alert%>");				
				window.document.form1.Series_Name.focus();
				return false;
			}
			return true;			
		}
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<%
	If Request.QueryString("Series_Id") <> nil Then
		Series_Id = trim(Request.QueryString("Series_Id"))
		If Len(Series_Id) > 0 Then
			sqlstr="Select Series_Name,User_Id From Series Where Series_Id = " & Series_Id
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				Series_Name = trim(rssub("Series_Name"))				
				user_id= trim(rssub("user_id"))				
			End If
			set rssub=Nothing
		End If
	End If	
%>
<table dir="<%=dir_var%>" border="0" width="480" cellspacing="0" cellpadding="0" align="center">
<tr >
   <td  align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="1" cellspacing="0">
	   <tr>		 
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(Series_Id) > 0 Then%><span id=word1 name=word1><!--עדכון--><%=arrTitles(1)%></span><%Else%><span id="word2" name=word2><!--הוספת--><%=arrTitles(2)%></span><%End If%>&nbsp;<span id=word4 name=word4>סדרה</span>&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td bgcolor="#e6e6e6" background="../../images/Line_3D.gif" height=2 width="100%"></td></tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%>
<table width="480" cellspacing="1" cellpadding="2" align=center border="0">
<form name=form1 id=form1 action="addSeries.asp?add=1" target="_self" method="post">
<input type=hidden name=Series_Id id=Series_Id value="<%=Series_Id%>">
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts" name="Series_Name" id="Series_Name" value="<%=vFix(Series_Name)%>" dir="<%=dir_obj_var%>" size=100 maxLength=100>	
	</td>
	<td width="150" nowrap align="center">&nbsp;<b><span id=word3 name=word3><!--שם סדרה-->שם סדרה</span></b>&nbsp;</td>	
</tr>
<TR>
	<td align="<%=align_var%>" width=330 nowrap>
	<select dir="<%=dir_obj_var%>" name="user_id" class="norm" style="width:200px" ID="user_id">
	<option value="0" >בחר עובד</option>
<%	sqlstr = "SELECT User_id,FIRSTNAME,LASTNAME FROM USERS WHERE (active =1) ORDER BY LASTNAME,FIRSTNAME"
		set rs_users = con.getRecordSet(sqlstr)
		while not rs_users.eof	%>
	<option value="<%=trim(rs_users(0))%>" <%If trim(user_id) = trim(rs_users(0)) Then%> selected <%End If%>><%=rs_users(1)%>&nbsp;<%=rs_users(2)%></option>
<%	rs_users.moveNext
		wend
		set rs_users = nothing %>
	</select>
<td width="150" nowrap align="center">&nbsp;<b><span id="Span1" name=word3>מנהל יעד</span></b>&nbsp;</td>	
</tr>
<TR>
	<td colspan=2 nowrap align="center">&nbsp;<b><span id="Span2" name=word3>
אנשים ממושבים</span></b>&nbsp;</td>	
</tr>
<TR><td colspan=2 width=100%>
<table border=0 cellpadding=2 cellspacing=2>

<%if Series_Id>0 then
	'sqlstr = "select distinct Users.USER_ID,FIRSTNAME,LASTNAME, IsSeriesSelected=CASE WHEN UserToSeries.Series_Id IS NOT NULL THEN 1 ELSE 0 END  from " & _
	'" Users left join UserToSeries on UserToSeries.User_id=Users.USER_ID where ImportanceId>0" & " and (UserToSeries.Series_Id="&  Series_Id &"   or UserToSeries.Series_Id is null)"
 sqlstr = "select Users.USER_ID,FIRSTNAME,LASTNAME,IsSeriesSelected=CASE WHEN dbo.[IsSelectedUserToSerias]("& Series_Id & ",USER_ID)IS NOT NULL THEN 1 ELSE 0 END  from Users where ImportanceId>0"

else
	sqlstr = "select distinct Users.USER_ID,FIRSTNAME,LASTNAME, IsSeriesSelected=0   from Users where ImportanceId>0 "
end if
'response.Write sqlstr
'response.end
		set rs_users = con.getRecordSet(sqlstr)
	'response.Write Series_Id
		while not rs_users.eof	%>

<tr>
	<td class="form_title" align="right" width=150 ><input type="checkbox" dir="ltr" Id="U_<%=rs_users("User_Id")%>"  name="U_<%=rs_users("User_Id")%>" <%if rs_users("IsSeriesSelected")=1 then%> checked <%end if%></td>
	<td align="<%=align_var%>" class="form_title" nowrap width="200" dir="<%=dir_obj_var%>">&nbsp;<%=rs_users("FIRSTNAME")%>&nbsp;<%=rs_users("LASTNAME")%>&nbsp;</td>
</tr>
<%	rs_users.moveNext
		wend
		set rs_users = nothing %>


</table>

</td></TR>

<tr><td height=35 colspan="2" nowrap></td></tr>
<tr><td align=center colspan="2">
<input type=button value="<%=arrButtons(2)%>" class="but_menu" style="width:90" onclick="window.close();" id=Button2 name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=submit value="<%=arrButtons(1)%>" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1></td></tr>
</form>
</table>
</td></tr></table>
</BODY>
</HTML>

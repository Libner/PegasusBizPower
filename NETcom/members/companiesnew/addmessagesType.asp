<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
    If Request.QueryString("add") <> nil Then
	response.Write (Request.QueryString("add"))
		If trim(Request.Form("messages_type_id")) = "" Then ' add type
			sqlstr = "Insert into messages_types (Organization_ID,messages_type_name,messages_type_Content) values (" &_
			trim(Request.Cookies("bizpegasus")("OrgID")) & ",'" & sFix(Request.Form("messages_type_name")) & "','"& sFix(request.Form("messages_type_Content")) &"')"
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.href = window.opener.document.location.href;
				window.close();
			//-->
			</SCRIPT>
		
	<%	Else ' update type
			messages_type_id = trim(Request.Form("messages_type_id"))
			sqlstr="Update messages_types set messages_type_name = '" & sFix(Request.Form("messages_type_name")) &"',messages_type_Content = '" & sFix(Request.Form("messages_type_Content")) & "' Where messages_type_id = " & messages_type_id
			'response.Write sqlstr
			'response.end
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.href = window.opener.document.location.href;
				window.close();
			//-->
			</SCRIPT>	
	<%	End If
	End If
	
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 35 Order By word_id"				
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
<title>BizPower</title>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
		function checkForm()
		{
			if(window.document.form1.messages_type_name.value == "")
			{
			   <%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס שם תבנית "
				Else
					str_alert = "Please insert the template name !!"
				End If   
				%>
				window.alert("<%=str_alert%>");				
				window.document.form1.messages_type_name.focus();
				return false;
			}
			return true;			
		}
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<%
	If Request.QueryString("messages_type_id") <> nil Then
		messages_type_id = trim(Request.QueryString("messages_type_id"))
		If Len(messages_type_id) > 0 Then
			sqlstr="Select messages_type_name,messages_type_Content From messages_types Where messages_type_id = " & messages_type_id
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				messages_type_name = trim(rssub("messages_type_name"))
				messages_type_Content=trim(rssub("messages_type_Content"))
								
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
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(item_id) > 0 Then%><span id=word1 name=word1><!--עדכון--><%=arrTitles(1)%></span><%Else%><span id="word2" name=word2><!--הוספת-->תבנית הודעה<%'=arrTitles(2)%></span><%End If%>&nbsp;<span id=word4 name=word4><!--סוג--><%'=arrTitles(4)%></span></td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%>
<table width="480" cellspacing="1" cellpadding="2" align=center border="0" dir="<%=dir_var%>">
<form name=form1 id=form1 action="addmessagesType.asp?add=1" target="_self" method="post">
<input type=hidden name=messages_type_id id=messages_type_id value="<%=messages_type_id%>">
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" dir="<%=dir_obj_var%>" class="texts" name="messages_type_name" id="messages_type_name" value="<%=vFix(messages_type_name)%>" dir=rtl size=100 maxLength=100>	
	</td>
	<td width="150" nowrap align="center">&nbsp;<b><span id=word3 name=word3><!--שם סוג-->שם<%'=arrTitles(3)%></span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<textarea  dir="<%=dir_obj_var%>" class="texts" name="messages_type_Content" id="messages_type_Content"  dir=rtl size=100 rows=5 cols=150><%=vFix(messages_type_Content)%></textarea>	
	</td>
	<td width="150" nowrap align="center">&nbsp;<b><span id="word4" name=word4><!--שם סוג-->תוכן הודעה</span></b>&nbsp;</td>	
</tr>
<tr><td height=5 colspan="2" nowrap></td></tr>
<tr><td align=center colspan="2">
<input type=button value="<%=arrButtons(2)%>" class="but_menu" style="width:90" onclick="window.close();" id=Button2 name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=submit value="<%=arrButtons(1)%>" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1></td></tr>
</form>
</table>
</td></tr></table>
</BODY>
</HTML>

<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->

<%
    If Request.QueryString("add") <> nil Then
		If trim(Request.Form("movement_type_id")) = "" Then ' add type
			sqlstr = "Insert into movements_types (Organization_ID,movement_type_name,type_id) values (" &_
			trim(Request.Cookies("bizpegasus")("OrgID")) & ",'" &_
			sFix(Request.Form("movement_type_name")) & "'," & Request.Form("type_id") & ")"
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				opener.focus();
				opener.window.location.reload();
				self.close();
			//-->
			</SCRIPT>
		
	<%	Else ' update type
			movement_type_id = trim(Request.Form("movement_type_id"))
			sqlstr="Update movements_types set movement_type_name = '" & sFix(Request.Form("movement_type_name")) &_
			"', type_id = " & Request.Form("type_id") & " Where movement_type_id = " & movement_type_id
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				opener.focus();
				opener.window.location.reload();
				self.close();
			//-->
			</SCRIPT>	
	<%	End If
	End If
	
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 43 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(trim(arr_title(0))) = arr_title(1)
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
			if(window.document.form1.movement_type_name.value == "")
			{
			   <%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס סוג תנועה"
				Else
					str_alert = "Please insert the name of movement  !!"
				End If   
				%>			
				window.alert("<%=str_alert%>");
				window.document.form1.movement_type_name.focus();
				return false;
			}
			return true;			
		}
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<%
	If trim(lang_id) = "1" Then
	     type_name = Array("","הוצאה" ,"הכנסה")    	
	Else
		 type_name = Array("","expense" ,"income")
	End If	
	If Request.QueryString("movement_type_id") <> nil Then
		movement_type_id = trim(Request.QueryString("movement_type_id"))
		If Len(movement_type_id) > 0 Then
			sqlstr="Select movement_type_name,type_id From movements_types Where movement_type_id = " & movement_type_id
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				movement_type_name = trim(rssub("movement_type_name"))	
				type_id	= trim(rssub("type_id"))		
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
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(movement_type_id) > 0 Then%><span id=word1 name=word1><!--עדכון--><%=arrTitles(1)%></span><%Else%><span id="word2" name=word2><!--הוספת--><%=arrTitles(2)%></span><%End If%>&nbsp;<span id=word4 name=word4><!--סוג תנועה--><%=arrTitles(4)%></span></td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%>
<table width="480" cellspacing="1" cellpadding="2" align=center border="0" dir="<%=dir_var%>">
<form name=form1 id=form1 action="addMovementType.asp?add=1" target="_self" method="post">
<input type=hidden name=movement_type_id id=movement_type_id value="<%=movement_type_id%>">
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<select class="norm" name="type_id" id="type_id" dir="<%=dir_obj_var%>">
	<%For i=1 To Ubound(type_name)%>
	<option value="<%=i%>" <%If trim(type_id) = trim(i) Then%>selected<%End If%>><%=type_name(i)%></option>
	<%Next%>
	</select>	
	</td>
	<td width="150" nowrap align="center">&nbsp;<b><span id=word3 name=word3><!--סוג תנועה--><%=arrTitles(3)%></span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts" name="movement_type_name" id="movement_type_name" value="<%=vFix(movement_type_name)%>"  dir="<%=dir_obj_var%>" size=100 maxLength=100>	
	</td>
	<td width="150" nowrap align="center">&nbsp;<b><span id="word5" name=word5><!--שם תנועה--><%=arrTitles(5)%></span></b>&nbsp;</td>	
</tr>
<tr><td height=35 colspan="2" nowrap></td></tr>
<tr><td align=center colspan="2">
<input type=button value="<%=arrButtons(2)%>" class="but_menu" style="width:90" onclick="window.close();" id=Button2 name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=submit value="<%=arrButtons(1)%>" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1></td></tr>
</form>
</table>
</td></tr></table>
</BODY>
</HTML>

<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
    If Request.QueryString("add") <> nil Then
		If trim(Request.Form("item_id")) = "" Then ' add type
			sqlstr = "Insert into messangers (Organization_ID,item_Name) values (" &_
			trim(Request.Cookies("bizpegasus")("OrgID")) & ",'" & sFix(Request.Form("item_Name")) & "')"
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.href = window.opener.document.location.href;
				window.close();
			//-->
			</SCRIPT>
		
	<%	Else ' update type
			item_id = trim(Request.Form("item_id"))
			sqlstr="Update messangers set item_Name = '" & sFix(Request.Form("item_Name")) & "' Where item_id = " & item_id
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
			if(window.document.form1.item_Name.value == "")
			{
			   <%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס עיסוק "
				Else
					str_alert = "Please insert the position name !!"
				End If   
				%>
				window.alert("<%=str_alert%>");				
				window.document.form1.item_Name.focus();
				return false;
			}
			return true;			
		}
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<%
	If Request.QueryString("item_id") <> nil Then
		item_id = trim(Request.QueryString("item_id"))
		If Len(item_id) > 0 Then
			sqlstr="Select item_Name From messangers Where item_id = " & item_id
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				item_Name = trim(rssub("item_Name"))				
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
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(item_id) > 0 Then%><span id=word1 name=word1><!--עדכון--><%=arrTitles(1)%></span><%Else%><span id="word2" name=word2><!--הוספת--><%=arrTitles(2)%></span><%End If%>&nbsp;<span id=word4 name=word4><!--עיסוק--><%=arrTitles(4)%></span>&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td bgcolor="#e6e6e6" background="../../images/Line_3D.gif" height=2 width="100%"></td></tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%>
<table width="480" cellspacing="1" cellpadding="2" align=center border="0">
<form name=form1 id=form1 action="addmessanger.asp?add=1" target="_self" method="post">
<input type=hidden name=item_id id=item_id value="<%=item_id%>">
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts" name="item_Name" id="item_Name" value="<%=vFix(item_Name)%>" dir="<%=dir_obj_var%>" size=100 maxLength=100>	
	</td>
	<td width="150" nowrap align="center">&nbsp;<b><span id=word3 name=word3><!--שם עיסוק--><%=arrTitles(3)%></span></b>&nbsp;</td>	
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

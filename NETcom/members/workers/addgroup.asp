<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
    If Request.QueryString("add") <> nil Then
        group_id = trim(Request.Form("group_id"))	
        group_name = sFix(Request.Form("group_name"))       
       
		If trim(Request.Form("group_id")) = "" Then ' add type
			sqlstr = "Insert into Users_Groups (Organization_ID,group_name) values (" &_
			trim(OrgId) & ",'" & trim(group_name) & "')"
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.href = window.opener.document.location.href;
				window.close();
			//-->
			</SCRIPT>
		
	<%	Else ' update
					
			sqlstr="Update Users_Groups set group_name = '" & trim(group_name) &_
			"' Where group_id = " & group_id
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
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 65 Order By word_id"				
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
			if(window.document.form1.group_name.value == "")
			{
				window.alert("!!נא להכניס קבוצה");
				window.document.form1.group_name.focus();
				return false;
			}			
				
			return true;
			   
		}
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<%
	If Request.QueryString("group_id") <> nil Then
		group_id = trim(Request.QueryString("group_id"))
		If Len(group_id) > 0 Then
			sqlstr="Select group_name From Users_Groups Where group_id = " & group_id
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				group_name = trim(rssub("group_name"))					
			End If
			set rssub=Nothing
		End If
	End If	
%>
<table border="0" width="480" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>">
<tr >
   <td  align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="1" cellspacing="0" ID="Table1">
	   <tr>		 
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(group_id) > 0 Then%><span id=word1 name=word1><!--עדכון--><%=arrTitles(1)%></span><%Else%><span id="word2" name=word2><!--הוספת--><%=arrTitles(2)%></span><%End If%>&nbsp;<span id=word4 name=word4><!--קבוצה--><%=arrTitles(4)%></span>&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%>
<table width="460" cellspacing="1" cellpadding="2" align=center border="0">
<form name=form1 id=form1 action="addgroup.asp?add=1" target="_self" method="post">
<input type=hidden name=group_id id=group_id value="<%=group_id%>">
<input type=hidden name=ORGANIZATION_ID id="ORGANIZATION_ID" value="<%=OrgId%>">
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" name="group_name" id="group_name" class="form" value="<%=vFix(group_name)%>" dir="<%=dir_obj_var%>" style="width:300" maxLength=50>	
	</td>
	<td width="150" nowrap align="<%=align_var%>">&nbsp;<b><span id=word3 name=word3><!--שם קבוצה--><%=arrTitles(3)%></span></b>&nbsp;</td>	
</tr>
<tr><td height=35 colspan="2" nowrap></td></tr>
<tr><td align=center colspan="2">
<input type=button value="<%=arrButtons(2)%>" class="but_menu" style="width:90" onclick="window.close();" id=Button2 name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=submit value="<%=arrButtons(1)%>" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1></td></tr>
</form>
</table>
</td></tr> 
</table>
</BODY>


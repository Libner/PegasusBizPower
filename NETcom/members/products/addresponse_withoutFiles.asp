<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
    If Request.QueryString("add") <> nil Then
        prodID = trim(Request.Form("prodID"))
		If trim(Request.Form("responseID")) = "" Then ' add type
			sqlstr = "Insert into Product_Responses (Product_ID,Organization_ID,Response_Title,Response_Content) values (" &_
			prodID & "," & OrgID & ",'" & sFix(Request.Form("Response_Title")) & "','" & sFix(Request.Form("Response_Content")) & "')"
			'Response.Write sqlstr
			'Response.End
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.reload(true);
				window.close();
			//-->
			</SCRIPT>
		
	<%	Else ' update type
			responseID = trim(Request.Form("responseID"))
			sqlstr="Update Product_Responses set Response_Title = '" & sFix(Request.Form("Response_Title")) &_
			"', Response_Content = '" & sFix(Request.Form("Response_Content")) & "'" &_
			" Where Response_ID = " & responseID & " And Product_ID = " & prodID
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.reload(true);
				window.close();
			//-->
			</SCRIPT>	
	<%	End If
	End If
	
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 71 Order By word_id"				
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
			if(window.document.form1.Response_Title.value == "")
			{
			   <%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס מענה "
				Else
					str_alert = "Please insert the reply !!"
				End If   
				%>
				window.alert("<%=str_alert%>");				
				window.document.form1.Response_Title.focus();
				return false;
			}
			if(window.document.form1.Response_Content.value.length > 5000)
			{
				window.alert("התוכן שהזנת הינו גדול ממספר התוים המקסימלי");
				return false;
			}
			return true;			
		}
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<%
    prodID = trim(Request.QueryString("prodID"))
	If Request.QueryString("responseID") <> nil Then
		responseID = trim(Request.QueryString("responseID"))		
		If Len(responseID) > 0 Then
			sqlstr="Select Response_Title, Response_Content From Product_Responses Where Response_ID = " &_
			responseID & " And Product_ID = " & prodID
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				Response_Title = trim(rssub("Response_Title"))
				Response_Content = trim(rssub("Response_Content"))				
			End If
			set rssub=Nothing
		End If
	End If	
%>
<table dir="<%=dir_var%>" border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
<tr >
   <td  align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="1" cellspacing="0">
	   <tr>		 
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(responseID) > 0 Then%><span id=word1 name=word1><!--עדכון--><%=arrTitles(1)%></span><%Else%><span id="word2" name=word2><!--הוספת--><%=arrTitles(2)%></span><%End If%>&nbsp;<span id=word3 name=word3><!--מענה--><%=arrTitles(3)%></span>&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%>
<table width="100%" cellspacing="1" cellpadding="2" align=center border="0">
<form name=form1 id=form1 action="addresponse.asp?add=1" target="_self" method="post">
<input type=hidden name=responseID id=responseID value="<%=responseID%>">
<input type=hidden name=prodID id="prodID" value="<%=prodID%>">
<tr>
	<td align="<%=align_var%>" width=100% valign=top>
	<input type="text" class="texts" name="Response_Title" id="Response_Title" value="<%=vFix(Response_Title)%>" dir="<%=dir_obj_var%>" style="width:450" maxLength=100>	
	</td>
	<td width="70" nowrap align="<%=align_var%>" valign=top>&nbsp;<b><span id=word4 name=word4><!--כותרת--><%=arrTitles(4)%></span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=100% valign=top>
	<textarea class="texts" name="Response_Content" id="Response_Content" dir="<%=dir_obj_var%>" style="width:450" rows=9><%=(Response_Content)%></textarea>
	</td>
	<td width="70" nowrap align="<%=align_var%>" valign=top>&nbsp;<b><span id="word5" name=word5><!--תוכן--><%=arrTitles(5)%></span></b>&nbsp;</td>	
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

<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
    If Request.QueryString("add") <> nil Then
		If trim(Request.Form("SMS_type_id")) = "" Then ' add type
			sqlstr = "Insert into SMS_types (SMS_type_name,SMS_type_text) values ('" & sFix(Request.Form("SMS_type_name")) & "','"&  sFix(Request.Form("sms_text"))& "')"
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.href = window.opener.document.location.href;
				window.close();
			//-->
			</SCRIPT>
		
	<%	Else ' update type
			SMS_type_id = trim(Request.Form("SMS_type_id"))
			sqlstr="Update SMS_types set SMS_type_name = '" & sFix(Request.Form("SMS_type_name")) & "',SMS_type_text='"& sFix(Request.Form("sms_text")) & "'  Where SMS_type_id = " & SMS_type_id
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
			if(window.document.form1.SMS_type_name.value == "")
			{
			   <%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס שם סוג "
				Else
					str_alert = "Please insert the type name !!"
				End If   
				%>
				window.alert("<%=str_alert%>");				
				window.document.form1.SMS_type_name.focus();
				return false;
			}
			return true;			
		}
		function checkLen(object, maxChars, divNum, langFlip, langFlipDec)
{
	if (langFlip)
	{
		charToDec = 126 - maxChars;
		if (object.value.search(/[^\x00-\x7E]/) >= 0) ;
		else
		{
			if (langFlipDec) charToDec -= langFlipDec;
			maxChars = 126 - charToDec;
		}
	}	

	elementId = "charCount" + divNum;
	//document.getElementById(elementId).style.display="block";	

	if (object.value.length > maxChars) 
	{
		object.value = object.value.substring(0,maxChars); 
		object.focus();
	}

	leftChars = maxChars - object.value.length;
	if (leftChars < 0) leftChars = 0;	
	document.getElementById(elementId).innerHTML = "נותרו לכם <font color='red'>" + leftChars + "</font> תווים";
	
	return true;	
}

//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<%
	If Request.QueryString("SMS_type_id") <> nil Then
		SMS_type_id = trim(Request.QueryString("SMS_type_id"))
		If Len(SMS_type_id) > 0 Then
			sqlstr="Select SMS_type_name,SMS_type_text From SMS_types Where SMS_type_id = " & SMS_type_id
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				SMS_type_name = trim(rssub("SMS_type_name"))		
				sms_text = trim(rssub("SMS_type_text"))		
			End If
			set rssub=Nothing
		End If
	End If	
%>
<table dir="<%=dir_var%>" border="0" width="480" cellspacing="0" cellpadding="0" align="center" ID="Table1">
<tr >
   <td  align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="1" cellspacing="0" ID="Table2">
	   <tr>		 
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(item_id) > 0 Then%><span id=word1 name=word1><!--עדכון--><%=arrTitles(1)%></span><%Else%><span id="word2" name=word2><!--הוספת--><%=arrTitles(2)%></span><%End If%>&nbsp;<span id=word4 name=word4><!--סוג--><%=arrTitles(4)%></span></td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%>
<table width="480" cellspacing="1" cellpadding="2" align=center border="0" dir="<%=dir_var%>" ID="Table3">
<form name=form1 id=form1 action="addSMSType.asp?add=1" target="_self" method="post">
<input type=hidden name=SMS_type_id id=SMS_type_id value="<%=SMS_type_id%>">
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" dir="<%=dir_obj_var%>" class="texts" name="SMS_type_name" id="SMS_type_name" value="<%=vFix(SMS_type_name)%>" dir=rtl size=100 maxLength=100>	
	</td>
	<td width="150" nowrap align="right">&nbsp;<b><span id=word3 name=word3><!--שם סוג--><%=arrTitles(3)%></span></b>&nbsp;</td>	
</tr>
<tr>
	<td align="right" width=330 nowrap><div id="charCount1"></div>
	<input dir="<%=dir_obj_var%>" type="text" class="texts" style="width:320" id="sms_text" name="sms_text" 
	value="<%=vFix(sms_text)%>" MaxLength="126" onkeyup="return checkLen(this, 126, 1, 1);" onchange="return checkLen(this, 126, 1, 1);"></td>
	<td width="150" nowrap align="right">&nbsp;<b>תוכן ההודעה </b>&nbsp;</td>	
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

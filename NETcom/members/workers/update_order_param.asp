<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%If Request.QueryString("add") <> nil Then
        Month_Min_Order = nFix(DBValueToStr(Request.Form("Month_Min_Order")))
        Order_Price = nFix(DBValueToStr(Request.Form("Order_Price")))	
    
       If isNumeric(Month_Min_Order) Then
			sqlstr="UPDATE Users SET Month_Min_Order = " & (Month_Min_Order)
			con.executeQuery(sqlstr) 
		End If
		
       If IsNumeric(Order_Price) Then
			sqlstr="UPDATE Users SET Order_Price = " & Order_Price & " "
			con.executeQuery(sqlstr) 
		End If  %>
		<SCRIPT LANGUAGE="javascript">
		<!--
			window.opener.document.location.href = window.opener.document.location.href;
			window.close();
		//-->
		</SCRIPT>	
<%End If%>
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
	  set rsbuttons=nothing	 	  %>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
		function checkForm()
		{		
				
			return true;			   
		}
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<table border="0" width="380" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>">
<tr >
   <td  align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="1" cellspacing="0">
	   <tr>		 
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%">&nbsp;עדכון יעד/עלות הזמנות&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width="100%">
<form name="form1" id="form1" action="update_order_param.asp?add=1" target="_self" method="post">
<table width="380" cellspacing="1" cellpadding="2" align="center" border="0">
<tr>
	<td align="<%=align_var%>" width="230" nowrap><input type="text" name="Month_Min_Order" id="Month_Min_Order" 
	class="form" value="<%=vFix(Month_Min_Order)%>" dir="<%=dir_obj_var%>" style="width: 100px" maxLength="3"></td>
	<td width="150" nowrap align="<%=align_var%>">&nbsp;<b>יעד הזמנות מינימלי</b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width="230" nowrap><input type="text" name="Order_Price" id="Order_Price" 
	class="form" value="<%=vFix(Order_Price)%>" dir="<%=dir_obj_var%>" style="width: 100px" maxLength="3"></td>
	<td width="150" nowrap align="<%=align_var%>">&nbsp;<b>שווי כספי של הזמנה</b>&nbsp;</td>	
</tr>
<tr><td height="35" colspan="2" nowrap></td></tr>
<tr><td align="center" colspan="2">
<input type=button value="<%=arrButtons(2)%>" class="but_menu" style="width:90" onclick="window.close();" id=Button2 name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=submit value="<%=arrButtons(1)%>" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1></td></tr>
</table>
</form>
</td></tr></table>
</BODY>
</html>
<%Set con = Nothing%>
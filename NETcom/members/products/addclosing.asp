<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%   If Request.QueryString("add") <> nil Then
        prodID = trim(Request.Form("prodID"))
        bgrColor = trim(Request.Form("bgrColor"))
		If trim(Request.Form("closingID")) = "" Then ' add type
			sqlstr = "Insert into Product_Closings (Product_ID,Organization_ID,Closing_Name,Closing_Color) " &_
			" Values (" & prodID & "," & OrgID & ",'" & sFix(Request.Form("Closing_Name")) & "','" & bgrColor & "')"
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
			closingID = trim(Request.Form("closingID"))
			sqlstr="Update Product_Closings set Closing_Name = '" & sFix(Request.Form("Closing_Name")) &_
			"',  Closing_Color = '" & sFix(bgrColor) & "' Where Closing_ID = " & closingID & " And Product_ID = " & prodID
			con.executeQuery(sqlstr) %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				window.opener.document.location.reload(true);
				window.close();
			//-->
			</SCRIPT>	
	<%	End If
	End If

	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 70 Order By word_id"				
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
	set rsbuttons=nothing%>
<html>
<head>
<!--#include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css" />
<SCRIPT LANGUAGE=javascript>
<!--
		function checkForm()
		{
			if(window.document.form1.Closing_Name.value == "")
			{
			   <%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס מצב סגירה "
				Else
					str_alert = "Please insert the closing status !!"
				End If   
				%>
				window.alert("<%=str_alert%>");				
				window.document.form1.Closing_Name.focus();
				return false;
			}
			return true;			
		}
		
		function setBgrColor()
		{
			var oldcolor = document.all.bgrColor.value;
			var newcolor = showModalDialog("../../../htmlarea/popups/select_color.html", oldcolor, "resizable: no; help: no; status: no; scroll: no;");
			if (newcolor != null) { document.all.bgrColor.value = "#"+newcolor; document.all.imgbgrColor.style.backgroundColor = "#"+newcolor;}
		}
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<%  prodID = trim(Request.QueryString("prodID"))
	If Request.QueryString("closingID") <> nil Then
		closingID = trim(Request.QueryString("closingID"))		
		If Len(closingID) > 0 Then
			sqlstr="Select Closing_Name, Closing_Color From Product_Closings Where Closing_ID = " &_
			closingID & " And Product_ID = " & prodID
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				Closing_Name = trim(rssub("Closing_Name"))
				bgrColor = trim(rssub("Closing_Color"))
			End If
			set rssub=Nothing
		End If
	Else
		bgrColor = "#53FE01"			
	End If	%>
<table dir="<%=dir_var%>" border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
<tr >
   <td  align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="1" cellspacing="0">
	   <tr>		 
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(closingID) > 0 Then%><!--עדכון--><%=arrTitles(4)%><%Else%><!--הוספת--><%=arrTitles(9)%><%End If%>&nbsp;<!--מצב סגירה--><%=arrTitles(8)%>&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%><form name=form1 id=form1 action="addclosing.asp?add=1" target="_self" method="post">
<input type=hidden name=closingID id=closingID value="<%=closingID%>">
<input type=hidden name=prodID id="prodID" value="<%=prodID%>">
<table width="100%" cellspacing="1" cellpadding="2" align=center border="0">
<tr>
	<td align="<%=align_var%>" width=100% valign=top>
	<input type="text" class="texts" name="Closing_Name" id="Closing_Name" value="<%=vFix(Closing_Name)%>" dir="<%=dir_obj_var%>" style="width:250" maxLength="50">	
	</td>
	<td width="70" nowrap align="<%=align_var%>" valign=top>&nbsp;<b><!--כותרת--><%=arrTitles(8)%></b>&nbsp;</td>	
</tr>
<tr>
	<td align="<%=align_var%>" width=100% valign=top>
	<input type="text" name="bgrColor" value="<%=bgrColor%>" size=7 style="height: 22px; left: 100px; top: 15px;" maxlength=7 ID="Text1">
	<BUTTON style="top: 15px;left:160px ;width: 22px; height: 22px; border: 1px none;"  unselectable="on" onclick="setBgrColor();" id=button4 name=button4><img id=imgbgrColor style="background-color:<%=bgrColor%>;" src="<%=Application("VirDir")%>/netcom/images/selcolor.gif" border=0></BUTTON>
	</td>
	<td width="70" nowrap align="<%=align_var%>" valign=top>&nbsp;<b><!--צבע--><%=arrTitles(8)%></b>&nbsp;</td>	
</tr>
<tr><td height=5 colspan="2" nowrap></td></tr>
<tr><td align=center colspan="2">
<input type=button value="<%=arrButtons(2)%>" class="but_menu" style="width:90" onclick="window.close();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=submit value="<%=arrButtons(1)%>" class="but_menu" style="width:90" onclick="return checkForm()"></td></tr>
</table></form>
</td></tr></table>
</BODY>
</HTML>
<%Set con = Nothing%>
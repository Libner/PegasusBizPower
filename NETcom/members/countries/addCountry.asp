<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%  relCountry= Request.Form("task_types")
    If Request.QueryString("add") <> nil Then
		If trim(Request.Form("Country_Id")) = "" Then ' add type
			sqlstr = "Insert into Countries (Country_Name) values ('"  & sFix(Request.Form("Country_Name")) & "')"
		   con.executeQuery(sqlstr)
		   sql="SELECT top 1 Country_Id  from Countries  order by Country_Id desc"
		set rs_tmp = con.getRecordSet(sql)
		if not rs_tmp.eof then
				Country_Id = rs_tmp("Country_Id")
			end if
			
		set rs_tmp = Nothing
		
			
				If Len(trim(relCountry)) > 0 Then 'לשלוח במייל משימה למותבים
		arrCountries = Split(relCountry, ",")
			If IsArray(arrCountries) Then
				For count=0 To Ubound(arrCountries)
				   If IsNumeric(arrCountries(count)) Then
						sqlstr="Insert Into Country_ToPegasus (Country_Id,Country_Pegasus) Values (" & Country_Id & "," & arrCountries(count) & ")"	  	
						con.executeQuery(sqlstr)		
				   End If				
				Next
			End If
		End If			
				
			 %>
			<SCRIPT LANGUAGE=javascript>
			<!--
				opener.focus();
				opener.window.location.reload();
				self.close();
			//-->
			</SCRIPT>
		
	<%	Else ' update type
			Country_Id = trim(Request.Form("Country_Id"))
			sqlstr="Update Countries set Country_Name = '" & sFix(Request.Form("Country_Name")) & "' Where Country_Id = " & Country_Id
			con.executeQuery(sqlstr)
			sqlRel="delete from Country_ToPegasus where Country_Id=" & Country_Id
			con.executeQuery(sqlRel)
			
		If Len(trim(relCountry)) > 0 Then 'לשלוח במייל משימה למותבים
		arrCountries = Split(relCountry, ",")
			If IsArray(arrCountries) Then
				For count=0 To Ubound(arrCountries)
				   If IsNumeric(arrCountries(count)) Then
						sqlstr="Insert Into Country_ToPegasus (Country_Id,Country_Pegasus) Values (" & Country_Id & "," & arrCountries(count) & ")"	  	
						con.executeQuery(sqlstr)		
				   End If				
				Next
			End If
		End If		
				 %>
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
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 32 Order By word_id"				
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
		
			if(window.document.form1.Country_Name.value == "")
			{
			
				<%
				If trim(lang_id) = "1" Then
					str_alert = "!!נא להכניס יעד "
				Else
					str_alert = "Please insert the group name !!"
				End If   
				%>
				window.alert("<%=str_alert%>");
				window.document.form1.Country_Name.focus();
				return false;
			}
		//	alert(window.document.form1.task_types.length)
			var i;
			
for (i = 0; i < window.document.form1.task_types.length; i++) {
  if (window.document.form1.task_types[i].checked) {
    flCkeck=1;
    return true;
  }
  else
  flCkeck=0;
  }
if (flCkeck==0)
{
 <%
			If trim(lang_id) = "1" Then
				str_alert = "! נא לבחור מדינה"
			Else
				
			End If	
	  %>
 	  window.alert("<%=str_alert%>");     
      return false; 
}
			
	/* if(window.document.all("task_types").value=='')
   {
      <%
			If trim(lang_id) = "1" Then
				str_alert = "! נא לבחור מדינה"
			Else
				
			End If	
	  %>
 	  window.alert("<%=str_alert%>");     
      return false; 
   }*/
			return true;			
		}
//-->
</SCRIPT>
</HEAD>
<body style="margin:0px;background:#E5E5E5" onload="window.focus()">     
<form name=form1 id=form1 action="addCountry.asp?add=1" target="_self" method="post">

<%
	If Request.QueryString("Country_Id") <> nil Then
		Country_Id = trim(Request.QueryString("Country_Id"))
		If Len(Country_Id) > 0 Then
			sqlstr="Select Country_Name From Countries Where Country_Id = " & Country_Id
			set rssub = con.getRecordSet(sqlstr)
			If not rssub.eof Then
				Country_Name = trim(rssub("Country_Name"))				
			End If
			set rssub=Nothing
		End If
	End If	
%>
<table dir="<%=dir_var%>" border="0" width="100%" cellspacing="0" cellpadding="0" align="center" ID="Table1">
<tr >
   <td  align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="1" cellspacing="0" ID="Table2">
	   <tr>	
	   <td class="page_title" align=left>
<input type=button value="<%=arrButtons(2)%>" class="but_menu" style="width:90" onclick="window.close();" id=Button2 name=Button2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=submit value="<%=arrButtons(1)%>" class="but_menu" style="width:90" onclick="return checkForm()" id=Button1 name=Button1></td>	 
		 <td class="page_title" align="<%=align_var%>" valign="middle" width="100%"><%If Len(Country_Id) > 0 Then%><span id=word1 name=word1><!--עדכון--><%'=arrTitles(1)%>עדכון יעד</span><%Else%><span id="word2" name=word2>הוספת יעד<!--הוספת--><%'=arrTitles(2)%></span><%End If%>&nbsp;<span id=word4 name=word4><!--קבוצה--><%'=arrTitles(4)%></span>&nbsp;</td>
	  </tr>
	 </table>
  </td>
</tr>
<tr><td height=15 nowrap></td></tr>
<tr><td width=100%>
<table width="480" cellspacing="1" cellpadding="2" align=center border="0" ID="Table3">
<input type=hidden name=Country_Id id=Country_Id value="<%=Country_Id%>">
<tr>
	<td align="<%=align_var%>" width=330 nowrap>
	<input type="text" class="texts" name="Country_Name" id="Country_Name" value="<%=vFix(Country_Name)%>" dir="<%=dir_obj_var%>" size=100 maxLength=100>	
	</td>
	<td width="150" nowrap align="center">&nbsp;<b><span id=word3 name=word3>יעד<!--שם קבוצה--><%'=arrTitles(3)%></span></b>&nbsp;</td>	
</tr>
<!--site country-->
<tr>
								<td align="right" nowrap class="td_admin_4" valign="top" dir=rtl>
								<table border=0 cellpadding=0 cellspacing=0>
				
							  <%	
							  if Country_Id="" then
							  sqlstr="SELECT   Country_Id,Country_Name,dbo.IsSelected (Country_Id,0) as IsSelected from " & pegasusDBName & ".dbo.Countries order by Country_Name"

							  else
							  
							  sqlstr="SELECT   Country_Id,Country_Name,dbo.IsSelected (Country_Id,"& Country_Id &") as IsSelected from " & pegasusDBName & ".dbo.Countries order by Country_Name"
		end if
				set rssub = con.getRecordSet(sqlstr)		 
		  	If not rssub.eof Then
		  	   arrSub = rssub.getRows()
		  	   recCount = rssub.RecordCount	
		  	   set rssub=Nothing
		  	   i=0			
		  	While i<=Ubound(arrSub,2)
		  		selId=trim(arrSub(0,i))
		  		selName=trim(arrSub(1,i))
		  		is_selected=trim(arrSub(2,i))
		  %>
		    <tr>
		  <%If i<=Ubound(arrSub,2) Then%>		 
		  <td width=15  align="center" nowrap><input type=checkbox id="task_types" name="task_types" <%If is_selected > 0 Then%> checked <%End If%> value="<%=selId%>"></td>		 
		  <td align="<%=align_var%>" width=170 nowrap>&nbsp;<%=selName%>&nbsp;</td>
		  <%Else%>
		  <td>&nbsp;</td><td>&nbsp;</td>
		  <%End If	
		        i = i+1	
		        If i<=Ubound(arrSub,2) Then		  	    
		  	    selId=trim(arrSub(0,i))
		  		selName=trim(arrSub(1,i))
		  		is_selected=trim(arrSub(2,i))		  		
		  %>	   
		   <td width=15  align="center" nowrap><input type=checkbox id="task_types" name="task_types" <%If is_selected > 0 Then%> checked <%End If%> value="<%=selId%>"></td>									  
		   <td align="<%=align_var%>" width=170 nowrap>&nbsp;<%=selName%>&nbsp;</td>
		    <%Else%>
		  <td>&nbsp;</td><td>&nbsp;</td>
		  <%End If%>
		   </tr>	
		  <%  i = i+1		
		  	  Wend		  	
			  End If		
			
	
          %>         
</table>
								</td>
								<td width="150" nowrap align="center" dir=rtl valign=top>מדינות (אתר)</td>
							</tr>
							<tr>
								<td bgcolor="#e6e6e6" height="10"></td>
							</tr>
<!--site country-->
<tr><td height=5 colspan="2" nowrap></td></tr>
</form>
</table>
</td></tr> </table>
</BODY>
</HTML>

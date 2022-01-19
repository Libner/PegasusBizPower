<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<script language="javascript" type="text/javascript" src="../../../Scripts/CalendarPopupH.js"></script>


<script language="javascript" type="text/javascript" >
<!--
function callCalendar(pf,pid)
	{

	cal1xx.select(pf,pid,'dd/MM/yyyy')
	}

function CheckFields(action)
{	
	if (document.frmMain.closeDate.value=='')		   	
	{
		<%
			If trim(lang_id) = "1" Then
				str_alert = "אנא הכנס תאריך סגירה"
			Else
				str_alert = "Please insert the closing date"
			End If	
		%>
		window.alert("<%=str_alert%>");
		return false;				
	}
	
	document.frmMain.submit();
	return true;

}

function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=157pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}
//-->
</script>  
<%appID = cLng(Request.QueryString("appID"))
  
	if appID<>nil and appID<>"" then
	'sqlStr = "select appeal_close_date, appeal_close_text from appeals where appeal_ID=" & appID  
	sqlstr = "EXECUTE dbo.get_appeals '','','','','" & OrgID & "','','','','','','','" & appID & "'"
	''Response.Write sqlStr
	set rs_appeals = con.GetRecordSet(sqlStr)
		if not rs_appeals.eof then
			If IsDate(rs_appeals("appeal_close_date")) Then
			closeDate = FormatDateTime(rs_appeals("appeal_close_date"),2)
			Else
			closeDate = Date()
			End If
			closeText = rs_appeals("appeal_close_text")
			companyId = rs_appeals("company_id")
			contactId = rs_appeals("contact_id")
			quest_id = rs_appeals("questions_id")				
		end if
		set rs_appeals = nothing	 
		If IsDate(closeDate) = false Then
			closeDate = FormatDateTime(Date(),2)   
		End If	
		
		If trim(contactId) <> "" Then
			sqlstr = "Select email from contacts WHERE contact_Id = " & contactId
			set rs_pr = con.getRecordSet(sqlstr)
			If not rs_pr.eof Then
				reponse_email = trim(rs_pr(0))				
			End If
			set rs_pr = Nothing
		End If
	end if
	'Response.Write company_id  

  if request.form("closeDate")<>nil then 'after form filling		
	
	  closeDate = "'" & Request.Form("closeDate") & "'"	
	  closeContent = sFix(trim(Request.Form("closeText")))
	  Closing_ID = trim(Request.Form("Closing_ID"))
	 
	  con.executeQuery("SET DATEFORMAT DMY")
	  sqlStr = "Update appeals set appeal_close_date = " & closeDate &_
	  ", appeal_status = '3', appeal_close_text = '" & closeContent &_
	  "', response_user_id = " & UserID & ", Closing_ID = '" & Closing_ID &_
	  "' Where appeal_ID=" & appID
	  con.GetRecordSet (sqlStr) 	%>	
		<script language="javascript" type="text/javascript" >
			<!--
			//window.opener.document.location.href = "appeals.asp?prodID=<%=quest_id%>";
			window.opener.history.back();
			window.close();
			//-->
		</script>
   <%End If
   
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 63 Order By word_id"				
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
	  set rstitle = Nothing   %>
<body style="margin:0px; background-color:#E6E6E6" onload="window.focus();">
<table border="0" cellpadding="0" cellspacing="1" width="100%" dir="<%=dir_var%>">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title" dir="<%=dir_obj_var%>">&nbsp;<!--סגירה--><%=arrTitles(1)%></td></tr>		   
   </table>
</td></tr> 
<tr><td width="100%">
<FORM name="frmMain" ACTION="appeal_close.asp?appID=<%=appID%>" METHOD="post" onSubmit="return CheckFields()">
<table align=center border="0" cellpadding="3" cellspacing="1" width="98%" align=center>
<tr><td height=10></td></tr>          
<tr>
	<td align="<%=align_var%>" width="100%">	
		<a href='' onclick='callCalendar(document.frmMain.closeDate,"AscloseDate");return false;' id="AscloseDate">
						<image src='../../images/calend.gif' border=0 ></a>&nbsp;<input dir="ltr" type="text" class="passw" size=10 id="closeDate" name="closeDate" value="<%=closeDate%>" maxlength=10 readonly>&nbsp;

	
	</td>
	<td align="<%=align_var%>" nowrap width="100">&nbsp;<!--תאריך סגירה--><%=arrTitles(2)%>&nbsp;</td>
</tr>
<%
sqlstr = "Select Closing_ID, Closing_Name FROM Product_Closings WHERE ORGANIZATION_ID = " & OrgID &_			
" And Product_ID = " & quest_id & " Order BY Closing_Name"
set rs_tmp = con.getRecordSet(sqlstr)
If not rs_tmp.eof Then
	str_tmp = rs_tmp.getstring(,,"'>", "</option><option value='", "")
	If Len(str_tmp) > 15 Then
			str_tmp = Left(str_tmp, Len(str_tmp)-15)
	End If
End If	
set rs_tmp = Nothing

If Len(str_tmp) > 0 Then%>
<tr>
	<td align="<%=align_var%>" width="100%">	
	<select name="Closing_ID" id="Closing_ID" class="norm" dir="<%=dir_obj_var%>">	
	<option value='<%Response.Write(str_tmp)%>	
	</select>	
	</td>
	<td align="<%=align_var%>" nowrap width="100">&nbsp;<!--מצב סגירה--><%=arrTitles(15)%>&nbsp;</td>
</tr>
<%End If%>
<tr valign=top>
	<td align="<%=align_var%>" width="100%">	
	<textarea dir="<%=dir_obj_var%>" type="text" class="texts" style="width:450" rows=9 id="closeText" name="closeText"><%=closeText%></textarea>
	</td>
	<td align="<%=align_var%>" nowrap width="100">&nbsp;<!--תוכן סגירה--><%=arrTitles(3)%>&nbsp;</td>
</tr>
<tr><td colspan="2" height=10></td></tr>
<tr>
<td colspan=2>
<table cellpadding=0 cellspacing=0 border=0 align=center width=80% align=center>
<tr valign=top>
<td width=50% align=left><A class="but_menu" href="#" style="width:90px" onclick="window.close();"><span id=word4 name=word4><!--ביטול--><%=arrTitles(4)%></span></A></td>
<td width=10 nowrap></td>
<td width=50% align=center><A class="but_menu" style="width:90px" href="#" onclick="return CheckFields(1);"><span id="word5" name=word5><!--סגור--><%=arrTitles(5)%></span></A></td>
<td width=10 nowrap></td>
</tr>
</table></td></tr>
<tr><td colspan="2" height=5></td></tr>
</table></form>
</td></tr></table>
   	<SCRIPT LANGUAGE="JavaScript">document.write(getCalendarStyles());</SCRIPT>
			<SCRIPT LANGUAGE="JavaScript" type="text/javascript">
            <!--
            var cal1xx = new CalendarPopup('CalendarDiv');
                cal1xx.showNavigationDropdowns();
                cal1xx.yearSelectStart
                cal1xx.offsetX = -50;
                cal1xx.offsetY = 0;
            //-->
			</SCRIPT>
					<DIV ID='CalendarDiv' name='CalendarDiv' STYLE='VISIBILITY:hidden;POSITION:absolute;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>
			

</body> 
</html>
<%set con = nothing%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%CRMCountryId = trim(Request.QueryString("CRMCountryId"))
contactID= trim(Request.QueryString("contactID"))
      If IsNumeric(CRMCountryId) Then
		Set rs = con.getRecordSet("SELECT Country_Name FROM dbo.Countries WHERE (Country_Id = " & cInt(CRMCountryId) & ")")
		If Not rs.EOF Then
			Country_Name = trim(rs(0))
		End If
		Set rs = Nothing
      End If
	


	if Request("Page")<>"" then
	Page=request("Page")
	else
		Page=1
	end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">	  
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE="javascript" type="text/javascript">
<!--
	function choose(AppealId, CountryName)
	{		
		if(window.opener.document.getElementById("RelationId") != null)
		{
			window.opener.document.getElementById("RelationId").value = AppealId;
		}	
			
		if(window.opener.document.getElementById("RelationName") != null)	
		{
			window.opener.document.getElementById("RelationName").value = unescape(CountryName)+'/'+ AppealId;
		}	
			
	
		
		self.close();
		return false;
	}

	function submitenter(myfield,e)
	{
		var keycode;
		if (window.event) keycode = window.event.keyCode;
		else if (e) keycode = e.which;
		else return true;

		if (keycode == 13)
		{
			myfield.form.submit();
			return false;
		}
		else
		return true;
	}
//-->
</SCRIPT>
</HEAD>
<BODY style="margin:0px" onload="window.focus()">
<table width="100%" align="center" cellpadding=0 cellspacing=0 border=0 ID="Table1">
<tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table2">
<tr>
	<td align="<%=align_var%>" class="page_title" width="100%" style="border-left: none">&nbsp;&nbsp;טפסי מתעניין לטיול שנפתחו ב-3 חודשים אחרונים<%'=Country_Name%></td>
</tr>
</table>
</td></tr>
<TR>
<td>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" ID="Table3">
<tr>
<td width="100%" bgcolor="#E4E4E4" align="middle" valign="top"><!-- start code --> 
		<FORM action="listInterestedForm.asp?CRMCountryId=<%=CRMCountryId%>" method="post" id="form_search" name="form_search" style="margin: 0px;">
		<table align="<%=align_var%>" border="0" cellpadding="1" cellspacing="1" bgcolor=white width="100%" dir="<%=dir_var%>" ID="Table4">
		
		<tr>						
			<td class="title_sort" align="center" width="5%">&nbsp;בחר&nbsp;</td>
			<td class="title_sort" align="center" width="20%">&nbsp;יעד&nbsp;</td>			
			<td class="title_sort" align="center" width="20%">&nbsp;הוזן ע"י&nbsp;</td>	
			<td class="title_sort" align="center" width="20%">&nbsp;חודש יציאה&nbsp;</td>	
			<td class="title_sort" align="center" width="20%">&nbsp;תוכן הפנייה&nbsp;</td>	
			<td class="title_sort" align="center" width="20%">&nbsp;ID&nbsp;</td>			
			<td class="title_sort" align="center" nowrap width="55%">&nbsp;תאריך&nbsp;</td>			
		</tr>
		
<%sqlstr = "SELECT APPEAL_ID, appeal_CountryId, APPEAL_DATE,Country_Name, (FIRSTNAME + Char(32) + LASTNAME) as USER_NAME,dbo.Get_FieldValue40103(APPEAL_ID) as FieldValue40103,dbo.Get_FieldValue40109(APPEAL_ID) as FieldValue40109  FROM dbo.APPEALS  left join Users  on APPEALS.User_Id=Users.User_Id  left join Countries on appeal_CountryId=Countries.Country_Id WHERE  (appeal_CountryId = " & cInt(CRMCountryId) & ")  and  CONTACT_ID=" & CONTACTID & "  and  QUESTIONS_ID=16504 and (APPEAL_STATUS=1 or APPEAL_STATUS=2 or APPEAL_STATUS=4) and APPEAL_DATE >= DATEADD(month, -3, GETDATE()) ORDER BY APPEAL_DATE desc"	
'response.Write sqlstr
'response.end
	set pr=con.GetRecordSet(sqlstr)    
	if not pr.eof then 
		recnom = 1		
		pr.PageSize = 1000			
		pr.AbsolutePage=Page		
		NumberOfPages = pr.PageCount
		page_size = pr.PageSize
		prArray = pr.GetRows(page_size)	
		recCount=pr.RecordCount 	
		pr.Close
		set pr=Nothing    
		i=1
		j=0
		do while (j<=ubound(prArray,2))			
			APPEAL_ID = trim(prArray(0,j))
			APPEAL_DATE = trim(prArray(2, j))					
			Country_Name = trim(prArray(3, j))	
			User_Name = trim(prArray(4, j))	
			field40103 = trim(prArray(5, j))	
			field40109 = trim(prArray(6, j))	
			
			%>	
		<tr bgcolor="#E6E6E6">			
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data"><input type="button" ID="btnSend" NAME="btnSend" value="בחר" onclick="return choose('<%=APPEAL_ID%>', escape('<%=altFix(Country_Name)%>'));" class="but_site"></td>
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data"><A class="link_categ" onclick="return choose('<%=APPEAL_ID%>', escape('<%=altFix(Country_Name)%>'));" href="javascript:void(0)">&nbsp;<%=Country_Name%>&nbsp;</a></td>			
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data"><A class="link_categ" onclick="return choose('<%=APPEAL_ID%>', escape('<%=altFix(Country_Name)%>'));" href="javascript:void(0)">&nbsp;<%=User_Name%>&nbsp;</a></td>
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data"><A class="link_categ" onclick="return choose('<%=APPEAL_ID%>', escape('<%=altFix(Country_Name)%>'));" href="javascript:void(0)">&nbsp;<%=field40109%>&nbsp;</a></td>
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data"><A class="link_categ" onclick="return choose('<%=APPEAL_ID%>', escape('<%=altFix(Country_Name)%>'));" href="javascript:void(0)">&nbsp;<%=field40103%>&nbsp;</a></td>
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data"><A class="link_categ" onclick="return choose('<%=APPEAL_ID%>', escape('<%=altFix(Country_Name)%>'));" href="javascript:void(0)">&nbsp;<%=APPEAL_ID%>&nbsp;</a></td>
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data"><A class="link_categ" onclick="return choose('<%=APPEAL_ID%>', escape('<%=altFix(Country_Name)%>'));" href="javascript:void(0)">&nbsp;<%=APPEAL_Date%>&nbsp;</a></td>
		</tr>
	<%	'pr.movenext
		recnom = recnom + 1
		j = j + 1
		loop%>			
	   
	<tr><td class="card" align="center" colspan="7" style="color:#5F435A;font-weight:600" dir="ltr">נמצאו&nbsp;<%=recCount%>&nbsp;רשומות</td></tr>					
<%Else 'not pr.eof%>
	<tr><td colspan=7 height=25></td></tr>
	<tr><td align=center dir=rtl colspan="7" style="color:#5F435A;font-size:14pt;font-weight:bold">לא ניתן לשמור הטופס כיוון שאין טופס מתעניין בטיול שתואם את טופס רישום החתום, לכן, עליך </td></tr>
	<tr><td colspan=7 height=15></td></tr>
	<tr><td align=center dir=rtl colspan="7" >
	<table border=0 cellpadding=0 cellspacing=7>

	<TR><td  align=right dir=rtl style="color:#5F435A;font-size:12pt;">1) לעדכן את מנהל המכירות שהתנהלות נציג לא בוצעה כהלכה</td></tr>
	<tr><td align=right dir=rtl style="color:#5F435A;font-size:12pt;">2) יש להסביר לנציג ולתקנו כדי שילמד להבא</td></tr>
	<tr><td align=right dir=rtl style="color:#5F435A;font-size:12pt;">3) יש לפתוח טופס מתעניין בטיול לפי הנוהל ורק אז ללחוץ על אישור</td></tr>
	</table> 
		
	
</td></tr>
	
<%End If 'not pr.eof %>
<!-- end code -->
</form></table>
</td></tr></table>
</td></tr></table>
</BODY>
</HTML>
<%set con =nothing%>
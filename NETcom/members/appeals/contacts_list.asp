<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%CompanyId = trim(Request.QueryString("CompanyId"))
      If IsNumeric(CompanyId) Then
		Set rs = con.getRecordSet("SELECT Company_Name FROM dbo.Companies WHERE (Company_Id = " & cInt(CompanyId) & ")")
		If Not rs.EOF Then
			CompanyName = trim(rs(0))
		End If
		Set rs = Nothing
		where_company_id = " AND Company_Id =  " & CompanyId
	  Else
		where_company_id = ""	
      End If
	
	if trim(Request.Form("contact_name")) <> "" Or trim(Request.QueryString("contact_name")) <> "" then
		contact_name = trim(Request.Form("contact_name"))
		if trim(Request.QueryString("contact_name")) <> "" then
			contact_name = trim(Request.QueryString("contact_name"))
		end if
		is_search = true
		
		where_contact_name = " and UPPER(CONTACTS.contact_name) LIKE '%"& UCase(sFix(contact_name)) &"%'"
	else
		where_contact_name = ""
	end if
	
	if trim(Request.Form("contact_cellular")) <> "" Or trim(Request.QueryString("contact_cellular")) <> "" then
		contact_cellular = trim(Request.Form("contact_cellular"))
		if trim(Request.QueryString("contact_cellular")) <> "" then
			contact_cellular = trim(Request.QueryString("contact_cellular"))
		end if
		is_search = true
		
		where_contact_cellular = " AND "
		If Len(where_contact_name) > 1 Then
			where_contact_cellular = " OR "
		End If		
		where_contact_cellular = where_contact_cellular & " UPPER(CONTACTS.cellular) LIKE '"& UCase(sFix(contact_cellular)) &"%'"
	else
		where_contact_cellular = ""
	end if
	
	if Request("Page")<>"" then
	Page=request("Page")
	else
		Page=1
	end if
	if Request.QueryString("all") <> nil then
		all = 1
		show_all = true		
	else
		all = 0
		show_all = false
	end if	%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">	  
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html;">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE="javascript" type="text/javascript">
<!--
	function choose(ContactId, ContactName)
	{		
		if(window.opener.document.getElementById("contactID") != null)
		{
			window.opener.document.getElementById("contactID").value = ContactId;
		}	
			
		if(window.opener.document.getElementById("contactName") != null)	
		{
			window.opener.document.getElementById("contactName").value = unescape(ContactName);
		}	
			
		if(window.opener.document.getElementById("contacter_body") != null)	
		{
			window.opener.document.getElementById("contacter_body").style.display = "inline";
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
<table width="100%" align="center" cellpadding=0 cellspacing=0 border=0>
<tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
	<td align="<%=align_var%>" class="page_title" nowrap width="70" style="border-right: none"><A class="but_menu" href='contacts_list.asp?all=1&amp;CompanyId=<%=CompanyId%>' target="_self" style="width: 70px;" >הצג הכול</a></td>
	<td align="<%=align_var%>" class="page_title" nowrap width="70" style="border-right: none"><A class="but_menu" href="#" onclick="form_search.submit();"style="width:100%">חפש</a></td>
	<td align="<%=align_var%>" class="page_title" width="100%" style="border-left: none">&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsMulti"))%>&nbsp;<%=CompanyName%></td>
</tr>
</table>
</td></tr>
<TR>
<td>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="100%" bgcolor="#E4E4E4" align="middle" valign="top" align="center"><!-- start code --> 
		<FORM action="contacts_list.asp?CompanyId=<%=CompanyId%>" method=post id=form_search name=form_search>
		<table align="center" border="0" cellpadding="1" cellspacing="1" bgcolor=white width="99.5%" dir="<%=dir_var%>">
		<tbody bgcolor="#E4E4E4">
		<tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		  <td><INPUT type="text" class="search" dir="<%=dir_obj_var%>" id="contact_cellular" name="contact_cellular" value="<%=trim(vFix(contact_cellular))%>" onKeyPress="return submitenter(this,event)" style="width:100%"></td>
          <td><INPUT type="text" class="search" dir="<%=dir_obj_var%>" id="contact_name" name="contact_name" value="<%=trim(vFix(contact_name))%>" onKeyPress="return submitenter(this,event)" style="width:100%"></td>
		</tr>
		</tbody>
		<tr>						
			<td class="title_sort" align="center" width="5%">&nbsp;בחר&nbsp;</td>
			<td class="title_sort" align="center" width="15%">&nbsp;תאריך עדכון&nbsp;</td>
			<td class="title_sort" align="center" width="15%">&nbsp;טלפון ישיר&nbsp;</td>
			<td class="title_sort" align="center" width="15%">&nbsp;טלפון נייד&nbsp;</td>			
			<td class="title_sort" align="center" nowrap width="50%">&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;</td>			
		</tr>
<%Dim PageSize : PageSize = 20
	  Dim startRow : startRow = ((Page - 1) * PageSize)
	  Dim endRow : endRow = startRow + PageSize
	  
	  sqlstr = "SELECT COUNT(contact_ID) FROM dbo.CONTACTS WHERE (ORGANIZATION_ID = " & cInt(OrgID) & ")" & _
	  where_company_id & where_contact_name	 & where_contact_cellular
	  'Response.Write sqlstr
	  'Response.End  
	  set rs_count = con.GetRecordSet(sqlstr)
	  recCount = cLng(rs_count(0))
	  set rs_count = Nothing
	  
	  sqlstr = "SELECT * FROM (SELECT contact_ID, contact_NAME, cellular, phone, date_update, " & _
	 " ROW_NUMBER() OVER (ORDER BY contact_NAME, date_update DESC) AS rownum " & _
	 "  FROM dbo.CONTACTS WHERE (ORGANIZATION_ID = " & cInt(OrgID) & ")" & _
	 where_company_id & where_contact_name & where_contact_cellular & _
	 "  ) TMP WHERE ((rownum >= " & startRow & ") AND (rownum < " & endRow & ")) ORDER BY rownum"	
	'Response.Write sqlstr
	'Response.End
	SET pr = con.GetRecordSet(sqlstr)    
	IF NOT pr.eof THEN 
		recnom = 1		
		WHILE NOT pr.EOF
			contactID = trim(pr(0))
			contactNAME = trim(pr(1))					
			contactCellular = trim(pr(2))	
			contactPhone = trim(pr(3)) 
			date_update = cDate(pr(4)) %>	
		<tr>			
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data" bgcolor="#E6E6E6"><input type="button" ID="btnSend" NAME="btnSend" value="בחר" onclick="return choose('<%=contactID%>', escape('<%=altFix(contactNAME)%>'));" class="but_site"></td>
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data" bgcolor="#E6E6E6"><A class="link_categ" onclick="return choose('<%=contactID%>', escape('<%=altFix(contactNAME)%>'));" href="javascript:void(0)">&nbsp;<%=day(date_update)%>/<%=month(date_update)%>/<%=mid(year(date_update),3,2)%>&nbsp;<%=FormatDateTime(date_update,4)%></a></td>						
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data" bgcolor="#E6E6E6"><A class="link_categ" onclick="return choose('<%=contactID%>', escape('<%=altFix(contactNAME)%>'));" href="javascript:void(0)">&nbsp;<%=contactPhone%>&nbsp;</a></td>
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data" bgcolor="#E6E6E6"><A class="link_categ" onclick="return choose('<%=contactID%>', escape('<%=altFix(contactNAME)%>'));" href="javascript:void(0)">&nbsp;<%=contactCellular%>&nbsp;</a></td>			
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data" bgcolor="#E6E6E6"><A class="link_categ" onclick="return choose('<%=contactID%>', escape('<%=altFix(contactNAME)%>'));" href="javascript:void(0)">&nbsp;<%=contactNAME%>&nbsp;</a></td>
		</tr>
	<%	pr.movenext
		recnom = recnom + 1	
		Wend	
		
	    NumberOfPages = Fix((recCount / PageSize)+0.9)
			
	    If NumberOfPages > 1 then%>
		<tr>
		<td width="100%" align="middle" colspan="5" nowrap dir="ltr" class="card">
			<table border="0" cellspacing="0" cellpadding="2">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRow") <> nil Then
	               numOfRow = Request.QueryString("numOfRow")
	           Else numOfRow = 1
	           End If
	         %>
	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center"><A class=pageCounter title="לדפים הקודמים" href="contacts_list.asp?all=<%=all%>&amp;contact_name=<%=vFix(contact_name)%>&amp;contact_cellular=<%=vFix(contact_cellular)%>&amp;page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>&CompanyId=<%=CompanyId%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td class="form" valign="top"><font color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle"><p style="PADDING-RIGHT: 2px; PADDING-LEFT: 2px; FONT-WEIGHT: bold; FONT-SIZE: 9pt; BACKGROUND: #74b3d5; CURSOR: default; COLOR: #ffffff" 
                 ><%=i+10*(numOfRow-1)%></p></td>
	                  <%else%>
	                     <td align="middle"><A class=pageCounter href="contacts_list.asp?all=<%=all%>&amp;contact_name=<%=vFix(contact_name)%>&amp;contact_cellular=<%=vFix(contact_cellular)%>&amp;page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>&CompanyId=<%=CompanyId%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>	            
					<td class="form" valign="top"><font color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center"><A class=pageCounter title="לדפים הבאים" href="contacts_list.asp?all=<%=all%>&amp;contact_name=<%=vFix(contact_name)%>&amp;contact_cellular=<%=vFix(contact_cellular)%>&amp;page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>&CompanyId=<%=CompanyId%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>					
<%End If
		pr.Close
		set pr=Nothing    %>	
	<tr><td class="card" align="center" colspan="5" style="color:#5F435A;font-weight:600" dir="ltr">נמצאו&nbsp;<%=recCount%>&nbsp;רשומות</td></tr>					
<%Else 'not pr.eof%>
	<tr><td align="center" colspan="5" bgcolor="#E4E4E4" style="font-size:12px;font-weight:600;color:red;">לא נמצאו רשומות</td></tr>	
<%End If 'not pr.eof %>
<!-- end code -->
</form></table>
</td></tr></table>
</td></tr></table>
</BODY>
</HTML>
<%set con =nothing%>
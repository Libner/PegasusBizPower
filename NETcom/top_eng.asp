<%
    dim titles_array(6)
    If trim(COMPANIES) = "1" Then
	titles_array(0) = trim(Request.Cookies("bizpegasus")("title1"))
	End If
	If trim(SURVEYS) = "1" Then
	titles_array(1) = trim(Request.Cookies("bizpegasus")("title2"))
	End If
	If trim(EMAILS) = "1" Then
	titles_array(2) = trim(Request.Cookies("bizpegasus")("title3"))
	End If
	If trim(WORK_PRICING) = "1" Then
	titles_array(3) = trim(Request.Cookies("bizpegasus")("title4"))
	End If
	If trim(chief) = "1" Then
	titles_array(4) = trim(Request.Cookies("bizpegasus")("title5"))
	End If
	If trim(TASKS) = "1" Then
	titles_array(5) = trim(Request.Cookies("bizpegasus")("title6"))
	End If
	titles_array(6) = "התקנה"
	
	dim links_array(6)
	If trim(COMPANIES) = "1" Then
	links_array(0) = "members/companies/companies.asp"	
	'links_array(5) = "members/tasks/default.asp?reciver_id=" & userID
	end if
	If trim(SURVEYS) = "1" Then
	links_array(1) = "members/products/questions.asp"
	End If
	If trim(EMAILS) = "1" Then
	links_array(2) = "members/pages/default.asp"
	End If
	If trim(WORK_PRICING) = "1" Then
	links_array(3) = "members/projects/workers.asp"	
	ElseIf trim(CASH_FLOW) = "1" Then
	links_array(3) = "members/projects/movements.asp?type_id=1"
	End If
	If trim(chief) = "1" Then
	links_array(4) = "members/companies/company_types.asp"
	End If
	If trim(TASKS) = "1" Then
	links_array(5) = "members/tasks/default.asp?T=IN"
	End If
	links_array(6) = "members/projects/download.asp"	
	
%>
<table cellpadding=0 cellspacing=0 width=100% border=0  dir=<%=dir_var%>>
<tr><td <%If lang_id = 1 Then%> align=right <%Else%> align=left <%End If%> bgcolor="#FFFFFF">
<table border="0" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" <%If lang_id = 1 Then%> dir=rtl <%Else%> dir=ltr <%End If%>>
<tr>          
    <% for i=0 to Ubound(titles_array)
       If Len(titles_array(i)) > 0 Then
    %>    
	<td  align="center" nowrap valign="top" >					  		
	<A class="link_top<%If trim(numOftab) = trim(i) Then%>_over<%End if%>" href='<%=links_array(i)%>'  target=_self><%=titles_array(i)%></A>	
	</td>
	<td width=5 nowrap></td>	
	<%
	   End If
	   Next%>
</tr>	
</table></td></tr> 
<tr><td align=right bgcolor="#FFFFFF" height=3 nowrap></td></tr>
</table>    
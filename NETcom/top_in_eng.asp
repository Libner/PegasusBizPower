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
	titles_array(6) = "�����"
	
	dim links_array(6)	
	If trim(COMPANIES) = "1" Then
	links_array(0) = "../companies/companies.asp"
	end if	
	If trim(TASKS) = "1" Then
	links_array(5) = "../tasks/default.asp?T=IN"
	End If
	If trim(SURVEYS) = "1" Then
	links_array(1) = "../products/questions.asp"
	End If
	If trim(EMAILS) = "1" Then
	links_array(2) = "../pages/default.asp"
	End If	
	If trim(WORK_PRICING) = "1" Then
	links_array(3) = "../projects/workers.asp"	
	ElseIf trim(CASH_FLOW) = "1" Then
	links_array(3) = "../projects/movements.asp?type_id=1"
	End If
	If trim(chief) = "1" Then
		If trim(COMPANIES) = "1" Then
		links_array(4) = "../companies/company_types.asp"
		ElseIf trim(TASKS) = "1" Then
		links_array(4) = "../companies/activity_types.asp"
		End If	
	End If
	links_array(6) = "../projects/download.asp"	
	
	dim in_links_array(6,8)
	If trim(COMPANIES) = "1" Then
	in_links_array(0,0) = "../companies/companies.asp"
	in_links_array(0,1) = "../companies/contacts.asp"
	in_links_array(0,2) = "../companies/companies_private.asp"	
	in_links_array(0,3) = "../projects/default.asp"	
	in_links_array(0,4) = "../projects/default_action.asp"	  'clali
	in_links_array(0,5) = "../companies/reports.asp"	  '�����
	end if	
	
	in_links_array(1,0) = "../products/questions.asp"		
	in_links_array(1,1) = "../appeals/appeals.asp"
	in_links_array(1,2) = "../appeals/appeal.asp"
	in_links_array(1,3) = "../reports/reports.asp?type=cl"	
	
	in_links_array(2,0) = "../pages/default.asp"
	in_links_array(2,1) = "../groups_clients/default.asp"		
	in_links_array(2,2) = "../products/products.asp"
	in_links_array(2,3) = "../appeals/feedbacks.asp"
	in_links_array(2,4) = "../reports/reports.asp?type=ml"
	
	if trim(TASKS) = "1" Then
	in_links_array(5,0) = "../tasks/default.asp?T=IN"
	in_links_array(5,1) = "../tasks/default.asp?T=OUT"
	end if
	
	If trim(WORK_PRICING) = "1" Then
	in_links_array(3,0) = "../projects/workers.asp"
	in_links_array(3,1) = "../projects/projects_list.asp"
	in_links_array(3,2) = "../projects/prices_list.asp"			
	in_links_array(3,7) = "../projects/projects_reports.asp"
	End If
	If trim(CASH_FLOW) = "1" Then
	in_links_array(3,3) = "../projects/movements.asp?type_id=1"
	in_links_array(3,4) = "../projects/movements.asp?type_id=2"
	in_links_array(3,5) = "../projects/movements_activate.asp"
	in_links_array(3,6) = "../projects/cash_flow.asp"
	End If		
	
	If trim(COMPANIES) = "1" Then
	in_links_array(4,0) = "../companies/company_types.asp"	
	in_links_array(4,1) = "../companies/messangers.asp"
	in_links_array(4,2) = "../projects/addform.asp"
	End If
	
	If trim(TASKS) = "1" Then
	in_links_array(4,3) = "../companies/activity_types.asp"
	End If
	
	in_links_array(4,4) = "../workers/jobs.asp"
	in_links_array(4,5) = "../workers/default.asp"
	If trim(CASH_FLOW) = "1" Then
	in_links_array(4,6) = "../projects/banks.asp"
	in_links_array(4,7) = "../projects/movement_types.asp"
	End If	
	
	in_links_array(6,0) = "../projects/download.asp"	
	
	dim in_array(6,8)	
	If trim(COMPANIES) = "1" Then
	in_array(0,0) = trim(Request.Cookies("bizpegasus")("CompaniesMulti")) &  " �������� "
	in_array(0,1) = trim(Request.Cookies("bizpegasus")("ContactsMulti")) & " �������� "
	in_array(0,2) = trim(Request.Cookies("bizpegasus")("CompaniesMulti"))& " ������ "  
	in_array(0,3) = trim(Request.Cookies("bizpegasus")("ProjectMulti"))	
	in_array(0,4) = trim(Request.Cookies("bizpegasus")("ActivitiesMulti"))	
	in_array(0,5) = "����� ������"
	end if	
	
	if trim(TASKS) = "1" Then
	in_array(5,1) = trim(Request.Cookies("bizpegasus")("TasksMulti")) & " ������"
	in_array(5,0) = trim(Request.Cookies("bizpegasus")("TasksMulti")) & " ������"
	End If
	
	in_array(6,0) = "����� ����� Bizpower"
	
	in_array(1,0) = "����� �����"			
	in_array(1,1) = "����� �����"
	in_array(1,2) = "����� �����"
	in_array(1,3) = "����� �����"
		
	in_array(2,0) = "���� �������"
	in_array(2,1) = "������ �����"		
	in_array(2,2) = "�����"	
	in_array(2,3) = "������ ������"		
	in_array(2,4) = "����� ������"
			
	If trim(WORK_PRICING) = "1" Then
	in_array(3,0) = "���� ������"
	in_array(3,1) = "���� " & trim(Request.Cookies("bizpegasus")("ProjectMulti"))
	in_array(3,2) = "����� " & trim(Request.Cookies("bizpegasus")("ProjectMulti"))
	in_array(3,7) = "�����"	
	End If
	If trim(CASH_FLOW) = "1" Then	
	in_array(3,3) = "������ ��������"
	in_array(3,4) = "������ ��������"
	in_array(3,5) = "����� ��� �����"	
	in_array(3,6) = "����� �������"	
	End If	
	
	If trim(COMPANIES) = "1" Then
	in_array(4,0) = "������ " & trim(Request.Cookies("bizpegasus")("CompaniesMulti"))
	in_array(4,1) = "������ " &trim(Request.Cookies("bizpegasus")("ContactsMulti"))
	in_array(4,2) = "���� " &trim(Request.Cookies("bizpegasus")("Projectone"))
	End If
	If trim(TASKS) = "1" Then
	in_array(4,3) = "���� " & trim(Request.Cookies("bizpegasus")("TasksMulti"))
	End If
	in_array(4,4) = "���� ������ �������"
	in_array(4,5) = "������ �������"
	If trim(CASH_FLOW) = "1" Then	
	in_array(4,6) = "�����"
	in_array(4,7) = "���� ������"
	End If	
%>
<table cellpadding=0 cellspacing=0 width=100% border=0  dir=<%=dir_var%>>
<tr><td <%If lang_id = 1 Then%> align=right <%Else%> align=left <%End If%> bgcolor="#FFFFFF">
<table border="0" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" <%If lang_id = 1 Then%> dir=rtl <%Else%> dir=ltr <%End If%>>
<tr>          
    <% dim i, j
       for i=0 to Ubound(titles_array) 
       If Len(titles_array(i)) > 0 Then
    %>    
	<td  align="center" nowrap valign="top" dir=rtl>					  		
	<A class="link_top<%If trim(numOftab) = trim(i) Then%>_over<%End if%>" href='<%=links_array(i)%>'  target=_self><%=titles_array(i)%></A>	
	</td>
	<td width=5 nowrap></td>	
	<%
	   End If
	   Next%>
</tr>	
</table></td></tr>
<tr>
      <td width="100%" height="20" nowrap bgcolor="#FFFFFF"></td>
</tr> 
<tr>
	<td width="100%" align="right">
	<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 bgcolor="#FFFFFF">
		<TR>			
			<TD width=100% style="border-bottom: solid 1px #6F6D6B;">&nbsp;</TD>
			<%For j=Ubound(in_array,2) To 0 Step -1%>
			<%If trim(in_links_array(numOfTab,j)) <> "" Then%>			
			<TD width="2" bgcolor=#FFFFFF nowrap style="border-bottom: solid 1px #6F6D6B;">&nbsp;</TD>
			<TD bgcolor="#F0C000" nowrap dir=rtl><a HREF="<%=in_links_array(numOfTab,j)%>" target="_self" class="title_tab<%If trim(numOfLink) = trim(j) Then%>_over<%End If%>"><%=in_array(numOfTab,j)%></a></TD>
			
			<%End If%>
			<%Next%>			
			
		</TR>		
	</TABLE>
	</td>
</tr>
</table>    
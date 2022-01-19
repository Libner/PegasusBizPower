<%
	wizard_id=trim(Request.Cookies("bizpegasus")("wizardId"))
	If trim(wizard_id) <> "" And IsNumeric(wizard_id) Then
	count_pages = 1
	sqlstr = "Select Count(page_id) FROM wizards WHERE wizard_id = " & wizard_id
	set rs_count = con.getRecordSet(sqlstr)
	If not rs_count.eof Then
		count_pages = trim(rs_count(0))
	End If	
	set rs_count = Nothing
	'procent = Fix(100/count_pages)+1	
%>
<table border="0" width="100%" height="48" cellspacing="0" cellpadding="0" bgcolor=white align=right ID="Table3">
  <tr>
    <td width="100%" height="2"></td>
  </tr>
  <tr>
    <td width="100%" height="42" align=right></td>
    <td nowrap height="42" align=right>
    <table border="0" width="98%" height="42" cellspacing="0"  cellpadding="0">
      <tr>      
      <%For i=count_pages To 1 Step -1%>
      <%
		 sqlstr = "Select Step_File, Step_Title FROM Wizards WHERE Page_ID = " & i &_
		 " And Wizard_ID = " & wizard_id
		 set rs_wiz = con.getRecordSet(sqlstr)
		 If not rs_wiz.eof Then
			Step_Title = trim(rs_wiz(1))
			Step_File = trim(rs_wiz(0))
			is_done = false
			If wizard_id=1 Or wizard_id=2 Then
				If trim(wpageID) <> "" And trim(i) = 1 Then
					is_done = true
				ElseIf trim(wgroupID) <> "" And trim(i) = 2 Then
					is_done = true			
				End If
			Else
				'Response.Write 	trim(wWorkers) And trim(i) = 1
				If trim(wprodID) <> "" And trim(i) = 1 Then
					is_done = true
				ElseIf trim(wpageID) <> "" And trim(i) = 2 Then
					is_done = true
				ElseIf trim(wgroupID) <> "" And trim(i) = 3 Then
					is_done = true			
				ElseIf trim(wWorkers) = "true" And trim(i) = 1 Then
					is_done = true	
				ElseIf trim(wCompanies) = "true" And trim(i) = 2 Then
					is_done = true
				ElseIf trim(wProjects) = "true" And trim(i) = 3 Then
					is_done = true
				ElseIf trim(wPricing) = "true" And trim(i) = 1 Then
					is_done = true															
				End If			
			End If	
			
			If trim(wizard_page_id) = trim(i) Then
			     bg_color = "#FFD011"
			ElseIf is_done Then
				 bg_color = "#706EA6"    
			Else
				 bg_color = "#D3D3D3" 
			End If
      %>
        <td width=3 nowrap></td> 
        <td><img src="../../images/arrow_<%If trim(wizard_page_id) = trim(i) Then%>cur<%ElseIf is_done Then%>vis<%End if%>.gif"></td>        
        <td align="right" nowrap style="background-color='<%=bg_color%>'">
        <table border="0" cellspacing="0" cellpadding="0" ID="Table5">
          <tr>
            <td align="right" nowrap>&nbsp;<A class="wizard<%If trim(wizard_page_id) = trim(i) Then%>_ov<%ElseIf is_done Then%>_cl<%End If%>" href="../wizard/<%=Step_File%>" target=_self><%=Step_Title%></a>&nbsp;</td>
          </tr>         
        </table>
        <td width=30 nowrap style="background-color='<%=bg_color%>'" align=center><img src="../../images/<%=i%><%If is_done = false Or (trim(wizard_page_id) = trim(i)) Then%>_cl<%End If%>.gif" width="18" height="23" border="0"></td>
        </td>               
       <%
		  End If
		  set rs_wiz = Nothing
       %> 
       <%Next%>      
      </tr>         
   </table>
   </td></tr>
  <tr>
    <td width="100%" height="2"></td>
  </tr>
  </table>
  <%End If%>
  
   
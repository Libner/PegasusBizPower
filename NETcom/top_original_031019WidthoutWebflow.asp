<table cellpadding=0 cellspacing=0 width="100%" border=0>
<tr><td align=right bgcolor="#FFFFFF">
<table border="0" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" align="<%=align_var%>" dir=<%=dir_obj_var%>>
<% 
arr_bars_top = session("arr_bars_top")
links_array = session("links_array")

If Len(numOftab) = 0 Then
	numOftab = -1
End If

If IsArray(arr_bars_top) Then %>
<tr>        
	<td align="center" nowrap valign="top" >					  		
	<A class="link_top<%If trim(numOftab) = "-1" Then%>_over<%End if%>" href="<%=strLocal%>/netCom/" target=_self>ראשי</A>	
	</td>
	<td width=5 nowrap></td>	
    <%For i=0 To Ubound(arr_bars_top)-1
		arr_bar = Split(arr_bars_top(i),",")
		barID = arr_bar(0) 	
		barTitle = arr_bar(1)
		barVisible = arr_bar(2)
		If trim(barVisible) = "1" Then
			'href = "href='" & strLocal & links_array(i+1) &"' "
			href = "href='" & links_array(i+1) &"' "
			dis = ""	%>    
	<td align="center" nowrap valign="top" >					  		
	<A class="link_top<%If trim(numOftab+1) = trim(barID) Then%>_over<%End if%>" <%=dis%> <%=href%> target=_self><%=barTitle%></A>	
	</td>
	<td width=5 nowrap></td>	
	<%End If%>
	<%Next%>	
</tr>	
<%End If%>
</table></td></tr> 
<tr><td align=right bgcolor="#FFFFFF" height=3 nowrap></td></tr>
</table>    
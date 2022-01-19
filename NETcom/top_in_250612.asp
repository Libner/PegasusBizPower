<table cellpadding="0" cellspacing="0" width="100%" border="0">
<tr><td align="right" bgcolor="#FFFFFF">
<table border="0" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" align="<%=align_var%>" dir=<%=dir_obj_var%>>
<tr>          
<%If Len(numOftab) = 0 Then
		numOftab = -1
	 End If %>
	<td align="center" nowrap valign="top" ><A class="link_top<%If trim(numOfTab) = "-1" Then%>_over<%End if%>" href="<%=strLocal%>/netCom/" target=_self>ראשי</A></td>
	<td width="5" nowrap></td>	
<%arr_bars_top = session("arr_bars_top")
	  links_array = session("links_array")
	  If IsArray(arr_bars_top) Then    
      For i=0 To Ubound(arr_bars_top)-1
		arr_bar = Split(arr_bars_top(i),",")
		barID = trim(arr_bar(0))
		barTitle = trim(arr_bar(1))
		barVisible = arr_bar(2)	
		If trim(barVisible) = "1" Then
			href = "href='" & Application("VirDir") & "/" & trim(links_array(i+1)) &"' "
			dis = "" %>    
	<td  align="center" nowrap valign="top" dir="rtl"><a class="link_top<%If trim(barID) = trim(numOfTab+1) Then%>_over<%End if%>" 
	<%=href%> <%=dis%> target=_self><%=barTitle%></a></td>
	<td width="5" nowrap></td>	
	<%End If%>
	<%Next%>
	<%End If%>
</tr>	
</table></td></tr>
<tr>
      <td width="100%" height="20" nowrap bgcolor="#FFFFFF"></td>
</tr> 
<tr>
	<td width="100%">
	<TABLE width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF" align="<%=align_var%>" dir=<%=dir_obj_var%>>
		<TR>  <%
		        count = 0
		        arr_bars = session("arr_bars")
		        If IsArray(arr_bars) Then
		        For j=0 To Ubound(arr_bars)-1
				arr_bar = Split(arr_bars(j),",")
				barID = trim(arr_bar(0))
				barTitle = trim(arr_bar(1))
				barVisible = trim(arr_bar(2))
				barParent = trim(arr_bar(3))
				barURL = trim(arr_bar(4))
				barParentOrder = trim(arr_bar(5))
				
				If cInt(barParent) = cint(numOfTab)+1 Then
				If trim(barVisible) = "1" Then
					href = "href='" & Application("VirDir") & "/" & trim(barURL) &"' "
					dis = ""	%>    					
			<TD bgcolor="#F0C000" nowrap dir=rtl><a <%=href%> <%=dis%> target="_self" class="title_tab<%If trim(count) = trim(numOfLink) Then%>_over<%End If%>"><%=barTitle%></a></TD>		
			<TD width="2" bgcolor="#FFFFFF" nowrap style="border-bottom: solid 1px #6F6D6B;">&nbsp;</TD>
	<%			count = count + 1	
			End If
			End If%>
			<%Next%>
			<%End If%>
			<TD width="100%" style="border-bottom: solid 1px #6F6D6B;">&nbsp;</TD>						
		</TR>		
	</TABLE>
	</td>
</tr>
</table>    
<%	'if QUESTION_TYPE = "1" or QUESTION_TYPE = "2" then ' if question 	%>
  <tr>
	<td align="center">
	<TABLE align="center" height=340 BORDER=0 width="100%" CELLSPACING=0 CELLPADDING=2>				 
<%	
	sqlStr = "SELECT Field_Id, Field_Title, Field_Type, Field_Scale FROM FORM_FIELD Where product_id=" & prodId &_
	" And Field_Type IN (3,5,8,9,10,11,12) Order by Field_Order"
	set fields=con.GetRecordSet(sqlStr)
	If not fields.eof Then
		arr_fields = fields.getRows()
	End If
	set fields = Nothing
				
	If IsArray(arr_fields) Then
	For ff=0 To Ubound(arr_fields,2)
		Field_Id = trim(arr_fields(0,ff))
		Field_Title = trim(arr_fields(1,ff))
		Field_Type = trim(arr_fields(2,ff))
		Field_Scale = trim(arr_fields(3,ff))
		if Field_Type = "10" then '//nosee
		%>
		<tr><td height="5" nowrap></td></tr>
		<tr>
			<td dir="rtl" align="<%=align_var%>" class="title_table_admin" style="font-size:12pt;padding-right:5px;padding-left:5px"><%=Field_Title%></td>
		</tr>				
		<%			
		end if
		if Field_Type = "3" or Field_Type = "5" or Field_Type = "8" or Field_Type = "9" or Field_Type = "11"  or Field_Type = "12"  then				
		%>
		<tr>
			<td dir="rtl" align="<%=align_var%>" class="card5" style="font-size:11pt;padding-right:5px;padding-left:5px"><%=Field_Title%></td>
		</tr>			
		<tr>		
			<%Field_Title=""%>
			<td align=<%=td_align%>  valign=top>			
		    <%=GetFieldGraph(prodId,Field_Id,Field_Title,Field_Type,Field_Scale,pr_language,type_id)%>							
			</td>				
		</tr>
		<tr><td><hr size=1 color="Black"></td></tr>																							
		<%End If%>
<%	Next
End If%>				
	</TABLE>		
  </td>	
</tr>		
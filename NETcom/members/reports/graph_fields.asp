<tr><td width="640" height=5></td></tr>
	 <%if prodId=16504 then%>
	 	<tr>
			<td align="center" valign="top">
				<table width="640" border="0" cellspacing="0" ID="Table1" align=right>
					<tr><td height=5 nowrap></td></tr>
					<tr>
						<td dir="rtl" align="right" class="title_table_admin" style="font-size:12pt">&nbsp;<b>לאיזה יעד נסיעה הינך מתעניין?</b></td>
					</tr>
				</table>
			</td>
		</tr>
			<tr>					
			<td align="center" valign=top height=140>
				<table align="center" width="640" border=0 cellspacing=1 cellpadding=0 ID="Table2">
				 <%=GetDestinitionGraph()%>
 	</table>
			</td>	
		</tr>		
		<tr><td><hr size=1 color="Black"></td></tr>
 <%end if%>
<tr>
	<td align="center">
	<TABLE align="center" BORDER=0 width="640" CELLSPACING=1 CELLPADDING=1>				 
<%	
	sqlStr = "SELECT Field_Id,Field_Title,Field_Type,Field_Scale FROM Form_Field Where product_id=" & prodId &_
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
		
		If Field_Type = "10" Then '//nosee
		%>
		<tr>
			<td align="center" valign="top">
				<table width="640" border="0" cellspacing="0">
					<tr><td height=5 nowrap></td></tr>
					<tr>
						<td dir="rtl" align="<%=td_align%>" class="title_table_admin" style="font-size:12pt">&nbsp;<b><%=Field_Title%></b></td>
					</tr>
					<tr><td height=5 nowrap></td></tr>
				</table>
			</td>
		</tr>
		<%Else%>												
		<tr>					
			<td align="center" valign=top height=140>
				<table align="center" width="640" border=0 cellspacing=1 cellpadding=0>
				<%=GetFieldGraph(prodId,Field_Id,Field_Title,Field_Type,Field_Scale,pr_language,type_id)%>
				</table>
			</td>	
		</tr>		
		<tr><td><hr size=1 color="Black"></td></tr>
		<%End If%>
<%	Next
End If%>					
		</TABLE>		
	</td>	
</tr>
<%	
    sqlstr = "Select PRODUCT_ID FROM PRODUCTS WHERE PRODUCT_TYPE = 1 AND ORGANIZATION_ID=" & trim(OrgID) ' טופס של פרויקט
	set rs_pr = con.getRecordSet(sqlstr)
	If not rs_pr.eof Then
		prod_id = trim(rs_pr(0))
	End If
	set rs_pr = nothing
	
    set fields=con.GetRecordSet("SELECT Field_Id,Field_Title,Field_Type,Field_Size,Field_Align,Field_Scale,FIELD_MUST FROM FORM_FIELD Where product_id=" & prod_id &" Order by Field_Order")
	If not 	fields.EOF Then		
	do while not fields.EOF 
		Field_Id = fields(0)
		Field_Title = trim(fields(1))
		Field_Type = fields(2)
		Field_Size = fields(3)
		Field_Align = trim(fields(4))
		Field_Scale = fields(5)
		FIELD_MUST = fields(6)
		if FIELD_MUST then
			is_must_fields = true
		end if
	%>
		<%if trim(lang_id) = "2" then%>
			<tr><td width="100%" colspan=2 <%if trim(Field_Type) = "10" then%> class="title_form" <%Else%> class="form" <%End If%> bgcolor=#DADADA  style="padding-left:15px" align=left dir=ltr>&nbsp;
			<%if trim(Field_Type) = "5" then%>
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng","","" & Field_Title & "")%>
			<%end if%>
			<b><%=breaks(Field_Title)%>&nbsp;</b></td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="left" dir=ltr colspan=2 valign=middle width=100% class="form" bgcolor=#f0f0f0 style="padding-left:15px">					
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng","","" & Field_Title & "")%>						
				</td>
			</tr>
			<%end if%>
		<%else%>	
			<tr><td width="100%" colspan=2 <%if trim(Field_Type) = "10" then%> class="title_form" <%Else%> class="form" <%End If%> bgcolor=#DADADA  align=right valign="top" dir=ltr style="padding-right:15px;"><span dir="rtl"><b><%=breaks(Field_Title)%></b></span>
			<%if trim(Field_Type) = "5" then%><%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","","" & Field_Title & "")%><%end if%></td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="right" colspan=2 valign=middle width=100% class="form" bgcolor=#f0f0f0 style="padding-right:15px;">					
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","","" & Field_Title & "")%>					
				</td>
			</tr>
			<%end if%>
		<%end if%>
			
		
<%		fields.moveNext()
		loop
	set fields=nothing 
%>	
</table></td></tr>	
<%end if%>

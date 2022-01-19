<%	'if QUESTION_TYPE = "1" or QUESTION_TYPE = "2" then ' if question 	
		sqlstr = "EXECUTE get_project_field_value '" & OrgID & "','','"& project_ID & "',''"
		'Response.Write sqlStr
		set fields=con.GetRecordSet(sqlStr)
		do while not fields.EOF
			Field_Value = trim(fields("Field_Value"))			
			Field_Id = fields("Field_Id")
			Field_Title = trim(fields("Field_Title"))
			Field_Type = fields("Field_Type")			
			Field_Align = trim(fields("Field_Align"))
			Field_Size = fields("Field_Size")	
			Field_Scale = fields("Field_Scale")				
			%>									
		<%if Langu = "eng" then%>
			<tr><td width="100%" colspan=2 <%if trim(Field_Type) = "10" then%> class="title_form" <%Else%> class="form" <%End If%> bgcolor=#DADADA  style="padding-left:15px" align=left >&nbsp;
			<%if trim(Field_Type) = "5" then%>
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng", cStr(Field_Value), cStr(Field_Title))%>
			<%end if%>
			<b><%=breaks(Field_Title)%>&nbsp;</b></td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="left" colspan=2 valign=middle width=100% class="form" bgcolor=#f0f0f0 style="padding-left:15px">					
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng", cStr(Field_Value), cStr(Field_Title))%>						
				</td>
			</tr>
			<%end if%>
		<%else%>	
			<tr><td width="100%" colspan=2 <%if trim(Field_Type) = "10" then%> class="title_form" <%Else%> class="form" <%End If%> bgcolor=#DADADA  align=right valign="top" dir=ltr style="padding-right:15px;"><span dir="rtl"><b><%=breaks(Field_Title)%></b></span>
			<%if trim(Field_Type) = "5" then%><%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb", cStr(Field_Value), cStr(Field_Title))%><%end if%></td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="right" colspan=2 valign=middle width=100% class="form" bgcolor=#f0f0f0 style="padding-right:15px;">					
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb", cStr(Field_Value), cStr(Field_Title))%>					
				</td>
			</tr>
			<%end if%>
		<%end if%>
		
	<%		fields.moveNext()
			loop
		set fields=nothing
		
'end if ' QUESTION_TYPE = "1" or%>
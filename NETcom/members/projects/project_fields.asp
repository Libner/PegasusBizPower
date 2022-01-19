<%	'if QUESTION_TYPE = "1" or QUESTION_TYPE = "2" then ' if question 	
		If projectID <> nil Then
		sqlstr = "EXECUTE get_project_field_value '" & OrgID & "','','"& projectID & "',''"
		Elseif pricing_id <> nil Then
		sqlstr = "EXECUTE get_project_field_value '" & OrgID & "','','"& pricing_id & "',''"
		End If
		set fields=con.GetRecordSet(sqlstr)
		do while not fields.EOF 
			Field_Value = trim(fields("Field_Value"))
			Field_Id = fields("Field_Id")
			Field_Size = fields("Field_Size")
			If IsNumeric(Field_Size) Then
				Field_Size = cInt(Field_Size)
			Else
				Field_Size = 25
			End If	
			Field_Title = trim(breaks(fields("Field_Title")))
			Field_Type = fields("Field_Type")			
			Field_Align = trim(fields("Field_Align"))			
%>
		<tr>					
		<td align="<%=align_var%>" dir="<%=dir_obj_var%>" width="100%">		
		<%if Field_Type = "1" then  'text%>
		<span style="width:250;" class="Form_R" dir="<%=dir_obj_var%>">
		<%=Field_Value%>
		</span>	
		<%elseif Field_Type = "2" then  'textarea%>
		<span style="width:250;" class="Form_R" dir="<%=dir_obj_var%>">
		<%=breaks(trim(Field_Value))%>
		</span>	
		<%elseif Field_Type = "3" then  'select%>
		<span style="width:250;" class="Form_R" dir="<%=dir_obj_var%>">
		<%=Field_Value%>
		</span>	
		<%elseif Field_Type = "4" then  'scale with degree of importance%>
		<span style="width:250;" class="Form_R" dir="<%=dir_obj_var%>">
		<%=Field_Value%>
		</span>		
		<%elseif Field_Type = "5" then  'check%>
		<span style="width:250;" class="Form_R" dir="<%=dir_obj_var%>">	
		
		<%if Field_Value="on" and Field_Align="rtl" then%>пл<%elseif Field_Value="on" and Field_Align="ltr" then%>yes<%elseif Field_Align="rtl" then%>ам<%else%>no<%end if%>
		
		</span>	
		<%elseif Field_Type = "6" then  'date%>
		<span style="width:80;" class="Form_R" dir="<%=dir_obj_var%>">
		<%=Field_Value%>
		</span>	
		<%elseif Field_Type = "7" then  'number%>
		<span style="width:250;" class="Form_R" dir="<%=dir_obj_var%>">
		<%=Field_Value%>
		</span>	
		<%elseif Field_Type = "8" then  'radio%>
		<span style="width:250;" class="Form_R" dir="<%=dir_obj_var%>">
		<%=Field_Value%>
		</span>	
		<%elseif Field_Type = "9" then  'scale%>
		<span style="width:250;" class="Form_R" dir="<%=dir_obj_var%>">
		<%=Field_Value%>
		</span>	
		<%elseif Field_Type = "11" then  'scale%>
		<span style="width:250;" class="Form_R" dir="<%=dir_obj_var%>">
		<%=Field_Value%>
		</span>	
		<%elseif Field_Type = "12" then  'scale%>
		<span style="width:250;" class="Form_R" dir="<%=dir_obj_var%>">
		<%=Field_Value%>
		</span>	
		<%end if%>
		</td>
		<td dir="<%=dir_obj_var%>" align="<%=align_var%>" valign="top" width=130 nowrap><%=breaks(fields("Field_Title"))%></td>
		</tr>	
	<%		fields.moveNext()
			loop
		set fields=nothing
		
'end if ' QUESTION_TYPE = "1" or%>
<%	'if QUESTION_TYPE = "1" or QUESTION_TYPE = "2" then ' if question 	
	if quest_id=16735	then
		sqlstr = "EXECUTE get_field_value16735 '" & OrgID & "','','"& appId & "','" & quest_id & "',''"		
	else
		sqlstr = "EXECUTE get_field_value '" & OrgID & "','','"& appId & "','" & quest_id & "',''"		
	end if
	
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
			FIELD_MUST = fields("FIELD_MUST")
			if FIELD_MUST then
				is_must_fields = true
			end if	
				if quest_id=16735	then
			isTravelers=trim(fields("isTravelers"))
		else
			isTravelers=0
		end if
		
		
			
			%>
						
			<%if pr_language = "eng" then%>
			<tr><td width="100%" <%if trim(Field_Type) = "10" then%> colspan=2 class="title_form" <%Else%> class="form" <%End If%> bgcolor=#DADADA align=left dir=ltr style="padding-left:10px">
			<%if trim(Field_Type) = "5" then%>
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng", vFix(Field_Value), cStr(Field_Title))%>
			<%end if%>
			<b><%if FIELD_MUST then%><font color=red>*&nbsp;</font><%end if%><%=breaks(Field_Title)%></b>&nbsp;</td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="left" dir=ltr valign=middle width=100% class="form" bgcolor=#f0f0f0 style="padding-left:10px">					
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng", vFix(Field_Value), cStr(Field_Title))%>
				</td>
			</tr>
			<%end if%>
		<%else%>			
			<tr><td width="100%" <%if trim(Field_Type) = "10" then%> colspan=2 class="title_form" <%Else%> class="form" <%End If%> bgcolor=#DADADA  align=right valign="top" dir=ltr style="padding-right:15px"><span dir="rtl"><b><%if FIELD_MUST then%><font color=red>&nbsp;*&nbsp;</font><%end if%><%=breaks(Field_Title)%></b></span>
			<%if trim(Field_Type) = "5" then%><%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb", vFix(Field_Value), cStr(Field_Title))%><%end if%></td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="right" valign=middle width=100% class="form" bgcolor=#f0f0f0 style="padding-right:15px">	
				<%if (Field_Id=40660    and isTravelers>0) or Field_Id=40623 then '%>
				<%=con.GetFormFieldNoEdit(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb", vFix(Field_Value), cStr(Field_Title))%>
				<%else%>
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb", vFix(Field_Value), cStr(Field_Title))%>
				<%end if%>
				</td>
			</tr>
			<%end if%>
		<%end if%>
						
	<%		fields.moveNext()
			loop
		set fields=nothing
		
'end if ' QUESTION_TYPE = "1" or%>
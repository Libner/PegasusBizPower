<%	'if QUESTION_TYPE = "1" or QUESTION_TYPE = "2" then ' if question 	
		set fields=con.GetRecordSet("SELECT Field_Id,Field_Title,Field_Type,Field_Size,Field_Align,Field_Scale,FIELD_MUST FROM FORM_FIELD WHERE PRODUCT_ID=" & quest_id &" And ORGANIZATION_ID = "& OrgID &" Order by Field_Order")
		dim is_must_fields		
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
		<%if pr_language = "eng" then%>
			<tr><td width="100%" <%if trim(Field_Type) = "10" then%> colspan=2 class="title_form" <%Else%> class="form" <%End If%> bgcolor=#DADADA dir=ltr  style="padding-left:15px" align=left >
			<%if trim(Field_Type) = "5" then%>
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng","","" & Field_Title & "")%>
			<%end if%>
			<b><%if FIELD_MUST then%><font color=red>*&nbsp;</font><%end if%><%=breaks(Field_Title)%></b>&nbsp;</td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="left" dir=ltr valign=middle width=100% class="form" bgcolor=#f0f0f0 style="padding-left:10px">					
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng","","" & Field_Title & "")%>						
				</td>
			</tr>
			<%end if%>
		<%else%>	
			<tr><td width="100%" <%if trim(Field_Type) = "10" then%> colspan=2 class="title_form" <%Else%> class="form" <%End If%> bgcolor=#DADADA  align=right valign="top" dir=ltr style="padding-right:15px;"><span dir="rtl"><b><%if FIELD_MUST then%><font color=red>&nbsp;*&nbsp;</font><%end if%><%=breaks(Field_Title)%></b></span>
			<%if trim(Field_Type) = "5" then%><%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","","" & Field_Title & "")%><%end if%></td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="right" valign=middle width=100% class="form" bgcolor=#f0f0f0 style="padding-right:15px;" dir="<%=dir_var%>">					
					<%if Field_Id=40660 and isTravelers>0 then%>
						<%=con.GetFormFieldNoEdit(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","0","" & Field_Title & "")%>
			
					<%else%>
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","","" & Field_Title & "")%>					
			<%end if%>
				</td>
			</tr>
			<%end if%>
		<%end if%>
<%		fields.moveNext()
		loop
	set fields=nothing 
	
'end if ' QUESTION_TYPE = "1" or%>
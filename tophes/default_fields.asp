<%dim is_must_fields
  If IsNumeric(appID) And isNumeric(OrgID) And IsNumeric(quest_id) Then
	set fields = con.GetRecordSet("Exec dbo.get_field_value '" & OrgID & "','','" & appId & "','" & quest_id & "',''")
	do while not fields.EOF 
		Field_Id = trim(fields("Field_Id"))
		Field_Title = trim(fields("Field_Title"))
		Field_Type = trim(fields("Field_Type"))
		Field_Size = trim(fields("Field_Size"))
		Field_Align = trim(fields("Field_Align"))
		Field_Scale = trim(fields("Field_Scale"))
		FIELD_MUST = trim(fields("FIELD_MUST"))
		Field_Value = trim(fields("Field_Value"))
		if FIELD_MUST then
			is_must_fields = true
		end if	
		
		If pr_language = "eng" Then%>
			<tr><td width="100%" <%if trim(Field_Type) = "10" then%> colspan=2 class="form_header" <%Else%> class="Field_Title" <%End If%> bgcolor=#EBEBEB  style="padding-left:10px" align=left >
			<%if trim(Field_Type) = "5" then%>
			<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng", vFix(Field_Value), cStr(Field_Title))%>
			<%end if%>
			<b><%if FIELD_MUST then%><font color=red>*&nbsp;</font><%end if%><%=breaks(Field_Title)%></b>&nbsp;</td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="left" valign=middle width=100% class="Field_Title" bgcolor=#FFFFFF style="padding-left:10px">					
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng", vFix(Field_Value), cStr(Field_Title))%>
				</td>
			</tr>
			<%end if%>
		<%else%>	
			<tr><td width="100%" <%if trim(Field_Type) = "10" then%> colspan=2 class="form_header" <%Else%> class="Field_Title" <%End If%> bgcolor=#EBEBEB  align=right valign="top" dir=ltr style="padding-right:15px;"><span dir="rtl"><b><%if FIELD_MUST then%><font color=red>&nbsp;*&nbsp;</font><%end if%><%=breaks(Field_Title)%></b></span>
			<%if trim(Field_Type) = "5" then%><%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb", vFix(Field_Value), cStr(Field_Title))%><%end if%></td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="right" valign=middle width=100% class="Field_Title" bgcolor=#FFFFFF style="padding-right:15px;">					
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb", vFix(Field_Value), cStr(Field_Title))%>
				</td>
			</tr>
			<%end if%>
		<%end if%>
<%		fields.moveNext()
		loop
	set fields=nothing 
    ElseIf IsNumeric(quest_id) Then
		set fields = con.GetRecordSet("SELECT Field_Id,Field_Title,Field_Type,Field_Size,Field_Align," & _
		" Field_Scale,FIELD_MUST FROM FORM_FIELD Where product_id=" & cStr(cInt(quest_id)) & " Order by Field_Order")
	    do while not fields.EOF		
	    Field_Id = trim(fields("Field_Id"))
		Field_Title = trim(fields("Field_Title"))
		Field_Type = trim(fields("Field_Type"))
		Field_Size = trim(fields("Field_Size"))
		Field_Align = trim(fields("Field_Align"))
		Field_Scale = trim(fields("Field_Scale"))
		FIELD_MUST = trim(fields("FIELD_MUST"))
		if FIELD_MUST then
			is_must_fields = true
		end if			

		If pr_language = "eng" Then%>
			<tr><td width="100%" <%if trim(Field_Type) = "10" then%> colspan=2 class="form_header" <%Else%> class="Field_Title" <%End If%> bgcolor=#EBEBEB  style="padding-left:10px" align=left >
			<%if trim(Field_Type) = "5" then%>
			<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng","","" & Field_Title & "")%>
			<%end if%>
			<b><%if FIELD_MUST then%><font color=red>*&nbsp;</font><%end if%><%=breaks(Field_Title)%></b>&nbsp;</td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="left" valign=middle width=100% class="Field_Title" bgcolor=#FFFFFF style="padding-left:10px">					
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng","","" & Field_Title & "")%>						
				</td>
			</tr>
			<%end if%>
		<%else%>	
			<tr><td width="100%" <%if trim(Field_Type) = "10" then%> colspan=2 class="form_header" <%Else%> class="Field_Title" <%End If%> bgcolor=#EBEBEB  align=right valign="top" dir=ltr style="padding-right:15px;"><span dir="rtl"><b><%if FIELD_MUST then%><font color=red>&nbsp;*&nbsp;</font><%end if%><%=breaks(Field_Title)%></b></span>
			<%if trim(Field_Type) = "5" then%><%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","","" & Field_Title & "")%><%end if%></td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="right" valign=middle width=100% class="Field_Title" bgcolor=#FFFFFF style="padding-right:15px;">					
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","","" & Field_Title & "")%>					
				</td>
			</tr>
			<%end if%>
		<%end if%>
<%		fields.moveNext()
		loop
	set fields=nothing 
    End If%>
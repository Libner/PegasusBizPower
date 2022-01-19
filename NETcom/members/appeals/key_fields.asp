<%sqlstr = "EXECUTE get_field_value '" & OrgID & "','','"& appId & "','"& quest_id &"','1'"
	 set fields=con.GetRecordSet(sqlstr)
	 If not fields.EOF Then
		count = fields.recordCount
		i = 0
        do while not fields.EOF 
		i = i + 1
		Field_Value = trim(fields("Field_Value"))
		'Field_Id = fields("Field_Id")
		Field_Title = trim(breaks(fields("Field_Title")))
		Field_Type = fields("Field_Type")
		FIELD_EXCEPTION = fields("FIELD_EXCEPTION")
		Field_Align = trim(fields("Field_Align"))
		is_exception = false	%>
   <%if Field_Type <> "10" then  'question%>	
		<td align="<%=align_var%>" valign=top width=30%><a class="link_categ" HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" dir="<%=dir_obj_var%>">&nbsp;			
		<%if Field_Type = "1" then  'text%>
			<%=Field_Value%>
		<%elseif Field_Type = "2" then  'textarea%>
			<%=breaks(Field_Value)%>
		<%elseif Field_Type = "3" then  'select%>
			<%=Field_Value%>
		<%elseif Field_Type = "4" then  'scale with degree of importance
			if Field_Value <> "" then
				tmparr = Split(Field_Value," ")
				strValue = tmparr(0)
				If UBound(tmparr) > 0 Then
					strImportance = tmparr(1)
				Else
					strImportance = "H"
				End If
				select case strImportance
				case "H","ג"
					pf_exception = IMPORTANCE_H
				case "M","ב"
					pf_exception = IMPORTANCE_M
				case "L","נ"
					pf_exception = IMPORTANCE_L
				case "NA","לי"
					pf_exception = IMPORTANCE_NA
				case else
					pf_exception = IMPORTANCE_H
				end select
				is_exception = get_exception(pf_exception,strValue)
			end if
			if IsNumeric(strValue) then%>
			<%if pr_language = "eng" then%>&nbsp;&nbsp;Importance&nbsp;&nbsp;<%=strImportance%>&nbsp;&nbsp;score&nbsp;&nbsp;<%end if%><%if is_exception=true then%><font color=red><%end if%><b><%=strValue%></b><%if is_exception=true then%></font><%end if%><%if pr_language = "heb" then%>&nbsp;&nbsp;ציון&nbsp;&nbsp;<b><%=strImportance%></b>&nbsp;&nbsp;חשיבות&nbsp;&nbsp;<%end if%>
		<%	end if
			elseif Field_Type = "5" then  'check%>
			<%if Field_Value="on" then%><%=Field_Title%><%else%>&nbsp;<%end if%>
		<%elseif Field_Type = "6" then  'date%>
			<%=Field_Value%>	
		<%elseif Field_Type = "7" then  'number%>
			<%=Field_Value%>
		<%elseif Field_Type = "8" then  'radio%>
			<%=Field_Value%>	
		<%elseif Field_Type = "9" then  'scale%>
			<%=Field_Value%>		
		<%elseif Field_Type = "11" then  'scale%>
			<%=Field_Value%>		
		<%elseif Field_Type = "12" then  'scale%>
			<%=Field_Value%>				
		<%end if%>&nbsp;
		</a></td>							
		<%End If%>
	<%		fields.moveNext()
			loop
		set fields=nothing
	End If	%>
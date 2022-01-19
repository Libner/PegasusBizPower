<%	    sql = "EXECUTE get_field_value '" & OrgID & "','','"& appId & "','"& quest_id &"',''"
        'Response.Write sql
        'Response.End
        set fields=con.GetRecordSet(sql)
        
        '==added by Mila 20/04/2020 - competitors fields for tofes mitanhen
        '== show only one field (if was in trip last year field 40847 or not - field 40848)
        isfield40847Shown=false
         isfield40848Shown=false
         
		do while not fields.EOF 
		'==added by Mila 20/04/2020 
			if quest_id=16504 then 
				isShowField=false
			else			
				isShowField=true
			end if
			
			Field_Value = trim(fields("Field_Value"))
			Field_AppId = fields("Field_Id")
			Field_Title = trim(breaks(fields("Field_Title")))
			Field_Type = fields("Field_Type")
			FIELD_EXCEPTION = fields("FIELD_EXCEPTION")
			Field_Align = trim(fields("Field_Align"))
			If Len(Field_Align) = 0 Then
				Field_Align = dir_align
			End If	
			is_exception = false
			
		'==added by Mila 20/04/2020 
		
			if quest_id=16504 then 
				if Field_AppId=40847 and Field_Value<>"" then
					if not isfield40848Shown then
						isfield40847Shown=true
						isShowField=true
					end if
				end if
				if Field_AppId=40848 and Field_Value<>"" then
					if not isfield40847Shown then
						isfield40848Shown=true
						isShowField=true
					end if
				end if
					if Field_AppId<>40848 and Field_AppId<>40847 then					
						isShowField=true
					end if
			end if
			if isShowField then%>
						<%if pr_language = "eng" then%>
							<tr><td dir=ltr <%if Field_Type <> "10" then  'question%>class="form"<%Else%>class="title_form"<%End If%> width="100%" align=left bgcolor=#DADADA valign="top" style="padding-left:15px"><b><%=breaks(fields("Field_Title"))%></b></td></tr>
						<%elseif pr_language = "heb" then%>	
							<tr><td dir=rtl <%if Field_Type <> "10" then  'question%>class="form"<%Else%>class="title_form"<%End If%> width="100%" align=right bgcolor=#DADADA valign="top" style="padding-right:15px"><b><%=breaks(fields("Field_Title"))%></b></td></tr>						
						<%end if%>
						<%if Field_Type <> "10" then  'question%>
							<tr>
							<td width="85%" class="form" align=<%=td_align%> dir="<%=Field_Align%>" bgcolor=#f0f0f0 <%if pr_language = "eng" then%>style="padding-left:15px"<%Else%>style="padding-right:15px"<%End if%>>
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
								<%if Field_Value="on" and Field_Align="rtl" then%>כן<%elseif Field_Value="on" and Field_Align="ltr" then%>yes<%elseif Field_Align="rtl" then%>לא<%else%>no<%end if%>
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
							<%end if%>
							</td>
							</tr>
						<%End If%>
						
						<%End If%>
	<%		fields.moveNext()
			loop
		set fields=nothing %>
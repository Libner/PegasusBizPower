<%	'if QUESTION_TYPE = "1" or QUESTION_TYPE = "2" then ' if question 	
		str="SELECT * FROM Form_Field WHERE Product_ID="&prodId&" Order by Field_Order"
		set obj_value=con.GetRecordSet(str)
		rsCount = obj_value.RecordCount  
		ii=1    	
		DO WHILE not obj_value.eof
			Field_Id = obj_value("Field_Id")
			Product_ID = obj_value("Product_ID")
			Field_Title = trim(obj_value("Field_Title"))
			Field_Type = obj_value("Field_Type")
			Field_Size = obj_value("Field_Size")
			Field_Align = trim(obj_value("Field_Align"))
			Field_Scale = obj_value("Field_Scale")
			Field_Key = obj_value("Field_Key")
			FIELD_EXCEPTION = obj_value("FIELD_EXCEPTION")
			dbOrder=obj_value("Field_Order")
			
  %>	
   <tr>
		<td align=center bgcolor="#dddddd">&nbsp;<a href="addform.asp?Id=<%=Field_ID%>&prodId=<%=Product_ID%>&delField=1#link<%=Field_Id%>" ONCLICK="return CheckDelField('<%=Field_ID%>')"><IMG SRC="../../images/delete_icon.gif" BORDER=0 name=word9 title="מחיקת שדה"></a>&nbsp;</td>
		<td align=center bgcolor="#dddddd">&nbsp;<a href="<%If trim(Field_Type) <> "10" Then%>addfield.asp<%Else%>addquest.asp<%End If%>?Id=<%=Field_ID%>&prodId=<%=Product_ID%>"><IMG SRC="../../images/edit_icon.gif" BORDER=0 name=word10 title="עריכת שדה"></a>&nbsp;</td>		
		<td align=center bgcolor="#dddddd"><input type=checkbox <%If Field_Key = "True" Then%> checked <%End If%> OnClick="document.location.href='addform.asp?Id=<%=Field_ID%>&prodId=<%=Product_ID%>&editKey=1'" ></td>
		<!--td align=center bgcolor="#dddddd">&nbsp;</td>
		<td align=center bgcolor="#dddddd">&nbsp;</td-->
		<td align=<%=td_align%> width=100% bgcolor=#E5E5E5 nowrap dir="<%=dir_obj_var%>">
		<TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=2 bgcolor=white dir="<%=dir_obj_var%>">
		<%if pr_language = "eng" then%>
			<tr><td width="100%" <%if trim(Field_Type) = "10" then%> colspan=2 class="title_form" <%Else%> class="form" <%End If%> bgcolor=#DADADA  style="padding-left:10px" dir=ltr align=left >
			<%if trim(Field_Type) = "5" then%>
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng","","" & Field_Title & "")%>
			<%end if%>
			<b><%=breaks(Field_Title)%></b></td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="left" dir=ltr valign=middle width=100% class="form" bgcolor=#f0f0f0 style="padding-left:10px">					
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng","","" & Field_Title & "")%>						
				</td>
			</tr>
			<%end if%>
		<%else%>	
			<tr><td width="100%" <%if trim(Field_Type) = "10" then%> colspan=2 class="title_form" <%Else%> class="form" <%End If%> bgcolor=#DADADA  align=right valign="top" dir="<%=dir_var%>" style="padding-right:15px;"><span dir="rtl"><b><%=breaks(Field_Title)%></b></span>
			<%if trim(Field_Type) = "5" then%><%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","","" & Field_Title & "")%><%end if%></td></tr>
			<%if trim(Field_Type) <> "5" And trim(Field_Type) <> "10" then%>
			<tr>		
				<td align="right" valign=middle width=100% class="form" bgcolor=#f0f0f0 style="padding-right:15px;" dir="<%=dir_var%>">					
				<%=con.GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","","" & Field_Title & "")%>					
				</td>
			</tr>
			<%end if%>
		<%end if%>
		</TABLE>		
		</td>
   		<td bgcolor=#f0f0f0 align=center dir="ltr" valign=top>
			<!--table border="0" cellpadding="0" cellspacing="1" width="100%">
				<tr>
					<td valign="top" align="left" width=50%>
					<%if ii>1 then%><a href="addform.asp?up=1&place=<%=dbOrder%>&Id=<%=Field_Id%>&prodId=<%=prodId%>#link<%=Field_Id%>"><img src="../../images/up.gif" border="0"><%else%><%end if%>
					</td>
					<td valign="bottom" align=<%=td_align%> width=50%>
					<%if ii<rsCount then%><a href="addform.asp?down=1&place=<%=dbOrder%>&Id=<%=Field_Id%>&prodId=<%=prodId%>#link<%=Field_Id%>"><img src="../../images/down.gif" border="0"><%else%><%end if%>
					</td>
				</tr>
			</table-->
			  <%if  ii = 1 then 'FIRST%>
			  <table border=0 cellspacing=0 cellpadding=0 width="100%">					    
			     <tr>
			        <td width="20%" nowrap align="right"><font color="#060165"><B>&nbsp;<%=ii%>&nbsp;</b></font></td>
			        <td width="15%" nowrap>&nbsp;</td>
			        <td width="15%" nowrap align="right"><a href="addform.asp?down=1&Id=<%=Field_Id%>&place=<%=dbOrder%>&prodId=<%=prodId%>"><img src="../../images/down.gif"  border=0 alt="להעביר שורה אחת למתה"></a>
			        </td>
			        <td width="35%" nowrap align="center" valign="middle"><input type="text" name="toPlace-<%=ii%>" value="" onKeyPress="return getNumbers(this)" class=Form style="width:100%""></td>		    			
			        <td width="15%" nowrap align="center" valign="middle"><a href="" onclick="return moveTo('<%=ii%>','<%=Field_Id%>','<%=dbOrder%>')"><img src="../../images/arrow_top_bot.gif" align="top" border=0 alt="להעביר למקום הצויין"></a></td>
			     </tr>
			  </table>
			<%elseif ii = count_ then'LAST%>
			  <table border=0 cellspacing=0 cellpadding=0 width="100%">
			     <tr>			        
			        <td width="20%" nowrap align="right"><font color="#060165"><B>&nbsp;<%=ii%>&nbsp;</b></font></td>
			        <td width="15%" nowrap align="right"><a href="addform.asp?up=1&Id=<%=Field_Id%>&place=<%=dbOrder%>&prodId=<%=prodId%>"><img src="../../images/up.gif" border=0 alt="להעביר שורה אחת למעלה"></a>
			        </td>
			        <td width="15%" nowrap>&nbsp;</td>
			        <td width="35%" nowrap align="center" valign="middle"><input type="text" size="3" name="toPlace-<%=ii%>" value="" onKeyPress="return getNumbers(this)" class=Form style="width:100%""></td>		    			
			        <td width="15%" nowrap align="center" valign="middle"><a href="" onclick="return moveTo('<%=ii%>','<%=Field_Id%>','<%=dbOrder%>')"><img src="../../images/arrow_top_bot.gif" align="top" border=0 alt="להעביר למקום הצויין"></a></td>
			     </tr>					    
			  </table>
			<%else 'OTHERS%>
			  <table border=0 cellspacing=0 cellpadding=0 width="100%">
			     <tr>				        
			        <td width="20%" nowrap align="right"><font color="#060165"><B>&nbsp;<%=ii%>&nbsp;</b></font></td>
			        <td width="15%" nowrap align="center"><a href="addform.asp?up=1&Id=<%=Field_Id%>&place=<%=dbOrder%>&prodId=<%=prodId%>"><img src="../../images/up.gif" border=0 alt="להעביר שורה אחת למעלה"></a></td>					     
			        <td width="15%" nowrap align="center"><a href="addform.asp?down=1&Id=<%=Field_Id%>&place=<%=dbOrder%>&prodId=<%=prodId%>"><img src="../../images/down.gif" border=0 alt="להעביר שורה אחת למתה"></a></td>
			        <td width="35%" nowrap align="center" valign="middle"><input type="text" size="3" name="toPlace-<%=ii%>" value="" onKeyPress="return getNumbers(this)" class=Form style="width:100%""></td>		    			
			        <td width="15%" nowrap align="center" valign="middle"><a href="" onclick="return moveTo('<%=ii%>','<%=Field_Id%>','<%=dbOrder%>')"><img src="../../images/arrow_top_bot.gif" align="top" border=0 alt="להעביר למקום הצויין"></a></td>
			     </tr>
			  </table>
			<%end if%>			
		</td>
   </tr>    
   
<%    obj_value.MoveNext
	  ii=ii+1
	  loop
	  set obj_value = nothing
	'end if 'QUESTION_TYPE = "1" 	  %>
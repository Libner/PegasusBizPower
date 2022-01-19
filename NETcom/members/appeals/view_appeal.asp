
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<% 
	appid = Request.QueryString("appid")
	quest_id = Request.QueryString("quest_id")	
    arr_Status = session("arr_Status")
	
	If IsNumeric(appid) Then
	sqlstr = "EXECUTE get_appeals '','','','','" & OrgID & "','','','','','','','" & appid & "','',''"
'	Response.Write("sqlstr=" & sqlstr)
	set app = con.GetRecordSet(sqlstr)
	if not app.eof then	
	appeal_date = FormatDateTime(app("appeal_date"), 2) & " " & FormatDateTime(app("appeal_date"), 4)	
	companyName = app("company_Name")
	contactName = app("contact_Name")
	projectName = app("project_Name")
	userName = app("user_Name")
	productName = app("product_Name")
	quest_id = app("questions_id")
	companyId = app("company_id")
	contactId = app("contact_id")
	projectId = app("project_id")
	appeal_status = app("appeal_status")
	appeal_status_name = trim(app("appeal_status_name"))
	appeal_status_color = trim(app("appeal_status_color"))	
	private_flag = trim(app("private_flag"))
	appeal_close_date = trim(app("appeal_close_date"))
	appeal_close_text = trim(app("appeal_close_text"))
	response_user_id = trim(app("response_user_id"))
	If trim(response_user_id) <> "" Then
			sqlstr = "Select FIRSTNAME + char(32) + LASTNAME From USERS Where USER_ID = " & response_user_id
			set rswrk = con.getRecordSet(sqlstr)
			If not rswrk.eof Then
			  response_user_name = trim(rswrk(0))
			End If
			set rswrk = Nothing	
	End If		
	sqlstr =  "Select Langu, PRODUCT_DESCRIPTION, RESPONSIBLE, FILE_ATTACHMENT, ATTACHMENT_TITLE From Products WHERE PRODUCT_ID = " & quest_id
	'Response.Write sqlstr
	'Response.End
	set rsq = con.getRecordSet(sqlstr)
	If not rsq.eof Then		
		Langu = trim(rsq(0))
		PRODUCT_DESCRIPTION = trim(rsq(1))
		RESPONSIBLE = trim(rsq(2))
		attachment_file = trim(rsq(3))
		attachment_title = trim(rsq(4))
	End If
	set rsq = Nothing	
	if Langu = "eng" then
		td_align = "left"
		pr_language = "eng"
	else
		td_align = "right"
		pr_language = "heb"
	end if
	set app = nothing
		If trim(appeal_status) = "1"  And trim(UserID) = trim(RESPONSIBLE) Then
			sqlstring="UPDATE appeals set appeal_status = '2' WHERE ORGANIZATION_ID = " & OrgID & " AND appeal_id = " & appid 	
			'Response.Write(sqlstring)
			'Response.End 
			con.ExecuteQuery(sqlstring) 
			appeal_status = "2" : status_name  = "בטיפול" : status_num = 2
		End If		
	End If
	Else
		appeal_date = FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4)
	End If
	
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 24 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing  				
%>
<html>
<head>
<title><%=productName%></title>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body style="margin:0px" onload="window.focus();javascript:window.print();">
<table border="0" cellpadding="0" cellspacing="0" width="100%" align=center dir="<%=dir_var%>">
<tr>
<td class="Field_Title" dir="<%=dir_obj_var%>" style="font-size:16pt;line-height:26px" align=center><b><%=productName%></b></td>
</tr>
	<%if trim(PRODUCT_DESCRIPTION) <> "" then%>
	<tr>
		<td align=<%=td_align%> width=100% bgcolor=#FFFFFF colspan=2 dir="<%=dir_var%>">
		<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=3 width=100%>
			<TR>
				<td class="form_makdim" dir="<%=dir_obj_var%>" width=100% align="center">
				<%=breaks(PRODUCT_DESCRIPTION)%>
				</td>
			</tr>
			<%If Len(attachment_title) > 0 Then%>
			<TR>
				<td class="form_makdim" dir="rtl" width=100% align=<%=td_align%>>
				<a class="file_link" href="<%=strLocal%>/download/products/<%=attachment_file%>" target=_blank><%=attachment_title%></a>
				</td>
			</tr>	
			<%End If%>			
		</TABLE>	
		</td>
	</tr>
	<%end if%>
	<tr><td height=10 colspan=2 nowrap></td></tr>				
	<tr>
	<td width="100%" dir="<%=dir_var%>" align="<%=align_var%>" valign=top>
		<TABLE WIDTH=100% BORDER=1 CELLSPACING=1 CELLPADDING=3 bgcolor=#FFFFFF style="border-collapse:collapse" bordercolor="#999999">	
		   <tr>
			 <td align="<%=align_var%>" width=100% class="Field_Title"><%=appID%></td>
			 <td align="<%=align_var%>" width=130 nowrap class="Field_Title"><b>&nbsp;ID&nbsp;</b></td>   
			</tr>
			<tr>
			 <td align="<%=align_var%>" width=100% class="Field_Title"><%=appeal_date%></td>
			 <td align="<%=align_var%>" width=130 nowrap class="Field_Title"><b>&nbsp;<!--תאריך הזנת הטופס--><%=arrTitles(9)%>&nbsp;</b></td>   
			</tr>
			<%If Len(userName) > 0 Then%>
			<tr>
			 <td align="<%=align_var%>" width=100% class="Field_Title"><%=userName%></td>
			 <td align="<%=align_var%>" width=130 nowrap class="Field_Title"><b>&nbsp;<!--הזנה ע"י עובד--><%=arrTitles(10)%>&nbsp;</b></td>   
			</tr>
			<%End If%>
			<tr>
			 <td align="<%=align_var%>" width=100% class="Field_Title">
			  <span class="status_num" style="background-color:<%=trim(appeal_status_color)%>; text-align:center"><%=appeal_status_name%></span>
			 </td>
			 <td align="<%=align_var%>" width=130 nowrap class="Field_Title"><b>&nbsp;<!--סטטוס--><%=arrTitles(11)%>&nbsp;</b></td>   
		    </tr>
		    <%If trim(appeal_status) = "3" Then%>
			<tr>
			 <td align="<%=align_var%>" width=100% class="Field_Title"><%=appeal_close_date%></td>
			 <td align="<%=align_var%>" width=130 nowrap class="Field_Title"><b>&nbsp;<!--תאריך סגירה--><%=arrTitles(26)%>&nbsp;</b></td>   
			</tr>	
			<tr>
			 <td align="<%=align_var%>" width=100% class="Field_Title"><%=appeal_close_text%></td>
			 <td align="<%=align_var%>" width=130 nowrap class="Field_Title"><b>&nbsp;<!--תוכן סגירה--><%=arrTitles(27)%>&nbsp;</b></td>   
			</tr>	
			<tr>
			 <td align="<%=align_var%>" width=100% class="Field_Title"><%=response_user_name%></td>
			 <td align="<%=align_var%>" width=130 nowrap class="Field_Title"><b>&nbsp;<!--ע''י עובד--><%=arrTitles(29)%>&nbsp;</b></td>   
			</tr>								    
		    <%End If%>		
			<%If isNumeric(companyId) And  trim(private_flag) = "0" Then%>
			<tr>
			 <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=100% class="Field_Title"><%=companyName%></td>
			 <td align="<%=align_var%>" width=130 nowrap class="Field_Title"><b>&nbsp;<!--קישור ל--><%=arrTitles(12)%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;</b></td>   
			</tr>
			<%End If%>	
			<%If isNumeric(contactId) Then%>
			<tr>
			 <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=100% class="Field_Title"><%=contactName%></td>
			 <td align="<%=align_var%>" width=130 nowrap class="Field_Title"><b>&nbsp;<!--קישור ל--><%=arrTitles(13)%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;</b></td>   
			</tr>
			<%End If%>			
			<%If isNumeric(projectID) And trim(projectName) <> "" Then%>
			<tr>
			 <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=100% class="Field_Title"><%=projectName%></td>
			 <td align="<%=align_var%>" width=130 nowrap class="Field_Title"><b>&nbsp;<!--קישור ל--><%=arrTitles(14)%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%>&nbsp;</b></td>   
			</tr>
			<%End If%>				 				
			</table></td></tr>	
			<tr><td height=10 colspan=2 nowrap></td></tr> 	
			<tr>
				<td width="100%" colspan="2" align="<%=align_var%>">
					<table width=100% border=1 style="border-collapse:collapse" cellspacing=0 cellpadding=3 bordercolor=#999999 >				
					<%	    
					sql = "EXECUTE get_field_value '" & OrgID & "','','"& appId & "','"& quest_id &"',''"
					'Response.Write sql
					'Response.End
					set fields=con.GetRecordSet(sql)
					do while not fields.EOF 
						Field_Value = trim(fields("Field_Value"))
						'Field_Id = fields("Field_Id")
						Field_Title = trim(breaks(fields("Field_Title")))
						Field_Type = fields("Field_Type")
						FIELD_EXCEPTION = fields("FIELD_EXCEPTION")
						Field_Align = trim(fields("Field_Align"))
						If Len(Field_Align) = 0 Then
							Field_Align = dir_align
						End If	
						is_exception = false
						%>
						<%if pr_language = "eng" then%>
							<tr><td dir=ltr <%if Field_Type <> "10" then  'question%>class="Field_Title"<%Else%>class="form_header"<%End If%> width="100%" align=left bgcolor=#EBEBEB valign="top" style="padding-left:5px"><b><%=breaks(fields("Field_Title"))%></b></td></tr>
						<%elseif pr_language = "heb" then%>	
							<tr><td dir=rtl <%if Field_Type <> "10" then  'question%>class="Field_Title"<%Else%>class="form_header"<%End If%> width="100%" align=right bgcolor=#EBEBEB valign="top" style="padding-right:5px"><b><%=breaks(fields("Field_Title"))%></b></td></tr>						
						<%end if%>
						<%if Field_Type <> "10" then  'question%>
							<tr>
							<td width="85%" class="Field_Title" align=<%=td_align%> dir="<%=Field_Align%>" bgcolor=#FFFFFF <%if pr_language = "eng" then%>style="padding-left:5px"<%Else%>style="padding-right:5px"<%End if%>>
							<%if Field_Type = "1" then  'text%>
								<%=Field_Value%>
							<%elseif Field_Type = "2" then  'textarea%>
								<%=breaks(Field_Value)%>
							<%elseif Field_Type = "3" then  'select%>
								<%=Field_Value%>
							<%elseif Field_Type = "5" then  'check%>
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
				<%		fields.moveNext()
						loop
					set fields=nothing %>
				</TABLE></td></tr>
			</table></td></tr>    
		<tr><td colspan=2 nowrap>&nbsp;</td></tr>
</table>
<%set con=nothing%>
</body>
</html>

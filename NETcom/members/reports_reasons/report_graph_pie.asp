<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
	  lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
 	  perSize = trim(Request.Cookies("bizpegasus")("perSize"))
	  is_groups = trim(Request.Cookies("bizpegasus")("ISGROUPS"))
	
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"  :  	align_var = "left"  :  dir_obj_var = "ltr"  :  	self_name = "Self"
	Else
		dir_var = "ltr"  :  	align_var = "right"  :  	dir_obj_var = "rtl"  :  self_name = "עצמי"
	End If
	 
	start_date = trim(Request("dateStart"))		
	end_date = trim(Request("dateEnd"))
	start_date_ =  Month(start_date) & "/" & Day(start_date) & "/" & Year(start_date)
    end_date_ =  Month(end_date) & "/" & Day(end_date) & "/" & Year(end_date) 
	prodId = trim(Request.Form("quest_id"))
	UserID = trim(Request("user_id"))	
    CompanyID = trim(Request("company_id"))
	ProjectID = trim(Request("project_id"))
	mechanismID = trim(Request("mechanism_id"))	
	type_reports=trim(Request.QueryString("type_reports"))
	
	If trim(type_reports) = "cl" Then
		type_id = 2
	Else
	    type_id = 1
	End If
	
	reasons=""
	'only for tofes mitan'en
	
			'הגורם ליצירת הטופס
			'---------------P2932 - type of form- REASONS ---------------
				'select creation reason-------------------------------
				'added by Mila 22/10/2019-----------------------------			
	if prodId<>"" then
		if cint(prodId)=16504 then
			reasons=Request.Form("reasons")
			reasons=replace(reasons," ","")
		end if
	end if
	'-----------------------------------------------------------
	appealsCount = 0
	sqlstr = "EXEC  dbo.get_appeals_count_reasons @OrgID='" & OrgID & "',  @start_date='" & start_date_ & "',  @end_date='" & end_date_ &"'," & _
	" @company_id='" & CompanyID &"', @contact_id='" & ContactId & "', @project_id='" & ProjectID & "', @product_id='" & prodId &"'," & _
	" @UserId='" & UserID & "', @IsGroups='" & is_groups & "',@reasons='"	& reasons & "'"
	'Response.Write sqlstr
	'Response.End 
	set rs_count = con.getRecordSet(sqlstr)
	if not rs_count.eof then
		appealsCount = cLng(rs_count(0))
	end if
	set rs_count = nothing	
	
	strTitle = ""
	If trim(CompanyID) <> "" Then
		sqlstr = "Select Company_Name From Companies WHERE Company_ID = " & CompanyID
		set rs_com = con.getRecordSet(sqlstr)
		if not rs_com.eof then
			strTitle = strTitle & "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & ":&nbsp;<font color=""#666699"">" & trim(rs_com(0)) & "</font>"
		end if
		set rs_com = nothing
	End If

	If trim(ProjectID) <> "" Then
		sqlstr = "Select Project_Name From Projects WHERE Project_ID = " & ProjectID
		set rs_project = con.getRecordSet(sqlstr)
		if not rs_project.eof then
			strTitle = strTitle & "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("Projectone")) & ":&nbsp;<font color=""#666699"">" & trim(rs_project(0)) & "</font>"
		end if
		set rs_project = nothing
	End If

	If trim(mechanismID) <> "" Then
		sqlstr = "Select mechanism_name, project_name From mechanism Inner Join Projects On mechanism.project_id = projects.project_id WHERE mechanism_id = " & mechanismID
		set rs_mech = con.getRecordSet(sqlstr)
		if not rs_mech.eof then
			If lang_id = 1 Then
				strTitle = strTitle & "<br>&nbsp;מנגנון:&nbsp;<font color=""#666699"">" & trim(rs_mech(1)) & " - " & trim(rs_mech(0)) & "</font>"
			Else
				strTitle = strTitle & "<br>&nbsp;Sub-Project:&nbsp;<font color=""#666699"">" & trim(rs_mech(1)) & " - " & trim(rs_mech(0)) & "</font>"
			End If			
		end if
		set rs_mech = nothing
	End If	
	
	Function GetFieldGraph(pProduct_Id, pField_Id, pField_Title, pField_Type, pField_Scale, pLanguage, pType_ID) 
		Dim col(9)
		col(0) = "#65A825"
		col(1) = "#19B4D0"
		col(2) = "#E87117"
		col(3) = "#6E15C3"
		col(4) = "#44B4AD"
		col(5) = "#011D66"
		col(6) = "#F84637"
		col(7) = "#1DF3FD"
		col(8) = "#EBB20D"
		colind = 0
		If pLanguage = "eng" Then
			td_align = "left"
		Else
			td_align = "right"
			string_sort = "DESC"
		End If
		Select Case pField_Type
		Case "1" ' one string text
			GetFieldGraph = ""
		Case "2" 'textarea
			GetFieldGraph = ""
		Case "3", "8", "11" 'select
			sqlstr = "Exec dbo.FIELD_STATISTICS_PROC_Reasons '" & pProduct_Id & "','" & pField_Id & "','" & pType_ID &_
			"','" & start_date & "','" & end_date & "','" & UserID & "','" & is_groups & "','" & CompanyID &_
			"','" & ProjectID & "','" & mechanismID & "','" & contactid & "','','"	& reasons & "'"
			Set rsSelect = con.GetRecordSet(sqlstr)
	        
			selAll = 0
			If Not rsSelect.EOF Then
				Do While Not rsSelect.EOF
					selAll = selAll + CInt(rsSelect("FIELD_QUANTITY"))
				rsSelect.MoveNext
				Loop
				rsSelect.MoveFirst
				Do While Not rsSelect.EOF
					If selAll <> 0 Then
						wd = CInt((CInt(rsSelect("FIELD_QUANTITY")) / selAll) * 100)
					Else
						wd = 0
					End If
					strDataSeries = strDataSeries & wd & ";;"
	                
				rsSelect.MoveNext
				colind = colind + 1
				Loop
				rsSelect.MoveFirst
				Do While Not rsSelect.EOF
						strDataLabels = strDataLabels & rsSelect("FIELD_SELECT") & ";;"
				rsSelect.MoveNext
				Loop
			End If  'not rsSelect.EOF                       
			Set rsSelect = Nothing
	        If Len(strDataSeries) > 0 Then
				GetFieldGraph = "<img src=""../../chart/chart.asp?Title="& Server.URLEncode(pField_Title) &"&strDataSeries=" & Server.URLEncode(strDataSeries)  & "&strDataLabels=" & Server.URLEncode(strDataLabels) & """>"
			Else
				GetFieldGraph = ""
			End If	
		Case "9", "12" 'radio
			sqlstr = "Exec dbo.QUESTION_VALUES_PROC_Reasons '" & pProduct_Id & "','" & pField_Id & "','" & pType_ID &_
			"','" & start_date & "','" & end_date & "','" & UserID & "','" & is_groups & "','" & CompanyID &_
			"','" & ProjectID & "','" & mechanismID & "','" & contactid & "','"	& reasons & "'"
			Set rsSelect = con.GetRecordSet(sqlstr)
			selAll = rsSelect.RecordCount
	                  
			If pLanguage = "heb" Then
				fori = CInt(pField_Scale)
				fotto = 1
				forstep = -1
			Else
				fori = 1
				fotto = CInt(pField_Scale)
				forstep = 1
			End If
	        
			For i = fori To fotto Step forstep
				rsSelect.Filter = "FIELD_VALUE='" & i & "'"
				ScQuant = rsSelect.RecordCount
				If ScQuant <> 0 Then
					wd = CInt((ScQuant / selAll) * 100)
					strDataSeries = strDataSeries & wd & ";;"
					strDataLabels = strDataLabels & i & ";;"
				Else
					wd = 0
				End If			         
			Next			
			Set rsSelect = Nothing
			
			If Len(strDataSeries) > 0 Then
				GetFieldGraph = "<img src=""../../chart/chart.asp?Title="& Server.URLEncode(pField_Title) &"&strDataSeries=" & Server.URLEncode(strDataSeries)  & "&strDataLabels=" & Server.URLEncode(strDataLabels) & """>"
			Else
				GetFieldGraph = ""
			End If	
		Case "5" 'check
			sqlstr = "Exec dbo.QUESTION_VALUES_PROC_Reasons '" & pProduct_Id & "','" & pField_Id & "','" & pType_ID &_
			"','" & start_date & "','" & end_date& "','" & UserID & "','" & is_groups & "','" & CompanyID &_
			"','" & ProjectID & "','" & mechanismID & "','" & ContactId & "','','"	& reasons & "'"
			Set rsSelect = con.GetRecordSet(sqlstr)			
			selAll = rsSelect.RecordCount
			rsSelect.Filter = "FIELD_VALUE='on'"
			ScQuant = rsSelect.RecordCount
			Set rsSelect = Nothing
			If ScQuant <> 0 Then
				wd = CInt((ScQuant / selAll) * 100)
				strDataSeries = strDataSeries & wd & ";;"
				strDataLabels = strDataLabels & "כן;;"
			Else
				wd = 0
			End If		
	                
			ScQuant = selAll - ScQuant
			If ScQuant <> 0 Then
				wd = CInt((ScQuant / selAll) * 100)
				strDataSeries = strDataSeries & wd & ";;"
				strDataLabels = strDataLabels & "לא;;"
			Else
				wd = 0
			End If			
			
			If Len(strDataSeries) > 0 Then	        
				GetFieldGraph = "<img src=""../../chart/chart.asp?Title="& Server.URLEncode(pField_Title) &"&strDataSeries=" & Server.URLEncode(strDataSeries)  & "&strDataLabels=" & Server.URLEncode(strDataLabels) & """>"
			Else
				GetFieldGraph = ""
			End If			
	        
		Case "6" 'date
			GetFieldGraph = ""
		Case "7" 'number
			GetFieldGraph = ""     
	        
		End Select
		'Response.Write sqlstr & "<br>"
		'Response.End
	End Function			
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<link href="../../../include/styles.css" rel="STYLESHEET" type="text/css">
</head>
<body marginwidth="0" marginheight="0" topmargin="0" leftmargin="0" rightmargin="0">
<table border="0" width="640" cellspacing="0" cellpadding="0" align=center>
<tr>
<td width="640">
<!--#INCLUDE FILE="logo_top.asp"-->
</td></tr> 
</table>
<table border="0" width="640" cellspacing="0" cellpadding="0" align=center>
<!-- start code --> 
	<tr>
		<td class="card5" style="font-size:16pt" align=center width=100% dir="<%=dir_obj_var%>"><b>
        <%If trim(lang_id) = "1" Then%>&nbsp; דוח התפלגות שאלות (עוגה)&nbsp;
        <%Else%>&nbsp;Pie distribution report&nbsp;<%End If%></b></td>
	</tr>
	<tr>
		<td class="card5" style="font-size:13pt" dir="<%=dir_obj_var%>" width=100% align="center"><b>			
<%
			sqlStr = "Select Product_Name, QUESTIONS_ID, Langu FROM PRODUCTS WHERE PRODUCT_ID=" & prodId &_
			" AND ORGANIZATION_ID = " & OrgID
			''Response.Write sqlStr
			''Response.End			
			set prod = con.GetRecordSet(sqlStr)
				if not prod.eof then
					productName=prod("Product_Name")					
					quest_id = prod("QUESTIONS_ID")
					if prod("Langu") = "eng" then
						dir_align = "ltr"
						td_align = "left"
						pr_language = "eng"
					else
						dir_align = "rtl"
						td_align = "right"
						pr_language = "heb"
					end if
				end if
			set prod = nothing%>
			&nbsp;<%=productName%>&nbsp;<%=strTitle%></b></td>
	</tr>
	<tr><td class="card5" style="font-size:13pt" dir="<%=dir_obj_var%>" width=100% align="center"><b><%=start_date%>  -  <%=end_date%></b></td></tr>
	<%'---------------P2932 - type of form- REASONS ---------------
   	If reasons<>"" Then
		sqlstr = "SELECT  Reason_Title FROM  Appeals_CreationReasons where QUESTIONS_ID=16504 and Reason_Id in (" & reasons & ") order by Reason_Order"
			set rs_Reason = con.getRecordSet(sqlstr)
			If Not rs_Reason.eof Then
				arr_Reason = rs_Reason.getRows()
			End If
			Set rs_Reason = Nothing
			If isArray(arr_Reason) Then
			%>
			<tr ><td class="card5" style="font-size:13pt" dir="rtl" width=100% align="center"><b> הגורם ליצירת הטופס:
 				<%For mm=0 To Ubound(arr_Reason,2)
					if mm>0 then%>
					,&nbsp;
					<%end if%>
					<%= trim(arr_Reason(0,mm))%>
				<%next%>
				</b></td></tr>
<%
			end if
	end if
	%>
	<tr><td class="card5" style="font-size:13pt" dir="<%=dir_obj_var%>" width=100% align="center"><b>מספר משובים - <%=appealsCount%></b></td></tr>	
	<!--#INCLUDE FILE="graph_fields_pie.asp"-->				
	<tr><td width="100%" height=5></td></tr>	
	<!--#INCLUDE FILE="bottom_inc.asp"-->	
</TABLE>
</body>
<%set con = nothing%>
</html>
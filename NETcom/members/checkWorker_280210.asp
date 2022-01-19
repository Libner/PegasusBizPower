<%
if Request.Form("username")<>nil then
	loginname=Request.Form("username")
	password=Request.Form("password")
	wizard_id=Request.Form("wizard_id")
	
	sqlstring = "SELECT USER_ID, USERS.ORGANIZATION_ID, USERS.LOGINNAME, USERS.PASSWORD,"&_
	" CHIEF ,FIRSTNAME,LASTNAME,ORGANIZATIONS.ORGANIZATION_NAME,RowsInList,ORGANIZATION_LOGO, LANG_ID, " &_
	" USERS.WORK_PRICING,USERS.SURVEYS,USERS.EMAILS,USERS.COMPANIES,USERS.TASKS,USERS.CASH_FLOW, "&_
	" IsNULL(ORGANIZATIONS.FILE_UPLOAD, 0) as FILE_UPLOAD, IsNULL(USERS.EDIT_APPEAL, 0) as EDIT_APPEAL, "&_
	" isNULL(UseBizLogo, 0) as UseBizLogo FROM USERS INNER JOIN ORGANIZATIONS ON " &_
	" USERS.ORGANIZATION_ID=ORGANIZATIONS.ORGANIZATION_ID "&_
	" WHERE ORGANIZATIONS.ACTIVE = 1 and USERS.ACTIVE = 1 and USERS.loginName='"& sFix(loginname) &_
	"' AND USERS.Password='"& sFix(password) &"'"
	'Response.Write sqlstring
	'Response.End
	set worker=con.GetRecordSet(sqlstring)
	if worker.EOF  then	
		Response.Redirect strLocal & "?error=true"
	else	
		Response.Cookies("bizpegasus") = "" 'clear cookie
	    'Response.Cookies("bizpegasus").Domain = Request.ServerVariables("SERVER_NAME")
	    
		Response.Cookies("bizpegasus")("UserId") = trim(worker("USER_ID"))
		Response.Cookies("bizpegasus")("wizardId") = wizard_id
		Response.Cookies("bizpegasus")("OrgID") = trim(worker("ORGANIZATION_ID"))
		Response.Cookies("bizpegasus")("LANGID") = worker("LANG_ID")
		Response.Cookies("bizpegasus")("ORGNAME") = trim(worker("ORGANIZATION_NAME"))
		Response.Cookies("bizpegasus")("UserName") = trim(worker("FIRSTNAME")) & " " & trim(worker("LASTNAME"))
		Response.Cookies("bizpegasus")("RowsInList") = trim(worker("RowsInList"))	
		Response.Cookies("bizpegasus")("perSize") = worker("ORGANIZATION_LOGO").ActualSize
		Response.Cookies("bizpegasus")("Chief") = trim(worker("Chief"))
		Response.Cookies("bizpegasus")("WORKPRICING") = trim(worker("WORK_PRICING"))		
		Response.Cookies("bizpegasus")("SURVEYS") = trim(worker("SURVEYS"))		
		Response.Cookies("bizpegasus")("EMAILS") = trim(worker("EMAILS"))		
		Response.Cookies("bizpegasus")("COMPANIES") = trim(worker("COMPANIES"))		
		Response.Cookies("bizpegasus")("TASKS") = trim(worker("TASKS"))		
		Response.Cookies("bizpegasus")("CASHFLOW") = trim(worker("CASH_FLOW"))
		Response.Cookies("bizpegasus")("FILEUP") = trim(worker("FILE_UPLOAD"))	
		Response.Cookies("bizpegasus")("EDITAPPEAL") = trim(worker("EDIT_APPEAL"))	
		Response.Cookies("bizpegasus")("UseBizLogo") = trim(worker("UseBizLogo"))		
		
		sqlStr = "SELECT TOP 1 Bar_Id FROM bar_organizations WHERE Bar_ID = 10 AND IsNull(is_visible,0) = 1 AND ORGANIZATION_ID= "& trim(worker("ORGANIZATION_ID"))
		'Response.Write sqlStr
		'Response.End
		set rs_companies = con.GetRecordSet(sqlStr)
		If not rs_companies.eof Then
			is_companies = 1
		Else
			is_companies = 0	
		End If
		set rs_companies = Nothing		
		
		sqlStr = "SELECT TOP 1 Bar_Id FROM bar_organizations WHERE Bar_ID = 43 AND IsNull(is_visible,0) = 1 AND ORGANIZATION_ID= "& trim(worker("ORGANIZATION_ID"))
		'Response.Write sqlStr
		'Response.End
		set rs_groups = con.GetRecordSet(sqlStr)
		If not rs_groups.eof Then
			is_groups = 1
		Else
			is_groups = 0	
		End If
		set rs_groups = Nothing
		
		sqlStr = "SELECT TOP 1 Bar_Id FROM bar_users WHERE Bar_ID = 46 AND IsNull(is_visible,0) = 1 AND ORGANIZATION_ID= "& trim(worker("ORGANIZATION_ID")) &_
		" AND User_ID = " & trim(worker("USER_ID"))
		'Response.Write sqlStr
		'Response.End
		set rs_meetings = con.GetRecordSet(sqlStr)
		If not rs_meetings.eof Then
			is_meetings = 1
		Else
			is_meetings = 0	
		End If
		set rs_meetings = Nothing		
		
		Response.Cookies("bizpegasus")("ISCOMPANIES") = cStr(is_companies)
		Response.Cookies("bizpegasus")("ISGROUPS") = cStr(is_groups)
		Response.Cookies("bizpegasus")("ISMEETINGS") = cStr(is_meetings)		
		
		Response.Cookies("bizpegasus").Expires = Now() + 30
		Session.Timeout=300
		
		UserId = trim(Request.Cookies("bizpegasus")("UserId"))
		OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))	
	
		'sql_obj="Select bar_id, bar_title, is_visible From dbo.bar_users_table('" & OrgID & "','" & UserID & "')" &_
		'" WHERE parent_id IS NULL Order By parent_order, bar_order"
		sql_obj = "Exec dbo.get_user_bars '" & OrgID & "','" & UserID & "',0"
		'Response.Write sql_obj
		set rs_obj=con.getRecordSet(sql_obj)
		if not rs_obj.eof then
				arr_bars_top = Split(rs_obj.Getstring(,,",",";") , ";")   
				session("arr_bars_top") = arr_bars_top
		set rs_obj = Nothing
		end if
		   
		'sql_obj="Select bar_id, bar_title, is_visible, parent_id, bar_url, parent_order From "&_
		'" dbo.bar_users_table('" & OrgID & "','" & UserID & "')" &_
		'" WHERE parent_id IS NOT NULL Order By Parent_Order, bar_order"
		sql_obj = "Exec dbo.get_user_bars '" & OrgID & "','" & UserID & "',1"
		set rs_obj=con.getRecordSet(sql_obj)
		if not rs_obj.eof then
				arr_bars = Split(rs_obj.Getstring(,,",",";") , ";")  
				session("arr_bars") = arr_bars 								
		set rs_obj = Nothing
		end if 
		   
		links_array = Array()
		'sqlstr = "Select br.bar_id, br.bar_order, (SELECT TOP 1 dbo.bars.bar_url FROM dbo.bar_users Inner Join " &_
		'" dbo.bar_organizations	On dbo.bar_users.bar_id = dbo.bar_organizations.bar_id LEFT OUTER JOIN dbo.bars " &_
		'" ON dbo.bar_users.bar_id = dbo.bars.bar_id	Where dbo.bar_users.organization_id = " & OrgID &_
		'" And dbo.bar_organizations.organization_id = " & OrgID & " And dbo.bar_users.user_id = " & UserID &_
		'" And dbo.bars.in_use = '1' And dbo.bars.parent_id = br.bar_id) FROM bars br WHERE br.parent_id IS NULL AND br.in_use = 1"
		sqlstr = "Exec dbo.get_user_bars_links '" & OrgID & "','" & UserID & "'"
		'Response.Write sqlstr
		'Response.End 
		set rs_parents = con.getRecordSet(sqlstr)
		If not rs_parents.eof Then
			redim links_array(rs_parents.recordCount)
			While not rs_parents.eof
				parent_id = trim(rs_parents(0))
				parent_order = trim(rs_parents(1))
				links_array(parent_order) = trim(rs_parents(2))				
			rs_parents.moveNext
			Wend
		set rs_parents = Nothing
		End If 
		session("links_array") = links_array
			
		sql_obj="Select object_id,title_organization_one,title_organization_multi from titles WHERE ORGANIZATION_ID=" & trim(Request.Cookies("bizpegasus")("OrgID")) & " order by object_id"
		'Response.Write sql_obj
		'Response.End
		set rs_org=con.getRecordSet(sql_obj)
		do while not rs_org.eof
			select case rs_org("object_id")
				case 1
					Response.Cookies("bizpegasus")("Projectone")=trim(rs_org("title_organization_one"))
					Response.Cookies("bizpegasus")("ProjectMulti")=trim(rs_org("title_organization_multi"))
				case 2
					Response.Cookies("bizpegasus")("CompaniesOne")=trim(rs_org("title_organization_one"))
					Response.Cookies("bizpegasus")("CompaniesMulti")=trim(rs_org("title_organization_multi"))
				case 3
					Response.Cookies("bizpegasus")("ContactsOne")=trim(rs_org("title_organization_one"))
					Response.Cookies("bizpegasus")("ContactsMulti")=trim(rs_org("title_organization_multi"))
				case 4
					Response.Cookies("bizpegasus")("TasksOne")=trim(rs_org("title_organization_one"))
					Response.Cookies("bizpegasus")("TasksMulti")=trim(rs_org("title_organization_multi"))
				case 5
					Response.Cookies("bizpegasus")("ActivitiesOne")=trim(rs_org("title_organization_one"))
					Response.Cookies("bizpegasus")("ActivitiesMulti")=trim(rs_org("title_organization_multi"))					
 			End Select
		 
		rs_org.MoveNext
		loop 
		set rs_org = nothing	
    
	end if
	set worker = nothing
	
else	
	wizard_id = trim(Request.Cookies("bizpegasus")("wizardId"))	
end if

	If Request.QueryString("UserID") <> nil Then 'כניסה ממייל
		mUserID = trim(Request.QueryString("UserID"))
		If trim(mUserID) <> trim(Request.Cookies("bizpegasus")("UserId")) Then
			Response.Redirect strLocal
		End If
	End If

	If IsArray(session("arr_bars")) = false Or IsArray(session("arr_bars_top"))  = false Then
		Response.Redirect strLocal
	End If
	 
	If Len(Request.Cookies("bizpegasus")) = 0 Then
		Response.Redirect strLocal & "?error=true"
	End If 

    UserId = trim(Request.Cookies("bizpegasus")("UserId"))
    OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
	org_name = trim(Request.Cookies("bizpegasus")("ORGNAME"))
	user_name = trim(Request.Cookies("bizpegasus")("UserName"))
	RowsInList = Request.Cookies("bizpegasus")("RowsInList")	
	chief = trim(Request.Cookies("bizpegasus")("Chief"))
	perSize = trim(Request.Cookies("bizpegasus")("perSize"))
	WORK_PRICING = trim(Request.Cookies("bizpegasus")("WORKPRICING"))
	SURVEYS = trim(Request.Cookies("bizpegasus")("SURVEYS"))
	EMAILS = trim(Request.Cookies("bizpegasus")("EMAILS"))
	COMPANIES = trim(Request.Cookies("bizpegasus")("COMPANIES"))
	TASKS = trim(Request.Cookies("bizpegasus")("TASKS"))
	CASH_FLOW = trim(Request.Cookies("bizpegasus")("CASHFLOW"))	
	FILEUP = trim(Request.Cookies("bizpegasus")("FILEUP"))		
	lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))
	EDIT_APPEAL = trim(Request.Cookies("bizpegasus")("EDITAPPEAL"))
	is_companies = trim(Request.Cookies("bizpegasus")("ISCOMPANIES"))
	is_groups = trim(Request.Cookies("bizpegasus")("ISGROUPS"))
	is_meetings = trim(Request.Cookies("bizpegasus")("ISMEETINGS"))
		
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"  :  align_var = "left"  :  dir_obj_var = "ltr"
	Else
		dir_var = "ltr"  :  align_var = "right"  :  dir_obj_var = "rtl"
	End If
	    
    ' כניסה מהתוכנה Desktop
    formid = trim(Request.Form("formid"))
    form_type = trim(Request.Form("form_type"))
    newid = trim(Request.Form("newid"))
    task_inout = trim(Request.Form("task_inout"))
    task_date = trim(Request.Form("date_task"))
    'enter from web pegasus
    OpenContactId = trim(Request.Form("OpenContactId"))
    
    If formid <> "" Then
		if form_type = "1" then
			Response.Redirect "members/appeals/feedback.asp?appid=" & formid
		elseif form_type = "2" then
    		Response.Redirect "members/appeals/appeal_card.asp?appid=" & formid
    	end if	
    ElseIf task_inout <> "" Then
    	Response.Redirect "members/tasks/default.asp?T=" & task_inout & "&start_date=" & task_date & "&end_date=" & task_date
    ElseIf newid <> "" Then
		Response.Redirect "news.asp?ID=" & newid
	ElseIf OpenContactId <> "" And OpenContactId <> "0" Then
	    	Response.Redirect "members/companies/contact.asp?ContactId=" & OpenContactId
	ElseIf OpenContactId = "0" Then
	    	Response.Redirect "members/companies/default.asp"
    End If  %>
<script language="javascript">
<!-- 
	function SetCookie(sName, sValue)
	{  
	document.cookie = sName + "=" + escape(sValue);
	}
	SetCookie("ScreenWidth",screen.width);
// End -->
</script>
<%For each Item In Request.Cookies
		If Item = "ScreenWidth" Then
			strScreenWidth = Request.Cookies(item) 
		End if
	Next     			
	If strScreenWidth = "" Then strScreenWidth = 800	    %>
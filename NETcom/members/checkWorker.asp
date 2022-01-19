<%
'response.Write "strLocal="& strLocal
'response.end
'response.Write "check"
'Response.Cookies("bizpegasus").Expires
''response.Write ( Response.Cookies("bizpegasus").Expires)
'dim x,y
'for each x in Request.Cookies
'  response.write("<p>")
'  if Request.Cookies(x).HasKeys then
'    for each y in Request.Cookies(x)
'      response.write(x & ":" & y & "=" & Request.Cookies(x)(y))
'      response.write("<br />")
'    next
'  else
'    Response.Write(x & "=" & Request.Cookies(x) & "<br />")
'  end if
' response.write "</p>"
'next
if Request.Form("username")<>nil then
	loginname=Request.Form("username")
	password=Request.Form("password")
	wizard_id=Request.Form("wizard_id")
	
	sqlstring = "SELECT USER_ID, USERS.ORGANIZATION_ID, USERS.LOGINNAME, USERS.PASSWORD,"&_
	" CHIEF ,FIRSTNAME,LASTNAME,ORGANIZATIONS.ORGANIZATION_NAME,RowsInList,ORGANIZATION_LOGO, LANG_ID, " &_
	" USERS.WORK_PRICING,USERS.SURVEYS,USERS.EMAILS,USERS.COMPANIES,USERS.TASKS,USERS.CASH_FLOW, "&_
	" IsNULL(ORGANIZATIONS.FILE_UPLOAD, 0) as FILE_UPLOAD, IsNULL(USERS.EDIT_APPEAL, 0) as EDIT_APPEAL,IsNULL(Edit_DepartmentAppeal,0) as Edit_DepartmentAppeal, " &_
	" isNULL(salesControl_Edit,0) as salesControl_Edit,isNULL(salesControl_Points_Edit,0) as salesControl_Points_Edit, "&_
	" isNULL(EditFlyCompanies,0) as EditFlyCompanies,isNULL(EditAirports,0) as EditAirports,isNULL(EditApprovalStatus,0) as EditApprovalStatus,isNULL(EditCurrency,0) as EditCurrency, " & _
	" isNULL(EditFlightStatus,0) as EditFlightStatus, isNULL(Edit_IsNewLetters,0) as Edit_IsNewLetters, " &_
	" isNULL(ExportDataToExcel,0) as ExportDataToExcel,isNULL(EditMainData,0) as EditMainData,isNULL(InsertRowMainData,0) as InsertRowMainData, " &_
	" isNULL(DeleteRowMainData,0) as DeleteRowMainData,isNULL(UseBizLogo, 0) as UseBizLogo,isNULL(Add_Insurance, 0) as Add_Insurance, " &_
	" isNULL(IP_login, 0) as IP_login,isNULL( Add_GroupsTours, 0) as  Add_GroupsTours,isNULL( Sms_Write, 0) as  Sms_Write, " &_
	" isNULL(EmailGroupSend, 0) as  EmailGroupSend,isNULL(User_Screen_TourVisible, 0) as  User_Screen_TourVisible, " &_
	" 	isNULL(User_ScreenTourStatus, 0) as  User_ScreenTourStatus,isNULL(User_ScreenSendMail, 0) as  User_ScreenSendMail, " &_
	" isNULL(User_ScreenTourOrder, 0) as  User_ScreenTourOrder, isNULL(User_SendSms_GeneralScreen, 0) as  User_SendSms_GeneralScreen, " &_
	"  isNULL(User_SendSms_ContactScreen, 0) as  User_SendSms_ContactScreen,isNULL(User_OperationScreenView, 0) as  User_OperationScreenView, " &_
	" isNULL(ReportMaakavRishumView, 0) as  ReportMaakavRishumView,isNULL(ReportCloseProcView, 0) as  ReportCloseProcView, " &_
	" isNULL(DashBoardView, 0) as  DashBoardView,isNULL(UpdateFieldUserBlocked, 0) as  UpdateFieldUserBlocked, " &_
	" isNULL(Archive_appeal, 0) as  Archive_appeal, isNULL(UndoPNR, 0) as  UndoPNR, isNULL(IsSendFromTL, 0) as  IsSendFromTL, " &_
	"  JOB_ID,isNULL(IsTypingClient,0) as IsTypingClient,isNULL(InsertNewUser,0) as InsertNewUser, " &_
	" isNULL(PermissionsUploadFiles,0) as PermissionsUploadFiles,isNULL(PermissionsUploadFiles_View,0) as PermissionsUploadFilesView, " &_
	" isNULL(PermisionAdmin,0) as PermisionAdmin,isNull(UploadUserFiles,0) as UploadUserFiles,isNULL(VerificationCode,0) as VerificationCode, " &_
	" isNULL(VerificationCode_Confirmed,0) as VerificationCode_Confirmed,isNull(UploadUserFiles_View,0) as UploadUserFilesView, " &_
	" isNull(UploadUserFiles_Delete,0) as UploadUserFilesDelete,isNull(EditTourCodeTimeLimiting,0)as EditTourCodeTimeLimiting, " &_
	" isNull(AddAppealFromLogToursInteresting,0) as AddAppealFromLogToursInteresting,IsNull(EditPassword,0) as EditPassword, " &_
	" isNull(SendMarketingEmail,0) as SendMarketingEmail,isNull(AddTravelerToTour,0) as AddTravelerToTour,isNull(DelTravelerToTour,0) as DelTravelerToTour, " &_
	" isNull(UpdateTravelerToTour,0) as UpdateTravelerToTour, " & _
	" isNull(AddTraveler,0) as AddTraveler,isNull(DelTraveler,0) as DelTraveler,isNull(UpdateTraveler,0) as UpdateTraveler , " &_
	" isNull(DeleteVerificationCode,0) as DeleteVerificationCode, " &_
	" isNull(UpdateTimingMailReservationForm,0) as UpdateTimingMailReservationForm, " &_
	" isNull(SendMailReservationForm,0) as SendMailReservationForm, " &_
	" isNull(GetTravelersByDocket ,0) as GetTravelersByDocket " &_
	" FROM USERS INNER JOIN ORGANIZATIONS ON " &_
	" USERS.ORGANIZATION_ID=ORGANIZATIONS.ORGANIZATION_ID"&_
	" WHERE ORGANIZATIONS.ACTIVE = 1 and USERS.ACTIVE = 1 and User_Bloked=0 and USERS.loginName='"& sFix(loginname) &_
	"' AND USERS.Password='"& sFix(password) &"'"
'Response.Write sqlstring
	'Response.End
	set worker=con.GetRecordSet(sqlstring)
		if worker.EOF  then	
		set worker=Nothing
						sqlstring = "SELECT USER_ID from Users where   User_Bloked=1 and USERS.loginName='"& sFix(loginname) &"'"
						set worker=con.GetRecordSet(sqlstring)
						if not worker.EOF  then	
						set worker=Nothing
						Response.Redirect strLocal & "?error=true&blocked=true"
						end if
						sqlstring = "SELECT count(Users.USER_ID) as CountEnters from Users  left join FailedLogins on FailedLogins.LOGINNAME=USERS.loginName where  USERS.ACTIVE = 1 and User_Bloked=0 and USERS.loginName='"& sFix(loginname) &"' and datediff(mi,logintime,getDate())<10"
						set worker=con.GetRecordSet(sqlstring)
						if not worker.EOF  then
				
						CountEnters=worker("CountEnters")
					    if CountEnters=4 then
						con.ExecuteQuery("update users set User_Bloked=1 where  USERS.loginName='"& sFix(loginname) &"'")
						Response.Redirect strLocal & "?error=true&blocked=true"
						set worker=Nothing
						end if
						if CountEnters<4 then
						
							sqlU="Select top 1 User_Id from Users where LOGINNAME='"&sFix(loginname) &"'"
							set w=con.GetRecordSet(sqlU)
							if not w.EOF then
							UserId= w("User_Id")
							con.ExecuteQuery("insert into FailedLogins (LOGINNAME,User_Id) values ('"& sFix(loginname) &"'," & UserId &")")
							end if
												
						Response.Redirect strLocal & "?error=true&PasswordError=true"
						set worker=Nothing
						end if
						end if
						
						
						
						
		else	
	'	response.Write worker("VerificationCode")&":" & worker("VerificationCode_Confirmed")
'response.Write "code1"
	VerificationCode=worker("VerificationCode")
	VerificationCode_Confirmed= worker("VerificationCode_Confirmed")
			'		response.Write "VerificationCode"
			'		response.end
			if  trim(worker("IP_login"))=1 then
			sqlIp="Select top 1 IP from Ip_Address"
			set wIP=con.GetRecordSet(sqlIp)
			if not wIP.EOF then
			IPadd= wIP("IP")
			end if
		
			if 	trim(IPadd)= Request.ServerVariables("REMOTE_ADDR")  then

			else
				'response.Write "REMOTE_ADDR="& Request.ServerVariables("REMOTE_ADDR") &IPadd
			'response.end
set worker=Nothing
				Response.Redirect strLocal & "?error=true&ip=true"
			end if

		end if

		Response.Cookies("bizpegasus") = "" 'clear cookie
		'Response.Cookies("bizpegasus").Domain="pegasus.bizpower.co.il"
	    Response.Cookies("bizpegasus").Domain = Request.ServerVariables("SERVER_NAME")
	    
		Response.Cookies("bizpegasus")("UserId") = trim(worker("USER_ID"))
		Response.Cookies("bizpegasus")("wizardId") = wizard_id
		Response.Cookies("bizpegasus")("OrgID") = trim(worker("ORGANIZATION_ID"))
		Response.Cookies("bizpegasus")("LANGID") = worker("LANG_ID")
		Response.Cookies("bizpegasus")("ORGNAME") = trim(worker("ORGANIZATION_NAME"))
		Response.Cookies("bizpegasus")("UserName") = trim(worker("FIRSTNAME")) & " " & trim(worker("LASTNAME"))
		Response.Cookies("bizpegasus")("RowsInList") = trim(worker("RowsInList"))	
		Response.Cookies("bizpegasus")("perSize") = worker("ORGANIZATION_LOGO").ActualSize
		Response.Cookies("bizpegasus")("CHIEFPerm") = trim(worker("Chief"))
		Response.Cookies("bizpegasus")("CHIEF") = trim(worker("Chief"))
		Response.Cookies("bizpegasus")("EditPassword") = trim(worker("EditPassword"))
	
		Response.Cookies("bizpegasus")("AddInsurance") = trim(worker("Add_Insurance"))
		Response.Cookies("bizpegasus")("Add_GroupsTours") = trim(worker("Add_GroupsTours"))
		Response.Cookies("bizpegasus")("Sms_Write") = trim(worker("Sms_Write"))
		Response.Cookies("bizpegasus")("EmailGroupSend") = trim(worker("EmailGroupSend"))
  
    	Response.Cookies("bizpegasus")("ScreenTourVisible") = trim(worker("User_Screen_TourVisible"))
		Response.Cookies("bizpegasus")("ScreenTourStatus") = trim(worker("User_ScreenTourStatus"))
		Response.Cookies("bizpegasus")("ScreenSendMail") = trim(worker("User_ScreenSendMail"))
		Response.Cookies("bizpegasus")("ScreenExpires")=Now()
		Response.Cookies("bizpegasus")("ScreenTourOrder") = trim(worker("User_ScreenTourOrder"))
	' response.Write ("22="& request.Cookies("bizpegasus")("User_ScreenTourOrder") )
	 'response.end
	
		Response.Cookies("bizpegasus")("UserSendSmsGeneralScreen") = trim(worker("User_SendSms_GeneralScreen"))
		Response.Cookies("bizpegasus")("UserSendSmsContactScreen") = trim(worker("User_SendSms_ContactScreen"))
		Response.Cookies("bizpegasus")("UserOperationScreenView") = trim(worker("User_OperationScreenView"))
		
		Response.Cookies("bizpegasus")("ReportCloseProcView") = trim(worker("ReportCloseProcView"))
		Response.Cookies("bizpegasus")("GetTravelersByDocket") = trim(worker("GetTravelersByDocket"))
		
		Response.Cookies("bizpegasus")("DashBoardView") = trim(worker("DashBoardView"))
		Response.Cookies("bizpegasus")("UpdateFieldUserBlocked") = trim(worker("UpdateFieldUserBlocked"))
	
		Response.Cookies("bizpegasus")("Archive_appeal") = trim(worker("Archive_appeal"))
		Response.Cookies("bizpegasus")("UndoPNR") = trim(worker("UndoPNR"))
		Response.Cookies("bizpegasus")("IsSendFromTL") = trim(worker("IsSendFromTL"))

		Response.Cookies("bizpegasus")("ReportMaakavRishumView") = trim(worker("ReportMaakavRishumView"))
	 
		Response.Cookies("bizpegasus")("WORKPRICING") = trim(worker("WORK_PRICING"))		
		Response.Cookies("bizpegasus")("SURVEYS") = trim(worker("SURVEYS"))		
		Response.Cookies("bizpegasus")("EMAILS") = trim(worker("EMAILS"))		
		Response.Cookies("bizpegasus")("COMPANIES") = trim(worker("COMPANIES"))		
		Response.Cookies("bizpegasus")("TASKS") = trim(worker("TASKS"))		
		Response.Cookies("bizpegasus")("CASHFLOW") = trim(worker("CASH_FLOW"))
		Response.Cookies("bizpegasus")("FILEUP") = trim(worker("FILE_UPLOAD"))	
		Response.Cookies("bizpegasus")("EDITAPPEAL") = trim(worker("EDIT_APPEAL"))	
		Response.Cookies("bizpegasus")("Edit_DepartmentAppealPerm") = trim(worker("Edit_DepartmentAppeal"))	
		Response.Cookies("bizpegasus")("EditDepartmentAppeal") = trim(worker("Edit_DepartmentAppeal"))	
		
		Response.Cookies("bizpegasus")("salesControl_Edit") = trim(worker("salesControl_Edit"))	
		Response.Cookies("bizpegasus")("salesControlEdit") = trim(worker("salesControl_Edit"))	
		Response.Cookies("bizpegasus")("salesControl_Points_Edit") = trim(worker("salesControl_Points_Edit"))	
		Response.Cookies("bizpegasus")("salesControlPointsEdit") = trim(worker("salesControl_Points_Edit"))	
	
		Response.Cookies("bizpegasus")("EditFlyCompanies") = trim(worker("EditFlyCompanies"))	
		Response.Cookies("bizpegasus")("EditAirports") = trim(worker("EditAirports"))	
		Response.Cookies("bizpegasus")("EditApprovalStatus") = trim(worker("EditApprovalStatus"))	
		Response.Cookies("bizpegasus")("EditCurrency") = trim(worker("EditCurrency"))	
		Response.Cookies("bizpegasus")("EditFlightStatus") = trim(worker("EditFlightStatus"))	
		Response.Cookies("bizpegasus")("Edit_IsNewLetters") = trim(worker("Edit_IsNewLetters"))	
		Response.Cookies("bizpegasus")("ExportDataToExcel") = trim(worker("ExportDataToExcel"))	
		
		Response.Cookies("bizpegasus")("EditMainData") = trim(worker("EditMainData"))	
		Response.Cookies("bizpegasus")("InsertRowMainData") = trim(worker("InsertRowMainData"))	
		Response.Cookies("bizpegasus")("DeleteRowMainData") = trim(worker("DeleteRowMainData"))	
		
		Response.Cookies("bizpegasus")("IsTypingClient") = trim(worker("IsTypingClient"))	
		Response.Cookies("bizpegasus")("InsertNewUser") = trim(worker("InsertNewUser"))	
		Response.Cookies("bizpegasus")("PermissionsUploadFiles") = trim(worker("PermissionsUploadFiles"))	
		Response.Cookies("bizpegasus")("UploadUserFiles") = trim(worker("UploadUserFiles"))	
		Response.Cookies("bizpegasus")("PermissionsUploadFilesView") = trim(worker("PermissionsUploadFilesView"))	
		Response.Cookies("bizpegasus")("UploadUserFilesView") = trim(worker("UploadUserFilesView"))	
		Response.Cookies("bizpegasus")("UploadUserFilesDelete") = trim(worker("UploadUserFilesDelete"))	
		Response.Cookies("bizpegasus")("EditTourCodeTimeLimiting") = trim(worker("EditTourCodeTimeLimiting"))
	
		Response.Cookies("bizpegasus")("PermisionAdmin") = trim(worker("PermisionAdmin"))	
	
		Response.Cookies("bizpegasus")("AddAppealFromLogToursInteresting") = trim(worker("AddAppealFromLogToursInteresting"))	
		Response.Cookies("bizpegasus")("SendMarketingEmail") = trim(worker("SendMarketingEmail"))	
		
		Response.Cookies("bizpegasus")("UseBizLogo") = trim(worker("UseBizLogo"))		
		Response.Cookies("bizpegasus")("JobId") = trim(worker("JOB_ID"))
		
		
		Response.Cookies("bizpegasus")("AddTravelerToTour") = trim(worker("AddTravelerToTour"))
		Response.Cookies("bizpegasus")("DelTravelerToTour") = trim(worker("DelTravelerToTour"))
		Response.Cookies("bizpegasus")("UpdateTravelerToTour") = trim(worker("UpdateTravelerToTour"))
		
			
		Response.Cookies("bizpegasus")("AddTraveler") = trim(worker("AddTraveler"))
		Response.Cookies("bizpegasus")("DelTraveler") = trim(worker("DelTraveler"))
		Response.Cookies("bizpegasus")("UpdateTraveler") = trim(worker("UpdateTraveler"))
		Response.Cookies("bizpegasus")("DeleteVerificationCode") = trim(worker("DeleteVerificationCode"))
				
		
		Response.Cookies("bizpegasus")("UpdateTimingMailReservationForm") = trim(worker("UpdateTimingMailReservationForm"))		
	Response.Cookies("bizpegasus")("SendMailReservationForm") = trim(worker("SendMailReservationForm"))		
				
		''Response.Cookies("bizpegasus").Expires =DateTime.Now.AddHours(1)

		'--insert into changes table
		sqlstr = "INSERT INTO Changes (Ip,Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
		" SELECT '"& Request.ServerVariables("REMOTE_ADDR")  &"', 'כניסה למערכת', IsNULL(U.FIRSTNAME, '') + ' ' + IsNULL(LASTNAME, ''), [USER_ID], 'כניסה', getDate(), " & trim(worker("USER_ID")) & _
		" FROM dbo.USERS U WHERE ([USER_ID] = " & trim(worker("USER_ID")) & ")"
		'response.Write sqlstr
		'response.end
		con.executeQuery(sqlstr)	
		
		UserId=worker("USER_ID")
		'DateLogIn=Now()
		sql_obj = "Exec dbo.insert_Users_LogIn '" & UserID &"'"
		'response.Write sql_obj
	'	response.end
	    con.executeQuery (sql_obj)
		
		
	'''	sqlstr="Insert into Users_LogIn  (USER_ID) SELECT User_Id WHERE not exists (select * from Users_LogIn where USER_ID = "& worker("USER_ID") &"  and LogInDate = getDate() )
		
	'       response.Write      sqlstr
     ' response.end      
	''con.executeQuery(sqlstr)	
				
		
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
		dEnd=FormatDateTime(Now(),2) &" 23:59:59"
		dEnd=DateAdd("h",12,Now())
		'response.Write "dEnd="& dEnd
		
		Response.Cookies("bizpegasus").Expires =dEnd ' DateDiff("s",Now(),dEnd)  '' Now() + 1
		'response.Write "ddd=" & DateDiff("s",Now(),dEnd) &"<BR>"
		'response.Write (Now() &":"& 	Now()+1)
		Session.Timeout=600
		
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
				'Response.Cookies("bizpegasus")("arr_bars_top") = arr_bars_top
		set rs_obj = Nothing
		end if
		   
		'sql_obj="Select bar_id, bar_title, is_visible, parent_id, bar_url, parent_order From "&_
		'" dbo.bar_users_table('" & OrgID & "','" & UserID & "')" &_
		'" WHERE parent_id IS NOT NULL Order By Parent_Order, bar_order"
		sql_obj = "Exec dbo.get_user_bars '" & OrgID & "','" & UserID & "',1"
				'response.Write sql_obj 

		set rs_obj=con.getRecordSet(sql_obj)
		if not rs_obj.eof then
				arr_bars = Split(rs_obj.Getstring(,,",",";") , ";")  
				session("arr_bars") = arr_bars 		
				'Response.Cookies("bizpegasus")("arr_bars") = arr_bars						
		set rs_obj = Nothing
		end if 
		   
		links_array = Array()
		'sqlstr = "Select br.bar_id, br.bar_order, (SELECT TOP 1 dbo.bars.bar_url FROM dbo.bar_users Inner Join " &_
		'" dbo.bar_organizations	On dbo.bar_users.bar_id = dbo.bar_organizations.bar_id LEFT OUTER JOIN dbo.bars " &_
		'" ON dbo.bar_users.bar_id = dbo.bars.bar_id	Where dbo.bar_users.organization_id = " & OrgID &_
		'" And dbo.bar_organizations.organization_id = " & OrgID & " And dbo.bar_users.user_id = " & UserID &_
		'" And dbo.bars.in_use = '1' And dbo.bars.parent_id = br.bar_id) FROM bars br WHERE br.parent_id IS NULL AND br.in_use = 1"
		sqlstr = "Exec dbo.get_user_bars_links '" & OrgID & "','" & UserID & "'"
	'	Response.Write sqlstr
		'Response.End 
		set rs_parents = con.getRecordSet(sqlstr)
		If not rs_parents.eof Then
			redim links_array(rs_parents.recordCount)
			While not rs_parents.eof
				parent_id = trim(rs_parents(0))
				parent_order = trim(rs_parents(1))
				links_array(parent_order) = trim(rs_parents(2))				
		'	response.Write (parent_id &":"& parent_order &":"& trim(rs_parents(2))	&":"&"<BR>")
			rs_parents.moveNext
			Wend
		'	response.end
		set rs_parents = Nothing
		End If 
		session("links_array") = links_array
		
		'<------------------------------------------------------------------------------------------------------->
		dim arr_status_()
	    
		sqlstr = "Select Max(status_id) From statuses"
		set rs_max = con.getRecordSet(sqlstr)
		If not rs_max.eof Then
			max_status = trim(rs_max(0))
		End If
		set rs_max = Nothing
		
		If IsNumeric(max_status) Then
			max_status = cInt(max_status)	
			Redim arr_status_(max_status,2)

			sqlstr = "SELECT status_id, status_name, status_color, status_order FROM statuses ORDER BY status_order"
			set rs_st = con.getRecordSet(sqlstr)
			if not rs_st.eof then			
				while not rs_st.eof 					
					arr_status_(trim(rs_st(3)),0) =  trim(rs_st(0))   'status_id
					arr_status_(trim(rs_st(3)),1) =  trim(rs_st(1))   'status_name
					arr_status_(trim(rs_st(3)),2) =  trim(rs_st(2))   'status_color
					rs_st.moveNext
				wend	
			end if
			set rs_st = nothing
			session("arr_Status") = arr_status_    
		End If
			
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
			con.executeQuery("insert into errorsCheck(itemDesc) values('mUserID <> UserId - UserId = "& Request.Cookies("bizpegasus")("UserId") &"')")			
			Response.Redirect strLocal
		End If
	End If

	If IsArray(session("arr_bars")) = false Or IsArray(session("arr_bars_top"))  = false Then
		con.executeQuery("insert into errorsCheck(itemDesc) values('session arr_bars expire - UserId = "& Request.Cookies("bizpegasus")("UserId") &"')")	
'response.Write "11111"
'--------response.end
	sql_obj = "Exec dbo.get_user_bars '" & Request.Cookies("bizpegasus")("OrgID") & "','" & Request.Cookies("bizpegasus")("UserId") & "',0"
	'	Response.Write sql_obj
'		response.end
		set rs_obj=con.getRecordSet(sql_obj)
		if not rs_obj.eof then
				arr_bars_top = Split(rs_obj.Getstring(,,",",";") , ";")   
				session("arr_bars_top") = arr_bars_top
		set rs_obj = Nothing
		end if
		   
		'sql_obj="Select bar_id, bar_title, is_visible, parent_id, bar_url, parent_order From "&_
		'" dbo.bar_users_table('" & OrgID & "','" & UserID & "')" &_
		'" WHERE parent_id IS NOT NULL Order By Parent_Order, bar_order"
		sql_obj = "Exec dbo.get_user_bars '" &  Request.Cookies("bizpegasus")("OrgID")  & "','" & Request.Cookies("bizpegasus")("UserId")  & "',1"
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
		sqlstr = "Exec dbo.get_user_bars_links '" & Request.Cookies("bizpegasus")("OrgID")   & "','" & Request.Cookies("bizpegasus")("UserId") & "'"
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
	'			dim arr_status_()
	    
		sqlstr = "Select Max(status_id) From statuses"
		set rs_max = con.getRecordSet(sqlstr)
		If not rs_max.eof Then
			max_status = trim(rs_max(0))
		End If
		set rs_max = Nothing
		
		If IsNumeric(max_status) Then
			max_status = cInt(max_status)	
			Redim arr_status_(max_status,2)

			sqlstr = "SELECT status_id, status_name, status_color, status_order FROM statuses ORDER BY status_order"
			set rs_st = con.getRecordSet(sqlstr)
			if not rs_st.eof then			
				while not rs_st.eof 					
					arr_status_(trim(rs_st(3)),0) =  trim(rs_st(0))   'status_id
					arr_status_(trim(rs_st(3)),1) =  trim(rs_st(1))   'status_name
					arr_status_(trim(rs_st(3)),2) =  trim(rs_st(2))   'status_color
					rs_st.moveNext
				wend	
			end if
			set rs_st = nothing
			session("arr_Status") = arr_status_    
		End If
'------------------------------end create new session
	'	Response.Redirect strLocal
	else
	'Response.Cookies("bizpegasus").Expires =Now() -1
	'Response.Cookies("bizpegasus") = ""
	'Response.Cookies("bizpegasus")("UserId")=""
	'''response.Write "112222---555999"
	End If
	
	If Len(Request.Cookies("bizpegasus")) = 0  or  Request.Cookies("bizpegasus")("ScreenExpires")= "" Then
		con.executeQuery("insert into errorsCheck(itemDesc) values('bizpegasus = 0 - UserId = "& Request.Cookies("bizpegasus")("UserId") &"')")	
		Response.Redirect strLocal & "?error=true"
	End If 
			if len(VerificationCode)>1 and VerificationCode_Confirmed="0" then
			'send VerificationCode
			 sms_phone = "036374000"
	    		cmdSelectUsers="SELECT USER_ID,MOBILE_PRIVATE as MOBILE,FIRSTNAME,LASTNAME FROM   USERS " & _
                "  where ACTIVE=1  and (MOBILE_PRIVATE<>'' or MOBILE_PRIVATE<>null) and User_Bloked=0 and  USER_ID=" & trim(Request.Cookies("bizpegasus")("UserId"))
               set dr_m = con.getRecordSet(cmdSelectUsers)
               if not dr_m.eof Then
							
								UserID = dr_m("USER_ID")
								MOBILE = dr_m("MOBILE")
								'  MOBILE = "0507740302"
						
						   If Trim(MOBILE) <> "" Then
                              'VerificationCode = random.Next(100000)
       						Randomize
							VerificationCode = Int((rnd*10000))+1

									If Left(MOBILE, 3) = "050" Or Left(MOBILE, 3) = "052" Or Left(MOBILE, 3) = "053" Or Left(MOBILE, 3) = "054" Or Left(MOBILE, 3) = "055" Or Left(MOBILE, 3) = "058" Or Left(MOBILE, 3) = "077" Then
										cmdInsert ="SET DATEFORMAT DMY; SET NOCOUNT ON;update USERS SET SendSMS_VerificationCode=GetDate(),VerificationCode_Confirmed=0,VerificationCode=" & VerificationCode & " WHERE USER_ID=" & UserID
										con.executeQuery(cmdInsert)
				                 
										SMS_content = " קוד האימות שנשלח כעת מאתר פגסוס " & ":" & VerificationCode
									' Dim sendUrl, strResponse, getUrl As String
									 postWS = ""
										postWS = "uid=2575&un=pegasus&msglong=" & Server.UrlEncode(SMS_content) & "&charset=iso-8859-8" 
										postWS = postWS &"&from=" & sms_phone & "&post=2&list=" & Server.UrlEncode(MOBILE) 
										postWS = postWS & "&desc=" & Server.UrlEncode("pegasus")
										'response.Write "postWS="&  postWS &"<BR>"
										sendUrl = "http://www.micropay.co.il/ExtApi/ScheduleSms.php"
										set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
										
	xmlhttp.open "POST", sendUrl,  false
	xmlhttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
     xmlhttp.send postWS
'
strStatus = xmlhttp.Status
strRetval = xmlhttp.responseText
set xmlhttp = nothing
								
									
										' ---block send sms---
							
										'response.Write strRetval
									'	response.end
										set xmlhttp = Nothing
									end if
							end if
           set  dr_m=Nothing
   

	end if
			'SendSMS_VerificationCode
			'VerificationCode
			
			
	    Response.Redirect "VerificationPage.aspx"
		end if

    UserId = trim(Request.Cookies("bizpegasus")("UserId"))
    OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
    JobId = trim(Request.Cookies("bizpegasus")("JobId"))
	org_name = trim(Request.Cookies("bizpegasus")("ORGNAME"))
	user_name = trim(Request.Cookies("bizpegasus")("UserName"))
	RowsInList = Request.Cookies("bizpegasus")("RowsInList")	
	chief = trim(Request.Cookies("bizpegasus")("Chief"))
	EditPassword=trim(Request.Cookies("bizpowerpegasus")("EditPassword"))
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
	Edit_DepartmentAppeal= trim(Request.Cookies("bizpegasus")("EditDepartmentAppeal"))
	salesControl_Edit= trim(Request.Cookies("bizpegasus")("salesControl_Edit"))
	
	salesControl_Points_Edit= trim(Request.Cookies("bizpegasus")("salesControl_Points_Edit"))
	
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
<%'ADDED BY MILA for CMPARE - autocomplete works not properly (show all - in wrong place (in left corner of the screen)
if instr(Request.ServerVariables("SCRIPT_NAME"),"/compare/")<=0 then%>
<script language="javascript" type="text/javascript">
<!-- 
	function SetCookie(sName, sValue)
	{  
	document.cookie = sName + "=" + escape(sValue);
	}
	SetCookie("ScreenWidth",screen.width);
// End -->
</script>
<%end if%>
<%For each Item In Request.Cookies
		If Item = "ScreenWidth" Then
			strScreenWidth = Request.Cookies(item) 
		End if
	Next     			
	If strScreenWidth = "" Then strScreenWidth = 800	    %>

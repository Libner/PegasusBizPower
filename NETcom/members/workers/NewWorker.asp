<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
 workerId = trim(Request.Cookies("bizpegasus")("UserId"))
  PermissionsUploadFiles = trim(Request.Cookies("bizpegasus")("PermissionsUploadFiles"))
  UploadUserFiles= trim(Request.Cookies("bizpegasus")("UploadUserFiles"))
 PermisionAdmin=trim(Request.Cookies("bizpegasus")("PermisionAdmin"))
  
 'response.Write workerId
PByDepartmentId=PermissionsByDepartmentId(workerId)
if OrgID<>nil and OrgID<>"" then
  sqlStr = "select ORGANIZATION_NAME from ORGANIZATIONS where ORGANIZATION_ID=" & OrgID  
  ''Response.Write sqlStr
  set rs_Org = con.GetRecordSet(sqlStr)  
  if not rs_Org.eof then
	ORGANIZATION_NAME = rs_Org("ORGANIZATION_NAME")						
  end if
  set rs_Org = nothing
end if

errLogin=false
errEmail=false
USER_ID = trim(Request("USER_ID"))

if request.form("FIRSTNAME")<>nil then 'after form filling
	 FIRSTNAME = sFix(request.form("FIRSTNAME"))
	 FIRSTNAME_ENG= sFix(request.form("FIRSTNAME_ENG"))
	 LASTNAME  = sFix(trim(Request.Form("LASTNAME")))
	 LASTNAME_ENG  = sFix(trim(Request.Form("LASTNAME_ENG")))
	 LOGINNAME = sFix(request.form("LOGINNAME"))
	 PASSWORD  = sFix(request.form("PASSWORD"))
     phone     = sFix(trim(Request.Form("phone")))
     mobile    = sFix(trim(Request.Form("mobile")))
     mobilePrivate    = sFix(trim(Request.Form("mobilePrivate")))
     fax       = sFix(trim(Request.Form("fax")))
     email     = sFix(trim(Request.Form("email")))
     emailPrivate     = sFix(trim(Request.Form("emailPrivate")))
     birthday= sFix(trim(Request.Form("birthday")))
'     response.Write birthday
    	if len(BIRTHDAY) then
			else
			 BIRTHDAY="NULL"
			end if
     Month_Min_Order  = sFix(trim(Request.Form("Month_Min_Order")))
     If Request.Form("EDIT_APPEAL") <> nil Then
		EDIT_APPEAL = sFix(trim(Request.Form("EDIT_APPEAL")))     
     Else
		EDIT_APPEAL = "0"
     End If
   
     
     if Request.Form("salesControl_Edit")<>nil then
			 if Request.Form("salesControl_Edit")="on" then
				salesControl_Edit="1"
			 else
				salesControl_Edit="0"
			 end if
     else
     salesControl_Edit="0"
     end if
     if Request.Form("EditFlyCompanies")<>nil then
			 if Request.Form("EditFlyCompanies")="on" then
				EditFlyCompanies="1"
			 else
				EditFlyCompanies="0"
			 end if
     else
     EditFlyCompanies="0"
     end if
 
	 if Request.Form("EditAirports")<>nil then
			 if Request.Form("EditAirports")="on" then
				EditAirports="1"
			 else
				EditAirports="0"
			 end if
     else
     EditAirports="0"
     end if
 
	 if Request.Form("EditApprovalStatus")<>nil then
			 if Request.Form("EditApprovalStatus")="on" then
				EditApprovalStatus="1"
			 else
				EditApprovalStatus="0"
			 end if
     else
     EditApprovalStatus="0"
     end if
     if Request.Form("EditCurrency")<>nil then
			 if Request.Form("EditCurrency")="on" then
				EditCurrency="1"
			 else
				EditCurrency="0"
			 end if
     else
     EditCurrency="0"
     end if
	 if Request.Form("EditFlightStatus")<>nil then
			 if Request.Form("EditFlightStatus")="on" then
				EditFlightStatus="1"
			 else
				EditFlightStatus="0"
			 end if
     else
     EditFlightStatus="0"
     end if
     
      if Request.Form("ExportDataToExcel")<>nil then
			 if Request.Form("ExportDataToExcel")="on" then
				ExportDataToExcel="1"
			 else
				ExportDataToExcel="0"
			 end if
     else
     ExportDataToExcel="0"
     end if
     
        if Request.Form("UndoPNR")<>nil then
			 if Request.Form("UndoPNR")="on" then
				UndoPNR="1"
			 else
				UndoPNR="0"
			 end if
     else
     UndoPNR="0"
     end if
     
           if Request.Form("IsSendFromTL")<>nil then
			 if Request.Form("IsSendFromTL")="on" then
				IsSendFromTL="1"
			 else
				IsSendFromTL="0"
			 end if
     else
     IsSendFromTL="0"
     end if
     
     
     
     
     
      if Request.Form("EditMainData")<>nil then
			 if Request.Form("EditMainData")="on" then
				EditMainData="1"
			 else
				EditMainData="0"
			 end if
     else
     EditMainData="0"
     end if
     
         if Request.Form("InsertRowMainData")<>nil then
			 if Request.Form("InsertRowMainData")="on" then
				InsertRowMainData="1"
			 else
				InsertRowMainData="0"
			 end if
     else
     InsertRowMainData="0"
     end if
     
       if Request.Form("DeleteRowMainData")<>nil then
			 if Request.Form("DeleteRowMainData")="on" then
				DeleteRowMainData="1"
			 else
				DeleteRowMainData="0"
			 end if
     else
     DeleteRowMainData="0"
     end if
     
     
     
 'response.Write "Edit_IsNewLetters="& Request.Form("Edit_IsNewLetters")
 'response.end
  if Request.Form("Edit_IsNewLetters")<>nil then
			 if Request.Form("Edit_IsNewLetters")="on" then
				Edit_IsNewLetters="1"
			 else
				Edit_IsNewLetters="0"
			 end if
     else
     Edit_IsNewLetters="0"
     end if
     
     if Request.Form("IsTypingClient")<>nil then
			 if Request.Form("IsTypingClient")="on" then
				IsTypingClient="1"
			 else
				IsTypingClient="0"
			 end if
     else
     IsTypingClient="0"
     end if
     
        if Request.Form("InsertNewUser")<>nil then
			 if Request.Form("InsertNewUser")="on" then
				InsertNewUser="1"
			 else
				InsertNewUser="0"
			 end if
     else
     InsertNewUser="0"
     end if
     
     
     
     'response.Write "IsTypingClient="&IsTypingClient
       if Request.Form("UploadUserFiles")<>nil then
			if Request.Form("UploadUserFiles")="on" then
				UploadUserFiles="1"
			else
				UploadUserFiles="0"
			end if
	 else
		 UploadUserFiles="0"
     end if
     'response.Write "UploadUserFiles="& UploadUserFiles
  '   response.end
     
      
 
 
	 
     
     if Request.Form("salesControl_Points_Edit")<>nil then
			if Request.Form("salesControl_Points_Edit")="on" then
				salesControl_Points_Edit="1"
			else
				salesControl_Points_Edit="0"
			end if
	 else
		 salesControl_Points_Edit="0"
     end if
     
     
     
     if Request.Form("Edit_DepartmentAppeal")<>nil then
     if Request.Form("Edit_DepartmentAppeal")="on" then
     Edit_DepartmentAppeal="1"
     else
     Edit_DepartmentAppeal="0"
     end if
     
  
     else
     Edit_DepartmentAppeal="0"
     end if
     
     job_pay = trim(Request.Form("job_id"))
     Dep_ID= trim(Request.Form("Dep_ID"))
     job_pay = Split(job_pay,"!")
     If IsArray(job_pay) Then
		job_id = job_pay(0)
	 Else
		job_id = "NULL"
	 End If		 
	 
     If Request.Form("chief") <> nil Then
		chief = "1"  
     Else
		chief = "0"
     End If
     If Request.Form("arch_app") <> nil Then
		arch_app = "1"  
     Else
		arch_app = "0"
     End If
     
          If Request.Form("IP_login") <> nil Then
		IP_login = "1"  
     Else
		IP_login = "0"
     End If
 
        If Request.Form("Add_Insurance") <> nil Then
		Add_Insurance = "1"  
     Else
		Add_Insurance = "0"
     End If
     If Request.Form("Add_GroupsTours") <> nil Then
		 Add_GroupsTours = "1"  
     Else
		 Add_GroupsTours = "0"
     End If
       If Request.Form("Sms_Write") <> nil Then
		 Sms_Write = "1"  
     Else
		 Sms_Write = "0"
     End If
        If Request.Form("EmailGroupSend") <> nil Then
		 EmailGroupSend = "1"  
     Else
		 EmailGroupSend = "0"
     End If
        If Request.Form("Screen_TourVisible") <> nil Then
		 Screen_TourVisible = "1"  
     Else
		 Screen_TourVisible = "0"
     End If
     
         If Request.Form("ScreenTourStatus") <> nil Then
		 ScreenTourStatus = "1"  
     Else
		 ScreenTourStatus = "0"
     End If
     
           If Request.Form("ScreenSendMail") <> nil Then
		 ScreenSendMail = "1"  
     Else
		 ScreenSendMail = "0"
     End If
        
           If Request.Form("ScreenTourOrder") <> nil Then
		 ScreenTourOrder = "1"  
     Else
		 ScreenTourOrder = "0"
     End If
    ' response.Write ScreenTourOrder
     'response.end
           If Request.Form("ScreenGeneralSendSms") <> nil Then
		 ScreenGeneralSendSms = "1"  
     Else
		 ScreenGeneralSendSms = "0"
     End If
          If Request.Form("ScreenContactSendSms") <> nil Then
		 ScreenContactSendSms = "1"  
     Else
		 ScreenContactSendSms = "0"
     End If
          If Request.Form("OperationScreenView") <> nil Then
		 OperationScreenView = "1"  
     Else
		 OperationScreenView = "0"
     End If
            If Request.Form("ReportMaakavRishumView") <> nil Then
		 ReportMaakavRishumView = "1"  
     Else
		 ReportMaakavRishumView = "0"
     End If
             If Request.Form("DashBoardView") <> nil Then
		 DashBoardView = "1"  
     Else
		 DashBoardView = "0"
     End If
              If Request.Form("UpdateFieldUserBlocked") <> nil Then
		 UpdateFieldUserBlocked = "1"  
     Else
		 UpdateFieldUserBlocked = "0"
     End If
     
     
     
         If Request.Form("ReportCloseProcView") <> nil Then
		 ReportCloseProcView = "1"  
     Else
		 ReportCloseProcView = "0"
     End If
     
     
     
    
     Order_Price = nFix(Request.Form("Order_Price"))
     ImportanceId= nFix(Request.Form("ImportanceId"))
     if IsNumeric(ImportanceId) then
     else
     ImportanceId=0
     end if
     COMPANIES = 0 : SURVEYS = 0 : EMAILS = 0 : WORK_PRICING = 0 : PROPERTIES = 0 : TASKS = 0 : CASH_FLOW = 0
	 if USER_ID=nil or USER_ID="" then 'new record in DataBase
		    sqlStr = "select LOGINNAME from USERS where LOGINNAME = '" & sFix(LOGINNAME) & "'"
		    set rs_LOGINNAME = con.GetRecordSet(sqlStr)
		    if  not rs_LOGINNAME.EOF then
		        errLogin=true
		    end if
		    set rs_LOGINNAME = nothing
		    sqlStr = "select EMAIL from USERS where EMAIL='" & email & "' AND ORGANIZATION_ID = " & OrgID
		    set rs_EMAIL = con.GetRecordSet(sqlStr)
		    if  not rs_EMAIL.EOF then
		        errEmail=true
		    end if
		    set rs_EMAIL = Nothing
		    'Response.Write ("<BR>errLogin="& errLogin)
		    'Response.Write ("<BR>errEmail="& errEmail)
		    'Response.End
		    if errLogin=false AND errEmail=false  then
				sqlStr = "SET NOCOUNT ON; SET DATEFORMAT DMY; Insert Into USERS (FIRSTNAME,LASTNAME,FIRSTNAME_ENG,LASTNAME_ENG,LOGINNAME,PASSWORD,ACTIVE,"&_
				"ORGANIZATION_ID,job_id,Department_Id,ImportanceId,TELEPHONE,MOBILE,FAX,EMAIL,Month_Min_Order,WORK_PRICING,SURVEYS,EMAILS,"&_
				"COMPANIES,TASKS,CASH_FLOW,CHIEF,Archive_appeal,IP_login,Add_Insurance, Add_GroupsTours,Sms_Write,EmailGroupSend,User_Screen_TourVisible,User_ScreenTourStatus,User_ScreenSendMail,User_ScreenTourOrder,User_SendSms_GeneralScreen,User_SendSms_ContactScreen,EDIT_APPEAL,Order_Price,User_OperationScreenView,ReportCloseProcView,ReportMaakavRishumView,DashBoardView,UpdateFieldUserBlocked,Edit_DepartmentAppeal,salesControl_Edit,salesControl_Points_Edit,EditFlyCompanies,EditAirports,EditApprovalStatus,EditCurrency,EditFlightStatus,ExportDataToExcel,EditMainData,InsertRowMainData,DeleteRowMainData,Edit_IsNewLetters,IsTypingClient,InsertNewUser,EMAIL_PRIVATE,MOBILE_PRIVATE,BIRTHDAY,PermissionsUploadFiles) " &_
				" values ('" & FIRSTNAME &"','"& LASTNAME &"','"& FIRSTNAME_ENG &"','"& LASTNAME_ENG &"','"& LOGINNAME &"','"& PASSWORD &"',1,"& OrgID &","&_
				job_id &"," & dep_Id &"," & ImportanceId &",'"& phone &"','"& mobile &"','"& fax &"','"& email &"','" & Month_Min_Order & "','" &_
				WORK_PRICING& "','" &SURVEYS& "','" &EMAILS& "','" &COMPANIES& "','" & TASKS & "','" &_
				CASH_FLOW & "','" & CHIEF  & "','" & arch_app & "'," & IP_login  & "," & Add_Insurance  & "," & Add_GroupsTours &","& Sms_Write &","& EmailGroupSend & ",'" & Screen_TourVisible & "','" & ScreenTourStatus & "','" & ScreenSendMail & "','" & ScreenTourOrder & "','" & ScreenGeneralSendSms & "','" & ScreenContactSendSms & "','" & EDIT_APPEAL & "','" & Order_Price & "','" & ReportMaakavRishumView  & "','" & ReportCloseProcView  & "','" & OperationScreenView &"','" & DashBoardView &"','" & UpdateFieldUserBlocked &"','" &Edit_DepartmentAppeal & "','" & salesControl_Edit &"' "& ",'" & salesControl_Points_Edit &"','"& EditFlyCompanies &"','"& EditAirports &"','"& ExportDataToExcel  &"','"& EditMainData &"','"& InsertRowMainData &"','"& DeleteRowMainData &"','"& EditApprovalStatus &"','"& EditCurrency &"','"& EditFlightStatus &"','"& Edit_IsNewLetters &"','"& IsTypingClient &"','"& InsertNewUser &"','"& emailPrivate  &"','"& mobilePrivate &"','" & birthday &"','"& UploadUserFiles &"'); SELECT @@IDENTITY AS NewID"
			'response.Write  "sqlStr="& sqlStr
			'response.end
				set rs_tmp = con.getRecordSet(sqlStr)
					User_ID = rs_tmp.Fields("NewID").value
				set rs_tmp = Nothing	
				
				'--insert into changes table
				sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
				" SELECT 'חשבון משתמש', IsNULL(U.FIRSTNAME, '') + ' ' + IsNULL(LASTNAME, ''), [USER_ID], 'הוספה', getDate(), " & UserID & _
				" FROM dbo.USERS U WHERE ([USER_ID] = " & User_ID & ")"
				con.executeQuery(sqlstr)								
										  
			 elseif errLogin=true then
				 LOGINNAME = "" %>
			     <script language="javascript" type="text/javascript">
			     <!--
				   alert('שם משתמש זה בשימוש, בחר שם משתמש אחר')
				   history.back();
			      //-->
			     </SCRIPT>			 			 
			 <%Response.End
			   elseif errEmail=true then
				  email=""%>
			      <script language="javascript" type="text/javascript">
			      <!--
				     alert('כתובת דואר אלקטרוני שציינת'+'\n'+'רשומה כבר במערכת') ;
				     history.back();
			      //-->
			      </SCRIPT>			 			 				
			  <%Response.End
			  end if			  
			 
		    
	 else			   
				 sqlStr = "select LOGINNAME from USERS where USER_ID <> "& USER_ID &" and LOGINNAME = '" & sFix(LOGINNAME) & "'"
				 set rs_LOGINNAME = con.GetRecordSet(sqlStr)
				 if not rs_LOGINNAME.EOF then
		             errLogin=true
		         end if
		         set rs_LOGINNAME = nothing
		    
				 sqlStr = "select EMAIL from USERS where USER_ID <> "& USER_ID &" and EMAIL='" & email & "' AND ORGANIZATION_ID = " & OrgID
		         set rs_EMAIL = con.GetRecordSet(sqlStr)
		         if not rs_EMAIL.EOF then
		             errEmail=true
		         end if
		         set rs_EMAIL = Nothing
		         
               Order_Price = nFix(Request.Form("Order_Price"))
				 ImportanceId= nFix(Request.Form("ImportanceId"))
				 if IsNumeric(ImportanceId) then
				 else
				 ImportanceId=0
				 end if
			'	response.Write "BIRTHDAY='"&BIRTHDAY  &":"
		
				'	response.Write "BIRTHDAY='"&BIRTHDAY  &":"
				 if errLogin=false AND errEmail=false  then
					sqlStr = " SET DATEFORMAT DMY;UPDATE USERS set FIRSTNAME='" & FIRSTNAME &"',LASTNAME='"& LASTNAME &_
					"', FIRSTNAME_ENG='"& FIRSTNAME_ENG &"',LASTNAME_ENG='"& LASTNAME_ENG &_
					"', LOGINNAME='"& LOGINNAME &"', PASSWORD = '"& PASSWORD &"',job_id="& job_id &",Department_Id="& dep_Id &",ImportanceId="& ImportanceId &_
					", TELEPHONE='"& phone &"',MOBILE='"& mobile &"',FAX='"& fax &"',EMAIL='"& email &_
					"', WORK_PRICING='" & WORK_PRICING & "',SURVEYS='" & SURVEYS & "',EMAILS='" & EMAILS &_
					"', COMPANIES='" & COMPANIES & "', TASKS='" & TASKS & "', CASH_FLOW='" & CASH_FLOW &_
					"', Month_Min_Order='" & Month_Min_Order & "', CHIEF = '" & CHIEF & "', Archive_appeal = '" & arch_app & "', IP_login = " & IP_login  & ", Add_Insurance = " & Add_Insurance & ",  Add_GroupsTours = " &  Add_GroupsTours & ",  Sms_Write = " &  Sms_Write & ", EmailGroupSend="& EmailGroupSend  & ", User_Screen_TourVisible="& Screen_TourVisible  & ", User_ScreenTourStatus="& ScreenTourStatus  & ", User_ScreenSendMail="& ScreenSendMail & ", User_ScreenTourOrder=" & ScreenTourOrder & ", User_SendSms_GeneralScreen="& ScreenGeneralSendSms& ", User_SendSms_ContactScreen="& ScreenContactSendSms & ", EDIT_APPEAL = '" & EDIT_APPEAL &_
					"', Order_Price = '" & Order_Price & "',User_OperationScreenView='"& OperationScreenView  & "',ReportCloseProcView='"& ReportCloseProcView & "',ReportMaakavRishumView='"& ReportMaakavRishumView & "',DashBoardView='"& DashBoardView &"',UpdateFieldUserBlocked='" & UpdateFieldUserBlocked &"',Edit_DepartmentAppeal='"& Edit_DepartmentAppeal &"',salesControl_Edit='"& salesControl_Edit &"',salesControl_Points_Edit='"& salesControl_Points_Edit & _
					"', EditFlyCompanies='" & EditFlyCompanies & "', EditAirports='" & EditAirports & "', EditApprovalStatus='" & EditApprovalStatus &_
					"', UndoPNR='" & UndoPNR &_
					"', IsSendFromTL='" & IsSendFromTL &_
    		        "', EditCurrency='" & EditCurrency & "', EditFlightStatus='" & EditFlightStatus &"',ExportDataToExcel='"& ExportDataToExcel  &"',EditMainData='"& EditMainData  &"',InsertRowMainData='"& InsertRowMainData  &"',DeleteRowMainData='"& DeleteRowMainData  & "', Edit_IsNewLetters='" & Edit_IsNewLetters & "',IsTypingClient='" &IsTypingClient  &"',InsertNewUser='"& InsertNewUser &"',MOBILE_PRIVATE='"& mobilePrivate  &"',EMAIL_PRIVATE='"& EMAILPrivate &"',BIRTHDAY="& BIRTHDAY &",UploadUserFiles='" &UploadUserFiles &"' WHERE USER_ID=" & USER_ID
					 
'	Response.Write sqlStr
'	Response.End
					con.GetRecordSet (sqlStr)	
					
				'--insert into changes table
				sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
				" SELECT 'חשבון משתמש', IsNULL(U.FIRSTNAME, '') + ' ' + IsNULL(LASTNAME, ''), [USER_ID], 'עדכון', getDate(), " & UserID & _
				" FROM dbo.USERS U WHERE ([USER_ID] = " & User_ID & ")"
			
			'response.Write sqlstr
			'response.end
				con.executeQuery(sqlstr)					
									     
				 elseif errLogin=true then
					LOGINNAME = ""%>
				    <script language="javascript" type="text/javascript">
				     <!--
					   alert('שם משתמש זה בשימוש, בחר שם משתמש אחר')
					   history.back();
				     //-->
				   </SCRIPT>			 			 
				 <%Response.End
				 elseif errEmail=true then
				   email=""%>
			       <script language="javascript" type="text/javascript">
			       <!--
				     alert('כתובת דואר אלקטרוני שציינת'+'\n'+'רשומה כבר במערכת') 
				     history.back();
			       //-->
			       </SCRIPT>			 			 				
			    <%Response.End
			    end if
			    			    
    end if
	sqlstr = "Delete FROM bar_users WHERE organization_id = " & OrgID & " AND USER_ID = " & USER_ID
	'Response.Write sqlstr
	'Response.End
	con.executeQuery(sqlstr)
			
	sqlstr="Select bar_id from bars Order by PARENT_ID, bar_Order"
	set rs_bars = con.getRecordSet(sqlstr)
	While not rs_bars.eof
		barID = trim(rs_bars(0))
		barVisible = Request.Form("is_visible"&barID)
		If trim(barVisible) = "on" Then
			barVisible = 1
		Else
			barVisible = 0
		End If	
		If barID = "1" Then	
			COMPANIES = barVisible		
		ElseIf barID = "2" Then
			SURVEYS = barVisible
		ElseIf barID = "3" Then		
			EMAILS = barVisible
		ElseIf barID = "4" Then	
			WORK_PRICING = barVisible
		ElseIf barID = "5" Then	
			PROPERTIES = barVisible
		ElseIf barID = "6" Then		
			TASKS = barVisible	
		ElseIf barID = "29" Then		
			CASH_FLOW = barVisible					
		End If			
		sqlstr = "Insert Into bar_users values (" & barID & "," & OrgID & "," & USER_ID & ",'" & barVisible & "')"
		con.executeQuery(sqlstr)
	rs_bars.moveNext
	Wend
	set rs_bars = Nothing   
	
	sqlStr = "UPDATE USERS SET WORK_PRICING='"&WORK_PRICING&"',SURVEYS='"&SURVEYS&"',EMAILS='"&EMAILS&_
	"', COMPANIES='"&COMPANIES&"', TASKS='"&TASKS&"', CASH_FLOW='"&CASH_FLOW&"', EDIT_APPEAL = '" & EDIT_APPEAL &_
	"' WHERE USER_ID="&USER_ID
	'Response.Write sqlStr
	'Response.End
	con.GetRecordSet (sqlStr)	
	
	'<--------------------------------------- שיוך לקבוצות ------------------------------------------------------>
	sqlstr="Delete FROM Users_to_Groups WHERE USER_ID="&USER_ID
	con.executeQuery(sqlstr)
	
	groups = trim(Request.Form("group_id")) & ","
	groups = Split(groups, ",")
	If IsArray(groups) And Ubound(groups) > 0 Then
		For i=0 To Ubound(groups)
			If trim(groups(i)) <> "" And IsNumeric(groups(i)) = true Then
				sqlstr="Insert Into Users_to_Groups (USER_ID,group_id) values ("&USER_ID&","&groups(i)&" )"
				con.executeQuery(sqlstr)
			End If 
		Next
	End If
	'<--------------------------------------- שיוך לקבוצות ------------------------------------------------------>
	
	'<--------------------------------------- אחראי בקבוצות ------------------------------------------------------>
	sqlstr="Delete FROM Responsibles_to_Groups WHERE Responsible_ID="&USER_ID
	con.executeQuery(sqlstr)
	
	r_groups = trim(Request.Form("R_GROUP_ID")) & ","
	r_groups = Split(r_groups, ",")
	If IsArray(r_groups) And Ubound(r_groups) > 0 Then
		For i=0 To Ubound(r_groups)
			If trim(r_groups(i)) <> "" And IsNumeric(r_groups(i)) = true Then
				sqlstr="Insert Into Responsibles_to_Groups (Responsible_ID,GROUP_ID) values ("&USER_ID&","&r_groups(i)&" )"
				con.executeQuery(sqlstr)
			End If 
		Next
	End If
	'<--------------------------------------- אחראי בקבוצות ------------------------------------------------------>
	'<----------------------------------------הוספת הראשת טפסים לקבוצות שנוספו ----------------------------------->
	If Request.Form("group_id") <> nil Then
	sqlstr = "Select DISTINCT Product_ID, Group_ID From Users_To_Products WHERE Organization_ID = " & OrgID & " And Group_ID IN (" & trim(Request.Form("group_id")) & ")"
	set rs_products = con.getRecordSet(sqlstr)
	while not rs_products.eof
	    sqlCheck = "Select Top 1 Product_ID From Users_To_Products WHERE User_ID = " & USER_ID &_
	    " And Product_ID = " & rs_products(0) & " And Group_ID = " & rs_products(1)
	    set rs_check = con.getRecordSet(sqlCheck)
	    If rs_check.eof Then
	    sqlInsert = "Insert Into Users_To_Products values (" & USER_ID & "," & rs_products(1) & "," & OrgID & "," & rs_products(0) & ")"
	    con.executeQuery(sqlInsert)
	    End If
	    set rs_check = Nothing
		rs_products.moveNext
	wend
	set rs_products = Nothing
	End If
	'<----------------------------------------הוספת הראשת טפסים לקבוצות שנוספו ----------------------------------->
	'<----------------------------------------למחוק הרשאה לטפסים מקבוצות שנמחקו ---------------------------------->
	sqlDelete = "Delete FROM Users_To_Products WHERE  User_ID = " & USER_ID &_
	" And Group_ID Not IN (Select Group_ID FROM Users_to_Groups WHERE USER_ID="&USER_ID&")"
	con.executeQuery(sqlDelete)
	'<----------------------------------------למחוק הרשאה לטפסים מקבוצות שנמחקו ---------------------------------->	
	
    Response.Redirect "addWorker.asp?User_ID=" & USER_ID       
end if
%>
<%
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 39 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
	arrTitlesD = rstitle.getRows()		
	redim arrTitles(Ubound(arrTitlesD,2)+1)
	For i=0 To Ubound(arrTitlesD,2)		
		arrTitles(arrTitlesD(0,i)) = arrTitlesD(1,i)		
	Next
	End If
	set rstitle = Nothing
	  
	sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	set rsbuttons = con.getRecordSet(sqlstr)
	If not rsbuttons.eof Then
	arrButtons = ";" & rsbuttons.getString(,,",",";")		
	arrButtons = Split(arrButtons, ";")
	End If
	set rsbuttons=nothing	 	  
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">

	<script language="javascript" type="text/javascript" src="../../../Scripts/CalendarPopupH.js"></script>
	<script language="javascript" type="text/javascript" >
<!--
function callCalendar(pf,pid)
	{

	cal1xx.select(pf,pid,'dd/MM/yyyy')
	}
	function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=157pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}
		//-->
	</SCRIPT>
</head>
<script language="javascript" type="text/javascript">
<!--
function CheckFields()
{
	if ( document.frmMain.FIRSTNAME.value=='' ||
	     document.frmMain.LOGINNAME.value=='' || 
	     document.frmMain.PASSWORD.value==''  ||
	     document.frmMain.JOB_ID.value==''  ||	     
	     document.frmMain.email.value==''
	     )		   
		{
			<%
			If trim(lang_id) = "1" Then
				str_alert = "'*' אנא מלאו כל השדות הצוינות בסימן"
			Else
				str_alert = "Please insert all requested fields !!"
			End If   
			%>			
			window.alert("<%=str_alert%>");						
			return false;				
		}
	else if (!checkEmail(document.all("email").value) && document.all("email").value != "")
		{
           <%
				If trim(lang_id) = "1" Then
					str_alert = "כתובת דואר אלקטרוני לא חוקית"
				Else
					str_alert = "The email address is not valid!"
				End If	
			%>
				window.alert("<%=str_alert%>");
				document.all("email").focus();
				return false;
		}	
		
		
			else if (!checkPassword(document.all("PASSWORD").value) && document.all("PASSWORD").value != "")
		{
           <%
				If trim(lang_id) = "1" Then
					str_alert = "סיסמה לא חוקית\n על כל סיסמה להיות בנויה מ 8 תוים המכילים בתוכם לפחות אות גדולה אחת, לפחות מספר אחד ולפחות תו אחד (!, @, #, וכו)"
				Else
					str_alert = "The password address is not valid!"
				End If	
			%>
				window.alert("<%=str_alert%>");
				document.all("PASSWORD").focus();
				return false;
		}	
		
	else
		{
		$("input").removeAttr("disabled");
		document.frmMain.submit();
		return true;
		}		
}

	function GetNumbers ()
	{
		var ch=event.keyCode;
		event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
	} 

	function check_all_bars(objChk,parentID)
	{
		input_arr = document.getElementsByTagName("input");	
		
		for(i=0;i<input_arr.length;i++)	{
		
			
			if(input_arr[i].type == "checkbox")
			{
				currparentId = "";
				objValue = new String(input_arr[i].id);			
				value_arr = objValue.split("!");
				currparentId =  value_arr[1];
				if(currparentId == parentID)
				{
					//input_arr(i).disabled = objChk.checked;
					input_arr[i].checked = objChk.checked;
				}	
			}	
		}
		return true;
	}
	

function checkPassword(pass)
{
 var str = pass;
    if (str.length != 8) {
        alert("על כל סיסמה להיות בנויה מ 8 תוים");
        return false;

}
var validChars=/^(?=.*?[A-Z])(?=.*?[a-z])(?=.*?[0-9])(?=.*?[^\w\s]).{8,}$/
return (validChars.test(pass))
}
function checkEmail(addr)
{
	addr = addr.replace(/^\s*|\s*$/g,"");
	if (addr == '') {
	return false;
	}
	var invalidChars = '\/\'\\ ";:?!()[]\{\}^|';
	for (i=0; i<invalidChars.length; i++) {
	if (addr.indexOf(invalidChars.charAt(i),0) > -1) {
		return false;
	}
	}
	for (i=0; i<addr.length; i++) {
	if (addr.charCodeAt(i)>127) {     
		return false;
	}
	}

	var atPos = addr.indexOf('@',0);
	if (atPos == -1) {
	return false;
	}
	if (atPos == 0) {
	return false;
	}
	if (addr.indexOf('@', atPos + 1) > - 1) {
	return false;
	}
	if (addr.indexOf('.', atPos) == -1) {
	return false;
	}
	if (addr.indexOf('@.',0) != -1) {
	return false;
	}
	if (addr.indexOf('.@',0) != -1){
	return false;
	}
	if (addr.indexOf('..',0) != -1) {
	return false;
	}
	var suffix = addr.substring(addr.lastIndexOf('.')+1);
	if (suffix.length != 2 && suffix != 'com' && suffix != 'net' && suffix != 'org' && suffix != 'edu' && suffix != 'int' && suffix != 'mil' && suffix != 'gov' & suffix != 'arpa' && suffix != 'biz' && suffix != 'aero' && suffix != 'name' && suffix != 'coop' && suffix != 'info' && suffix != 'pro' && suffix != 'museum') {
	return false;
	}
return true;
}
//-->
</script>  
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
		<script type="text/javascript" src="//code.jquery.com/jquery-1.12.4.js"></script>
<script>
$("frmMain").submit(function() {
alert("frrr")
    $("input").removeAttr("disabled");
});

</script>

<body >
<!--#include file="../../logo_top.asp"-->
<%numOftab = 4%>
<%numOfLink = 5%>
<%If trim(wizard_id) = "" Then%>
<!--#include file="../../top_in.asp"-->
<%End If%>
<table border="0"  bgcolor="#E6E6E6" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title">&nbsp;&nbsp;</td></tr>  		       	
   </table></td></tr>            
<%
If Len(USER_ID) > 0 Then
	sqlStr = "SELECT FIRSTNAME,LASTNAME,LOGINNAME,PASSWORD,OFFICE,TELEPHONE,MOBILE,FAX,EMAIL,Month_Min_Order,chief, "&_
	" job_id,WORK_PRICING,SURVEYS,EMAILS,COMPANIES,TASKS,CASH_FLOW,CHIEF, ISNULL(EDIT_APPEAL,0),"&_
	" isNULL(Voice_Operator,0),IsNUll(Order_Price,0),Add_Insurance, Add_GroupsTours ,Sms_Write,EmailGroupSend,IP_login,IsNULL(User_Screen_TourVisible,0),IsNULL(User_ScreenTourStatus,0),IsNULL(User_ScreenSendMail,0),ImportanceId,IsNULL(User_SendSms_GeneralScreen,0),IsNULL(User_SendSms_ContactScreen,0),IsNULL(User_ScreenTourOrder,0),IsNULL(User_OperationScreenView,0)," & _
	"IsNULL(ReportCloseProcView,0),IsNULL(ReportMaakavRishumView ,0),IsNULL(DashBoardView ,0),Department_Id, Edit_DepartmentAppeal,salesControl_Edit,salesControl_Points_Edit,EditFlyCompanies,EditAirports,EditApprovalStatus,EditCurrency,EditFlightStatus,Edit_IsNewLetters,IsNULL(ExportDataToExcel ,0) ,IsNULL(EditMainData ,0),IsNULL(InsertRowMainData ,0),IsNULL(DeleteRowMainData ,0),IsNULL(Archive_appeal ,0),IsNULL(UndoPNR ,0) ,IsNULL(IsSendFromTL,0),FIRSTNAME_ENG,LASTNAME_ENG,IsNULL(IsTypingClient,0), " & _
	" EMAIL_PRIVATE,MOBILE_PRIVATE,BIRTHDAY,isNull(UpdateFieldUserBlocked,0),IsNULL(InsertNewUser,0),dbo.getPermissions(Department_Id,"& workerId &",90011)as CHIEFPerm,dbo.getPermissions(Department_Id,"& workerId &",90019) as InsertNewUserPerm, dbo.getPermissions(Department_Id,"& workerId &",90012) as IsTypingClientPerm, dbo.getPermissions(Department_Id,"& workerId &",90013) as UpdateFieldUserBlockedPerm , " & _
	" dbo.getPermissions(Department_Id,"& workerId &",90014) as Add_GroupsToursPerm,dbo.getPermissions(Department_Id,"& workerId &",90015) as Sms_WritePerm," &_
	" dbo.getPermissions(Department_Id,"& workerId &",90016) as Edit_IsNewLettersPerm,dbo.getPermissions(Department_Id,"& workerId &",90017) as EmailGroupSendPerm, "& _
	" dbo.getPermissions(Department_Id,"& workerId &",90018) as DashBoardViewPerm ," &_
	" dbo.getPermissions(Department_Id,"& workerId &",90021) as Edit_DepartmentAppealPerm,dbo.getPermissions(Department_Id,"& workerId &",90022) as salesControl_Points_EditPerm," &_ 
	" dbo.getPermissions(Department_Id,"& workerId &",90023) as salesControl_EditPerm, " &_
	" dbo.getPermissions(Department_Id,"& workerId &",90031) as Screen_TourVisiblePerm,dbo.getPermissions(Department_Id,"& workerId &",90032) as ScreenTourStatusPerm, "& _
    " dbo.getPermissions(Department_Id,"& workerId &",90033) as ScreenSendMailPerm,dbo.getPermissions(Department_Id,"& workerId &",90034) as ScreenTourOrderPerm, "& _
	" dbo.getPermissions(Department_Id,"& workerId &",90041) as ScreenGeneralSendSmsPerm,dbo.getPermissions(Department_Id,"& workerId &",90042) as ScreenContactSendSmsPerm, "& _
	" dbo.getPermissions(Department_Id,"& workerId &",90051) as OperationScreenViewPerm, "& _
	" dbo.getPermissions(Department_Id,"& workerId &",90061) as ReportMaakavRishumViewPerm,dbo.getPermissions(Department_Id,"& workerId &",90062) as ReportCloseProcViewPerm, "& _ 
	" dbo.getPermissions(Department_Id,"& workerId &",90071) as EditFlyCompaniesPerm,"& _
	" dbo.getPermissions(Department_Id,"& workerId &",90072) as EditAirportsPerm,dbo.getPermissions(Department_Id,"& workerId &",90073) as EditApprovalStatusPerm,dbo.getPermissions(Department_Id,"& workerId &",90074) as EditCurrencyPerm,"& _
	" dbo.getPermissions(Department_Id,"& workerId &",90075) as EditFlightStatusPerm,dbo.getPermissions(Department_Id,"& workerId &",90076) as IsSendFromTLPerm,dbo.getPermissions(Department_Id,"& workerId &",90077) as UndoPNRPerm,"& _
	" dbo.getPermissions(Department_Id,"& workerId &",90078) as DeleteRowMainDataPerm,dbo.getPermissions(Department_Id,"& workerId &",90079) as InsertRowMainDataPerm,dbo.getPermissions(Department_Id,"& workerId &",90080) as EditMainDataPerm, "& _
	" dbo.getPermissions(Department_Id,"& workerId &",90020) as UploadUserFilesPerm ,isNULL(UploadUserFiles,0) as UploadUserFiles " & _
	" FROM dbo.USERS WHERE USER_ID=" & USER_ID  
	'Response.Write sqlStr
	'response.end
	Set rs_USERS = con.GetRecordSet(sqlStr)
	If Not rs_USERS.eof Then
		FIRSTNAME   = rs_USERS(0)
		LASTNAME    = rs_USERS(1)
		LOGINNAME   = rs_USERS(2)	
		PASSWORD_U  = rs_USERS(3)					
		office      = rs_USERS(4)
		phone       = rs_USERS(5)
		mobile      = rs_USERS(6)
		fax         = rs_USERS(7)
		email       = rs_USERS(8)
		Month_Min_Order  = rs_USERS(9)
		chief       = rs_USERS(10) 
		job_id    = rs_USERS(11)
		WORK_PRICING = rs_USERS(12)
		SURVEYS = rs_USERS(13)
		EMAILS = rs_USERS(14)
		COMPANIES = rs_USERS(15)
		TASKS = rs_USERS(16)
		CASH_FLOW = rs_USERS(17)
		CHIEF = rs_USERS(18)	
		EDIT_APPEAL = rs_USERS(19) 	
		Voice_Operator = rs_USERS(20) 	
		Order_Price = rs_USERS(21) 	
		Add_Insurance= rs_USERS(22) 	
		Add_GroupsTours=rs_USERS(23) 	
		Sms_Write=rs_USERS(24) 
		
		EmailGroupSend=rs_USERS(25) 
		IP_login=rs_USERS(26) 
		
		Screen_TourVisible=rs_USERS(27) 
		ScreenTourStatus=rs_USERS(28) 
		ScreenSendMail=rs_USERS(29) 
		ImportanceId=rs_USERS(30) 
		
		ScreenGeneralSendSms=rs_USERS(31) 
		ScreenContactSendSms=rs_USERS(32) 
		ScreenTourOrder=rs_USERS(33) 
		OperationScreenView=rs_USERS(34) 
		ReportCloseProcView=rs_USERS(35) 
		ReportMaakavRishumView=rs_USERS(36) 
		DashBoardView=rs_USERS(37) 
		
		dep_id=rs_USERS(38) 
		Edit_DepartmentAppeal  =rs_USERS(39)
		salesControl_Edit=rs_USERS(40)
		salesControl_Points_Edit=rs_USERS(41)
		EditFlyCompanies=rs_USERS(42)
		EditAirports=rs_USERS(43)
		EditApprovalStatus=rs_USERS(44)
		
		EditCurrency=rs_USERS(45)
		EditFlightStatus=rs_USERS(46)
		Edit_IsNewLetters=rs_USERS(47)
		ExportDataToExcel=rs_USERS(48)
		EditMainData=rs_USERS(49)
		
		InsertRowMainData=rs_USERS(50)
		DeleteRowMainData=rs_USERS(51)
		arch_app=rs_USERS(52)
		
		UndoPNR=rs_USERS(53)
		IsSendFromTL=rs_USERS(54)
		FIRSTNAME_ENG=rs_USERS(55)
		LASTNAME_ENG=rs_USERS(56)
		IsTypingClient=rs_USERS(57)
		
		EMAILPRIVATE=rs_USERS(58)
		MOBILEPRIVATE=rs_USERS(59)
		BIRTHDAY=rs_USERS(60)
		
		UpdateFieldUserBlocked=rs_USERS(61)
		InsertNewUser=rs_USERS(62)
		CHIEFPerm=rs_USERS(63)
		
		InsertNewUserPerm=rs_USERS(64)
		IsTypingClientPerm=rs_USERS(65)
		UpdateFieldUserBlockedPerm=rs_USERS(66)
		Add_GroupsToursPerm=rs_USERS(67)  
		
		Sms_WritePerm=rs_USERS(68)  
		Edit_IsNewLettersPerm=rs_USERS(69)  
		EmailGroupSendPerm=rs_USERS(70)  
		DashBoardViewPerm=rs_USERS(71)  
		Edit_DepartmentAppealPerm=rs_USERS(72)  
		salesControl_Points_EditPerm=rs_USERS(73)  
		salesControl_EditPerm=rs_USERS(74) 
		Screen_TourVisiblePerm=rs_USERS(75)  
		ScreenTourStatusPerm=rs_USERS(76)  
		ScreenSendMailPerm=rs_USERS(77)  
		ScreenTourOrderPerm=rs_USERS(78)  
		ScreenGeneralSendSmsPerm=rs_USERS(79)   
		ScreenContactSendSmsPerm=rs_USERS(80)   
		OperationScreenViewPerm=rs_USERS(81)   
		ReportMaakavRishumViewPerm=rs_USERS(82)   
		ReportCloseProcViewPerm=rs_USERS(83)   
		EditFlyCompaniesPerm=rs_USERS(84)   
		EditAirportsPerm=rs_USERS(85)  
		EditApprovalStatusPerm=rs_USERS(86)  
		EditCurrencyPerm=rs_USERS(87)  
		EditFlightStatusPerm=rs_USERS(88)  
		IsSendFromTLPerm=rs_USERS(89)  
		UndoPNRPerm=rs_USERS(90)  
		DeleteRowMainDataPerm=rs_USERS(91)  
		InsertRowMainDataPerm=rs_USERS(92)  
		EditMainDataPerm=rs_USERS(93)  
        UploadUserFilesPerm=rs_USERS(94)  
     '   UploadUserFiles=rs_USERS(95)  

	end if
	set rs_USERS = nothing	
	
	Groups=""		
	sqlstr="Select Group_Name From Users_to_Groups_view Where User_id = " & USER_ID & " Order By Group_Name"
	set rssub = con.getRecordSet(sqlstr)		   
	If not rssub.eof Then
		Groups = rssub.getString(,,",",",")		
	Else	
		Groups = ""
	End If		    
	set rssub=Nothing
	If Len(Groups) > 0 Then
		Groups = Left(Groups,(Len(Groups)-1))
	End If	
	
	user_groups=""		
	sqlstr="Select Group_ID From Users_to_Groups_view Where User_id = " & USER_ID & " Order By Group_id"
	set rssub = con.getRecordSet(sqlstr)		   
	If not rssub.eof Then
		user_groups = rssub.getString(,,",",",")		
	Else	
		user_groups = ""
	End If		    
	set rssub=Nothing
	If Len(user_groups) > 0 Then
		user_groups = Left(user_groups,(Len(user_groups)-1))
	End If	
	
	ResInGroups=""		
	sqlstr="Select Group_Name From Responsibles_to_Groups_view Where Responsible_id = " & USER_ID & " Order By Group_Name"
	set rssub = con.getRecordSet(sqlstr)		   
	If not rssub.eof Then
		ResInGroups = rssub.getString(,,",",",")	
	Else
		ResInGroups=""
	End If		    
	set rssub=Nothing
	If Len(ResInGroups) > 0 Then
		ResInGroups = Left(ResInGroups,(Len(ResInGroups)-1))
	End If		
	
	ResInGroupsID=""		
	sqlstr="Select Group_ID From Responsibles_to_Groups_view Where Responsible_id = " & USER_ID & " Order By Group_ID"
	set rssub = con.getRecordSet(sqlstr)		   
	If not rssub.eof Then
		ResInGroupsID = rssub.getString(,,",",",")	
	Else
		ResInGroupsID=""
	End If		    
	set rssub=Nothing
	If Len(ResInGroupsID) > 0 Then
		ResInGroupsID = Left(ResInGroupsID,(Len(ResInGroupsID)-1))
	End If
	End If
	
	If trim(Request.QueryString("JobId")) <> "" Then	
		job_id = trim(Request.QueryString("JobId"))
	End If
	If trim(Request.QueryString("ImportanceId")) <> "" Then	
		ImportanceId = trim(Request.QueryString("ImportanceId"))
	End If
	 %>
<tr><td width="100%"> 
<FORM name="frmMain" ACTION="addWorker.asp?USER_ID=<%=USER_ID%>" METHOD="post" onSubmit="return CheckFields()" ID="frmMain">
<table align=center border="0" width="100%" cellpadding="2" cellspacing="0">
<tr><td colspan=2 align=left width=100%><table border=0 cellpadding=0 cellspacing=0  align=left ID="Table1" width=100%>
<tr>
<td  width=10%><%if PermisionAdmin=1  then 'workerId=893 only itamar nir%><input type=button class="but_menu" style="width:190px" onclick="alert('לפני מעבר למסך מתן הרשאות בדוק\n אם סימנתה את מודול הגדרות ובתוכו חשבונות עובדים');window.open('CreatePermisions.aspx?Id=<%=USER_ID%>', 'CreatePermisions', 'width=1600,height=800,scrollbars=yes,toolbar=no' );" value="ניהול הרשאה למתן הרשאות"><%end if%></td>
<td width=1%></td>
<td align=left> <%IF UploadUserFiles=1 THEN%>
<input type=button class="but_menu" id="ufiles" name="ufiles" value="חוזי עבודה" onclick="window.open('UploadUserFiles.aspx?Id=<%=USER_ID%>', 'UploadUserFiles', 'width=1000,height=800,scrollbars=yes,toolbar=no' );"><%end if%> </td>
<td width=39% align="right"><input type=button class="but_menu" style="width:90px" onclick="document.location.href='default.asp';" value="<%=arrButtons(2)%>" id="Button3" name=Button2></td>
<td width=150 nowrap></td>
<td width=50% align=left><input type=button class="but_menu" style="width:90px" onclick="javascript:CheckFields(); return false;" value="<%=arrButtons(1)%>" id="Button4" name=Button1></td></tr>
<tr><td height=5 nowrap colspan=6></td></tr>
<tr><td height=1 nowrap colspan=6 bgcolor=#808080></td></tr>
<tr><td height=10 nowrap colspan=6></td></tr>
</table></td></tr>
<tr><td colspan=2 align=right width =100%>
<table border=0 cellpadding=1 cellspacing=1  align=right ID="Table2">

<tr>


	<td align="right" class="title_sort"><input type=text dir="ltr" name="emailPrivate" value="<%=vfix(emailPrivate)%>"  style="width:150" maxlength=100 style="font-family:arial" ID="emailPrivate"></td>
	<td align=right  class="title_sort">&nbsp;דואר אלקטרוני פרטי<%'=arrTitles(11)%>&nbsp;</td>

	<td align="right" class="title_sort"><input type=text dir="ltr" name="phone" value="<%=vfix(phone)%>"  style="width:150" maxlength=20 style="font-family:arial" ID="phone"></td>
	<td align=right  class="title_sort">&nbsp;<!--טלפון--><%=arrTitles(8)%>&nbsp;</td>


			<td align="right" width=12% class="title_sort">
	<select dir="<%=dir_obj_var%>" name="JOB_ID" class="norm" style="width:150px" ID="JOB_ID">
	<option value="0" ><%=arrTitles(15)%></option>
<%	sqlstr = "SELECT job_id,job_name FROM jobs WHERE (ORGANIZATION_ID = " & OrgID & ") ORDER BY job_id"
		set rs_jobs = con.getRecordSet(sqlstr)
		while not rs_jobs.eof	%>
	<option value="<%=trim(rs_jobs(0))%>" <%If trim(job_id) = trim(rs_jobs(0)) Then%> selected <%End If%>><%=rs_jobs(1)%></option>
<%	rs_jobs.moveNext
		wend
		set rs_jobs = nothing %>
	</select>
	</td>
	<td align=right class="title_sort" width=12%>&nbsp;<!--עיסוק--><%=arrTitles(7)%>&nbsp;<span style="color:#FF0000">*</span></td>


	<td align="right" class="title_sort"><input type=text dir="<%=dir_obj_var%>" name="FIRSTNAME" value="<%=vFix(FIRSTNAME)%>"  style="width:150" maxlength=20 style="font-family:arial" ID="Text5"></td>
	<td align=right class="title_sort">&nbsp;<!--שם פרטי--><%=arrTitles(3)%>&nbsp;<span style="color:#FF0000">*</span></td>

</tr>

<tr>

	<td align="right" class="title_sort"><input type=text dir="ltr" name="mobilePrivate" value="<%=vfix(mobilePrivate)%>"  style="width:150" maxlength=20 style="font-family:arial" ID="mobilePrivate"></td>
	<td align=right  class="title_sort"><!--נייד-->נייד פרטי<%'=arrTitles(9)%>&nbsp;</td>

<td align="right" class="title_sort"><input type=text dir="ltr" name="mobile" value="<%=vfix(mobile)%>"  style="width:150" maxlength=20 style="font-family:arial" ID="Text6"></td>
	<td align=right   class="title_sort"><!--נייד-->נייד עבודה<%'=arrTitles(9)%>&nbsp;</td>


	<td align="right"  width=12% class="title_sort">
	<select dir="<%=dir_obj_var%>" name="Dep_ID" class="norm" style="width:150px" ID="Dep_ID">
	<option value="0">שיוך למחלקה</option>
<%if Len(USER_ID) > 0 or  PermisionAdmin=1 then ' WorkerId=893
	sqlstr = "SELECT departmentId,departmentName FROM Departments ORDER BY PriorityLevel desc "
	else
		sqlstr = "SELECT departmentId,departmentName FROM Departments where departmentId in ("& PByDepartmentId &") ORDER BY PriorityLevel desc "
	end if
		set rs_dep = con.getRecordSet(sqlstr)
		while not rs_dep.eof	%>
	<option value="<%=trim(rs_dep(0))%>" <%If trim(dep_id) = trim(rs_dep(0)) Then%> selected <%End If%>><%=rs_dep(1)%></option>
<%	rs_dep.moveNext
		wend
		set rs_dep = nothing %>
	</select>
	</td>
	<td align="right"  class="title_sort" width=12% >&nbsp;שיוך למחלקה&nbsp;</td>
	<td align="right" class="title_sort"><input type=text dir="<%=dir_obj_var%>" name="LASTNAME" value="<%=vFix(LASTNAME)%>"  style="width:150" maxlength=20 style="font-family:arial" ID="Text3"></td>
	<td align=right  class="title_sort">&nbsp;<!--שם משפחה--><%=arrTitles(4)%>&nbsp;<span style="color:#FF0000">*</span></td>


	
</tr>
<tr>
	<td align="right" class="title_sort" width=12%><input type="text" size="3" maxlength="3" dir="ltr" style="font-family:arial" name="Order_Price" value="<%=nFix(Order_Price)%>" ID="Order_Price"></td>
	<td align="right"  width=12%  class="title_sort" dir=rtl>&nbsp;שווי כספי של הזמנה&nbsp;&nbsp;</td>
	<td align="right" class="title_sort"><input type=text dir="ltr" name="fax" value="<%=vfix(fax)%>"  style="width:150" maxlength=20 style="font-family:arial" ID="Text8"></td>
	<td align=right  class="title_sort">&nbsp;<!--פקס--><%=arrTitles(10)%>&nbsp;</td>

<td align="right" class="title_sort"><input type=text dir="ltr" name="LOGINNAME" value="<%=vFix(LOGINNAME)%>"  style="width:150" maxlength=20 style="font-family:arial" ID="LOGINNAME"></td>
	<td align=right  class="title_sort">&nbsp;<!--שם משתמש--><%=arrTitles(5)%>&nbsp;<span style="color:#FF0000">*</span></td>



<td align="right"  width=12% class="title_sort"><input type=text dir="ltr" name="FIRSTNAME_ENG" value="<%=vFix(FIRSTNAME_ENG)%>"  style="width:150" maxlength=20 style="font-family:arial" ID="Text4"></td>
	<td align=right  class="title_sort" width=12%>&nbsp; שם פרטי באנגלית<%'=arrTitles(3)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="right" class="title_sort" width=12%><input type=text dir="ltr" name="Month_Min_Order" value="<%=vfix(Month_Min_Order)%>" size="3" maxlength="3" style="font-family:arial" onkeypress="GetNumbers();" ID="Text1"></td>
	<td align="right"  width=12%  class="title_sort" >&nbsp;יעד הזמנות מינימלי&nbsp;</td>


	<td align="right" class="title_sort"><input type=text dir="ltr" name="email" value="<%=vfix(email)%>"  style="width:150" maxlength=100 style="font-family:arial" ID="Text9"></td>
	<td align="right" class="title_sort">&nbsp;<!--דואר אלקטרוני-->דואר אלקטרוני עבודה<%'=arrTitles(11)%>&nbsp;<span style="color:#FF0000">*</span></td>

	
	
		<td align="right" class="title_sort"><input type=text dir="ltr" name="PASSWORD" value="<%=vFix(PASSWORD_U)%>"  style="width:150" maxlength=20 style="font-family:arial" ID="PASSWORD"></td>
	<td align=right  class="title_sort">&nbsp;<!--סיסמה--><%=arrTitles(6)%>&nbsp;<span style="color:#FF0000">*</span></td>
	<td align="right"  width=12% class="title_sort"><input type=text dir="ltr>" name="LASTNAME_ENG" value="<%=vFix(LASTNAME_ENG)%>"  style="width:150" maxlength=20 style="font-family:arial" ID="Text2"></td>
	<td align=right  class="title_sort" width=12%>&nbsp;שם משפחה באנגלית<%'=arrTitles(4)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
<td align="right" class="title_sort"><input type="checkbox" dir="ltr" name="Add_Insurance" <%If trim(Add_Insurance) = "1" Then%> checked <%End If%> ID="Add_Insurance"></td>
	<td  align="right" class="title_sort" nowrap dir=rtl>ביטוח מול דיוויד שילד</td>
	<td align="right" class="title_sort"><input type="checkbox" dir="ltr" name="arch_app" <%If trim(arch_app) = "1" Then%> checked <%End If%> ID="arch_app"></td>
	<td align="right" class="title_sort" nowrap dir=rtl>מורשה העברה  טופס לארכיון /לטיפול חוזר</td>

	<td align="right" class="title_sort"><input type="checkbox" dir="ltr" name="IP_login" <%If trim(IP_login) = "1" Then%> checked <%End If%> ID="IP_login"></td>
	<td align="right" class="title_sort" nowrap dir=rtl>&nbsp;מורשה כניסה רק מי-IP&nbsp;&nbsp;</td>


	<td align="right" class="title_sort">
	 		<a href='' onclick='callCalendar(document.frmMain.birthday,"Asbirthday");return false;' id="Asbirthday">
			<image src='../../images/calend.gif' border=0 ></a>&nbsp;<input dir="ltr" type="text" class="passw" size=10 id="birthday" name="birthday" value="<%=birthday%>" maxlength=10 readonly>
</td>
	<td align=right  class="title_sort">&nbsp;תאריך לידה&nbsp;</td>
</tr>



<tr><tD colspan=8 height=1 width=100% style="background-color:#ffffff"></tD></tr>
<tr><td class="title" align="<%=align_var%>" colspan=8>הרשאות&nbsp;&nbsp;<span style="font-weight:500">ביצוע פעולות</span></td></tr>

<%'response.end%>
<tr><td colspan=8><table border=0 cellpadding=1 cellspacing=1  align=right ID="Table13" style="background-color:#ffffff" width=100%>
<tr>
<td class="title_sort" align=center width=14%>טיים לימיטים</td>
<td class="title_sort" align =center width=14%>מסך טפסים וסקרים -<br> דוח מעקב רישום ואחוזי סגירה</td>
<td class="title_sort" align=center width=14%>מסך אופרציה</td>
<td class="title_sort" align=center width=16% nowrap>מסך משובים</td>
<td class="title_sort" align=center width=14%>מסך מיקודי טיולים</td>
<td class="title_sort" align=center width=14%>דש בורד מכירות</td>
<TD class="title_sort" align =center width=14%>כללי</TD>

</tr>
<tr><td valign=top bgcolor="#e6e6e6">
<table border=0 cellpadding=0 cellspacing=0  align=right ID="Table9">
<tr>
	<td align="<%=align_var%>" class="title_<%if EditFlyCompaniesPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>חברת תעופה&nbsp;&nbsp;</td>
	<td align="<%=align_var%>" nowrap><input type="checkbox" dir="ltr" name="EditFlyCompanies" <%If trim(EditFlyCompanies) = "1" Then%> checked <%End If%> ID="EditFlyCompanies" <%if EditFlyCompaniesPerm=0 then%>disabled<%end if%>></td>
	</tr>
<tr>
	<td align="<%=align_var%>" class="title_<%if EditAirportsPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>נמל טיסה&nbsp;&nbsp;</td>
	<td align="<%=align_var%>" nowrap><input type="checkbox" dir="ltr" name="EditAirports" <%If trim(EditAirports) = "1" Then%> checked <%End If%> ID="EditAirports" <%if EditAirportsPerm=0 then%>disabled<%end if%>></td>
</tr>
<tr>
	<td align="<%=align_var%>" class="title_<%if EditApprovalStatusPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>סטטוס אישור מקומות טיסה&nbsp;&nbsp;</td>
	<td align="<%=align_var%>" nowrap><input type="checkbox" dir="ltr" name="EditApprovalStatus" <%If trim(EditApprovalStatus) = "1" Then%> checked <%End If%> ID="EditApprovalStatus" <%if EditApprovalStatusPerm=0 then%>disabled<%end if%>></td>
</tr>
<tr>
	<td align="<%=align_var%>" class="title_<%if EditCurrencyPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>מטבע עלות טיסה&nbsp;&nbsp;</td>
	<td align="<%=align_var%>" nowrap><input type="checkbox" dir="ltr" name="EditCurrency" <%If trim(EditCurrency) = "1" Then%> checked <%End If%> ID="EditCurrency" <%if EditCurrencyPerm=0 then%>disabled<%end if%>></td>
</tr>	
<tr>
	<td align="<%=align_var%>" class="title_<%if EditFlightStatusPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>סטטוס טיסות&nbsp;&nbsp;</td>
	<td align="<%=align_var%>" nowrap><input type="checkbox" dir="ltr" name="EditFlightStatus" <%If trim(EditFlightStatus) = "1" Then%> checked <%End If%> ID="EditFlightStatus" <%if EditFlightStatusPerm=0 then%>disabled<%end if%>></td>
</tR>

<tr>
	<td align="right" class="title_<%if IsSendFromTLPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>שליחת מיילים בשם משתמש אחר</td>
	<td align="right" nowrap><input type="checkbox" dir="ltr" name="IsSendFromTL" <%If trim(IsSendFromTL) = "1" Then%> checked <%End If%> ID="IsSendFromTL" <%if IsSendFromTLPerm=0 then%>disabled<%end if%>></td>

</tr><tr>
	<td align="right" class="title_<%if UndoPNRPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>ביטול נעילה PNR &nbsp;&nbsp;</td>
    <td align="right" nowrap><input type="checkbox" dir="ltr" name="UndoPNR" <%If trim(UndoPNR) = "1" Then%> checked <%End If%> ID="UndoPNR" <%if UndoPNRPerm=0 then%>disabled<%end if%>></td>

    </tr>
    <tr>
	<td align="right" class="title_<%if DeleteRowMainDataPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>הרשאת מחיקת שורה &nbsp;&nbsp;</td>
    <td align="right" nowrap><input type="checkbox" dir="ltr" name="DeleteRowMainData" <%If trim(DeleteRowMainData) = "1" Then%> checked <%End If%> ID="DeleteRowMainData" <%if DeleteRowMainDataPerm=0 then%>disabled<%end if%>></td>

</tr>
<tr>
	<td align="right" class="title_<%if InsertRowMainDataPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>הרשאת פתיחת שדות חדשים &nbsp;&nbsp;</td>
	<td align="right" nowrap><input type="checkbox" dir="ltr" name="InsertRowMainData" <%If trim(InsertRowMainData) = "1" Then%> checked <%End If%> ID="InsertRowMainData" <%if InsertRowMainDataPerm=0 then%>disabled<%end if%>></td>

</tr><tr>
	<td align="right" class="title_<%if EditMainDataPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>הרשאות שינוי שדות &nbsp;&nbsp;</td>
    <td align="right" nowrap><input type="checkbox" dir="ltr" name="EditMainData" <%If trim(EditMainData) = "1" Then%> checked <%End If%> ID="EditMainData" <%if EditMainDataPerm=0 then%>disabled<%end if%>></td>

</tr>

</table>
</td>
<td valign=top bgcolor="#e6e6e6">
<table border=0 cellpadding=0 cellspacing=0  align=right ID="Table5">
<tr>
	<td align="right" class="title_<%if ReportMaakavRishumViewPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap  dir=rtl valign=top>הפקת דוח מעקב רישום</td>
	<td align="right"  nowrap><input type="checkbox" dir="ltr" name="ReportMaakavRishumView" <%If trim(ReportMaakavRishumView) = "1" Then%> checked <%End If%> ID="ReportMaakavRishumView" <%if ReportMaakavRishumViewPerm=0 then%>disabled<%end if%>></td>
</tr>
<tr>
	<td align="right" class="title_<%if ReportCloseProcViewPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>הפקת דוח  מגמות ביקושים ואחוזי סגירה</td>
	<td align="right"  nowrap><input type="checkbox" dir="ltr" name="ReportCloseProcView" <%If trim(ReportCloseProcView) = "1" Then%> checked <%End If%> ID="ReportCloseProcView" <%if ReportCloseProcViewPerm=0 then%>disabled<%end if%>></td>
</tr>
</table>
</td>
<td valign=top bgcolor="#e6e6e6"><table border=0 cellpadding=0 cellspacing=0  align=right ID="Table14">
	<td align="right" class="title_<%if OperationScreenViewPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>מורשה לצפיה בכל הטיולים  &nbsp;&nbsp;</td>
	<td align="right"  nowrap><input type="checkbox" dir="ltr" name="OperationScreenView" <%If trim(OperationScreenView) = "1" Then%> checked <%End If%> ID="OperationScreenView" <%if OperationScreenViewPerm=0 then%>disabled<%end if%>></td>
</table></td>
<td valign=top bgcolor="#e6e6e6" align=right >
<table border=0 cellpadding=0 cellspacing=0 align=right>
<tr><td align="right" class="title_<%if ScreenGeneralSendSmsPerm=0 then%>disabled<%else%>show<%end if%>_form"  dir=rtl>שליחת סמס ממסך מרכז משובים</td>
<td align="right" valign=top><input type="checkbox" dir="ltr" name="ScreenGeneralSendSms" <%If trim(ScreenGeneralSendSms) = "1" Then%> checked <%End If%> ID="ScreenGeneralSendSms" <%if ScreenGeneralSendSmsPerm=0 then%>disabled<%end if%>></td>
</tr>
<tr>
<td align="right" class="title_<%if ScreenContactSendSmsPerm=0 then%>disabled<%else%>show<%end if%>_form"  dir=rtl>שליחת סמס ממסך איש קשר</td>
<td align="right" valign=top><input type="checkbox" dir="ltr" name="ScreenContactSendSms" <%If trim(ScreenContactSendSms) = "1" Then%> checked <%End If%> ID="ScreenContactSendSms" <%if ScreenContactSendSmsPerm=0 then%>disabled<%end if%>></td></tr>

</table>
</td>
<td valign=top bgcolor="#e6e6e6" align=right>
<table border=0 cellpadding=0 cellspacing=0 align=right>
<tr>
	<td align="right" nowrap dir=rtl class="title_<%if Screen_TourVisiblePerm=0 then%>disabled<%else%>show<%end if%>_form">להציג טיול בתצוגה</td>
	<td nowrap align=left><input type="checkbox" dir="ltr" name="Screen_TourVisible" <%If trim(Screen_TourVisible) = "1" Then%> checked <%End If%> ID="Screen_TourVisible" <%if Screen_TourVisiblePerm=0 then%>disabled<%end if%>></td>
</tr>
<tr>
	<td align="right" nowrap dir=rtl class="title_<%if ScreenTourStatusPerm=0 then%>disabled<%else%>show<%end if%>_form">שינוי סטטוס </td>
	<td  nowrap align=left><input type="checkbox" dir="ltr" name="ScreenTourStatus" <%If trim(ScreenTourStatus) = "1" Then%> checked <%End If%> ID="ScreenTourStatus" <%if ScreenTourStatusPerm=0 then%>disabled<%end if%>></td>
</tr>
<tr>
	<td align=right  dir=rtl class="title_<%if ScreenSendMailPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap>שליחת  אימייל ממסך ההגדרות</td>
	<td  nowrap align=left><input type="checkbox" dir="ltr" name="ScreenSendMail" <%If trim(ScreenSendMail) = "1" Then%> checked <%End If%> ID="ScreenSendMail" <%if ScreenSendMailPerm=0 then%>disabled<%end if%>></td>
</tr>
<tr>
	<td align=right  nowrap  dir=rtl class="title_<%if ScreenTourOrderPerm=0 then%>disabled<%else%>show<%end if%>_form"> שינוי מיקום שורות במסך הצפיה</td>
	<td nowrap align=left><input type="checkbox" dir="ltr" name="ScreenTourOrder" <%If trim(ScreenTourOrder) = "1" Then%> checked <%End If%> ID="ScreenTourOrder" <%if ScreenTourOrderPerm=0 then%>disabled<%end if%>></td>
</tr>
</table>
</td>
<td valign=top bgcolor="#e6e6e6" align=right><table border=0 cellpadding=0 cellspacing=0 align=right>
   <tr><td align=right  dir=rtl class="title_<%if Edit_DepartmentAppealPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap> שינוי שדה שיוך למחלקה <BR>בטופס רישום חתום</td>
<td  align=left valign=top><input type="checkbox" dir="ltr" name="Edit_DepartmentAppeal" <%If trim(Edit_DepartmentAppeal) = "1" Then%> checked <%End If%> ID="Edit_DepartmentAppeal" <%if Edit_DepartmentAppealPerm=0 then%>disabled<%end if%>>
   </td></tr>
   <tr >
	<td align="right" class="title_<%if salesControl_Points_EditPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>עדכון שדות  מסך הגדרות יעדים</td>
	<td align="right"><input type="checkbox" dir="ltr" name="salesControl_Points_Edit" <%If trim(salesControl_Points_Edit) = "1" Then%> checked <%End If%> ID="salesControl_Points_Edit" <%if salesControl_Points_EditPerm=0 then%>disabled<%end if%>></td>
	</tr><tr>
	<td align="right" class="title_<%if salesControl_EditPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>עדכון שדות שליטה חודשי / שנתי</td>
		<td align="right"><input type="checkbox" dir="ltr" name="salesControl_Edit" <%If trim(salesControl_Edit) = "1" Then%> checked <%End If%> ID="salesControl_Edit" <%if salesControl_EditPerm=0 then%>disabled<%end if%>></td>

</tr>

</table></td>
<TD valign=top bgcolor="#e6e6e6"  align=right><table border=0 cellpadding=0 cellspacing=0 ID="Table4">

<tr>
	<td align="right" class="title_<%if CHIEFPerm=0 then%>disabled<%else%>show<%end if%>_form" dir=rtl>&nbsp;מורשה מחיקה&nbsp;</td>
<td align="left" ><input type="checkbox" dir="ltr" name="chief" <%If trim(chief) = "1" Then%> checked <%End If%> ID="chief" <%if CHIEFPerm=0 then%>disabled<%end if%>></td>

</tr>
<tr>
	<td align="right" class="title_<%if InsertNewUserPerm=0 then%>disabled<%else%>show<%end if%>_form" dir=rtl>מורשה הוספת משתמש</td>
<td align="left" ><input type="checkbox" dir="ltr" name="InsertNewUser" <%If trim(InsertNewUser) = "1" Then%> checked <%End If%> ID="InsertNewUser"  <%if InsertNewUserPerm=0 then%>disabled<%end if%>></td>

</tr>
<tr>	<td align="right" class="title_<%if IsTypingClientPerm=0 then%>disabled<%else%>show<%end if%>_form"  dir=rtl valign=top nowrap>הקלדה אינטרנטית</td>
	<td align="left" ><input type="checkbox" dir="ltr" name="IsTypingClient" <%If trim(IsTypingClient) = "1" Then%> checked <%End If%> ID="IsTypingClient" <%if IsTypingClientPerm=0 then%> disabled <%end if%>></td>
</tr>
<tr>	<td align="right"  class="title_<%if UpdateFieldUserBlockedPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top >שחרור הנעילת משתמש</td>
	<td align="right"><input type="checkbox" dir="ltr" name="UpdateFieldUserBlocked" <%If trim(UpdateFieldUserBlocked) = "1" Then%> checked <%End If%> ID="UpdateFieldUserBlocked"  <%if UpdateFieldUserBlockedPerm=0 then%>disabled<%end if%>></td>
</tr>

<tr><td  align="right" class="title_<%if Add_GroupsToursPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl>רישום לטיולים של הקבוצות הסגורות</td>
	<td align="right" ><input type="checkbox" dir="ltr" name="Add_GroupsTours" <%If trim(Add_GroupsTours) = "1" Then%> checked <%End If%> ID="Add_GroupsTours" <%if Add_GroupsToursPerm=0 then%>disabled<%end if%>></td>

</tr>
<tr>
	<td align="right" class="title_<%if Sms_WritePerm=0 then%>disabled<%else%>show<%end if%>_form"dir=rtl>שינוי תוכן הודעת sms</td>
	<td align="right" ><input type="checkbox" dir="ltr" name="Sms_Write" <%If trim(Sms_Write) = "1" Then%> checked <%End If%> ID="Sms_Write" <%if Sms_WritePerm=0 then%>disabled<%end if%>></td>

</tr>
<tr>
	<td align="right" class="title_<%if Edit_IsNewLettersPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>לעדכון שדה מעוניין בקבלת דיוור</td>
	<td align="right"><input type="checkbox" dir="ltr" name="Edit_IsNewLetters" <%If trim(Edit_IsNewLetters) = "1" Then%> checked <%End If%> ID="Edit_IsNewLetters" <%if Edit_IsNewLettersPerm=0 then%>disabled<%end if%>></td>
</tr><tr>
	<td align="right" class="title_<%if EmailGroupSendPerm=0 then%>disabled<%else%>show<%end if%>_form" dir=rtl>שליחת  אימייל קבוצתי</td>

	<td align="right"><input type="checkbox" dir="ltr" name="EmailGroupSend" <%If trim(EmailGroupSend) = "1" Then%> checked <%End If%> ID="EmailGroupSend" <%if EmailGroupSendPerm=0 then%>disabled<%end if%>></td>

</tr>
<tr>	<td align="right"  class="title_<%if DashBoardViewPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>מורשה לצפיה בדש בורד איש קשר</td>
	<td align="right"><input type="checkbox" dir="ltr" name="DashBoardView" <%If trim(DashBoardView) = "1" Then%> checked <%End If%> ID="DashBoardView" <%if DashBoardViewPerm=0 then%>disabled<%end if%>></td>
</tr>
<tr>	<td align="right"  class="title_<%if UploadUserFilesPerm=0 then%>disabled<%else%>show<%end if%>_form" nowrap dir=rtl valign=top>חוזי עבודה</td>
	<td align="right"><input type="checkbox" dir="ltr" name="UploadUserFiles" <%If trim(UploadUserFiles) = "1" Then%> checked <%End If%> ID="UploadUserFiles"  <%if UploadUserFilesPerm=0 then%>disabled<%end if%>></td>
</tr>

<tr><td colspan=2 align=right><table align=right>
<tr><td align="right" class="title_show_form"> :דרגת חשיבות ידיעת המשוב</td></tr>
<tr>
<td align="right">
	<select dir="<%=dir_obj_var%>" name="ImportanceId" class="norm" style="width:150px" ID="ImportanceId">
	<option value="0" >בחר דרגת חשיבות<%'=arrTitles(15)%></option>
<%	sqlstr = "SELECT ImportanceId,ImportanceName FROM Importance  ORDER BY ImportanceId"
		set rs_Imp = con.getRecordSet(sqlstr)
		while not rs_Imp.eof	%>
	<option value="<%=trim(rs_Imp(0))%>" <%If trim(ImportanceId) = trim(rs_Imp(0)) Then%> selected <%End If%>><%=rs_Imp(1)%></option>
<%	rs_Imp.moveNext
		wend
		set rs_Imp = nothing %>
	</select>
</td>
</tr></table></td>
</tr>



</table>

</TD>



</tr>



</table>
</td></tr>
<%'response.end%>

</table>
</td></tr>
<%'----------response.end%>
</table></td></tr>


<%'---response.end%>

<tr style="background-color:#ffffff"><td height=1 nowrap colspan=3 height=1></td></tr>



<% If trim(is_groups) = "1" Then %>
<tr><td height=5 nowrap colspan=3></td></tr>
<tr><td height=1 nowrap colspan=3 bgcolor=#808080></td></tr>
<tr><td height=5 nowrap colspan=3></td></tr>
<tr><td colspan=3 align=center><table cellpadding=2 cellspacing=0 width=600 border=0>
<tr><td class="title" align="<%=align_var%>" colspan=3><!--הרשאות--><%=arrTitles(21)%></td></tr>
<tr>
    <td width=300 nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top>
    <!--העובד יוכל להזין את כל הטפסים המשויכים לקבוצות אלו.--><%=arrTitles(23)%></td>
	<td align="<%=align_var%>" valign=top>	
    <INPUT type=image src="../../images/icon_find.gif" style="vertical-align:top" name=groups_list id="groups_list" onclick='window.open("groups_list.asp?USER_ID=<%=USER_ID%>&user_groups=" + GROUP_ID.value,"TypesList","left=300,top=100,width=250,height=350,scrollbars=1"); return false;'>&nbsp;
    <span class="Form_R" dir="<%=dir_obj_var%>" style="width:190;line-height:16px" name="groups_names" id="groups_names"><%=Groups%></span>
    <input type=hidden name=GROUP_ID id=GROUP_ID value="<%=user_groups%>">
	</td>
	<td align="<%=align_var%>" class="title_show_form" nowrap width="100" valign=top><!--קבוצה--><%=arrTitles(14)%></td>
</tr> 
<tr>
    <td width=300 nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top>
   <!-- העובד יוכל לצפות או לעדכן את כל הטפסים שהוזנו ע"י העובדים בקבוצות אלו.--><%=arrTitles(24)%>
    </td>
	<td align="<%=align_var%>" valign=top>
    <INPUT type=image src="../../images/icon_find.gif" style="vertical-align:top" onclick='window.open("res_groups_list.asp?USER_ID=<%=USER_ID%>&res_groups_id=" + R_GROUP_ID.value,"TypesList","left=300,top=100,width=250,height=350,scrollbars=1"); return false;'>&nbsp;
    <span class="Form_R" dir="<%=dir_obj_var%>" style="width:190;line-height:16px" name="ResInGroups" id="ResInGroups"><%=ResInGroups%></span>
    <input type=hidden name=R_GROUP_ID id="R_GROUP_ID" value="<%=ResInGroupsID%>">
	</td>
	<td align="<%=align_var%>" class="title_show_form" nowrap width="100" valign=top><!--אחראי קבוצות--><%=arrTitles(17)%></td>
</tr>
<tbody name="edit_appeal_tbody" id="edit_appeal_tbody" <%If trim(ResInGroupsID) = "" Then%> style="display:none" <%End If%>>
<tr>
	<td colspan=3 align="<%=align_var%>" dir="<%=dir_var%>" class="title_show_form">&nbsp;<!--אחראי מורשה צפיה בלבד--><%=arrTitles(19)%><input type=radio name="EDIT_APPEAL" value="0" <%If trim(EDIT_APPEAL) = "0" Then%> checked <%End If%> ID="Radio1" style="position:relative;top:2px"></td>
</tr>
<tr>
	<td colspan=3 align="<%=align_var%>" dir="<%=dir_var%>" class="title_show_form">&nbsp;<!--אחראי מורשה עדכון--><%=arrTitles(20)%><input type=radio name="EDIT_APPEAL" value="1" <%If trim(EDIT_APPEAL) = "1" Then%> checked <%End If%> ID="Radio2" style="position:relative;top:2px"></td>
</tr>
</tbody>
<tr><td height=5 nowrap colspan=3></td></tr>
<tr><td height=1 nowrap colspan=3 bgcolor=#808080></td></tr>
</table></td></tr>
<%End If%>
<tr><td height=5 nowrap colspan=3></td></tr>
<tr><td class="title" align="<%=align_var%>" colspan=3 dir=rtl><!--הרשאות-->הרשאות כניסה למודולים&nbsp;&nbsp;<span style="font-weight:500"><%=arrTitles(22)%></span> (2 שורות של הטאבים העליונים של המערכת)</td></tr>
<tr><td colspan=3><table cellpadding=0 cellspacing=0 width="100%" border=0 dir=rtl>
<% i=1%>
<tr><td width=100%><table border="0" ID="Table12" cellpadding=0 cellspacing=0><tr>
<%if Len(USER_ID) > 0 then
sql_obj="Select bar_id,bar_title,parent_id,parent_title,dbo.getbarParentPerm("&dep_id &","& workerId &",parent_id) as ParentPerm,dbo.getbarParentPerm("&dep_id &","& workerId &",bar_id) as barPerm from dbo.bar_organizations_table('"&OrgID&"') "&_
" WHERE PARENT_ID IS NOT NULL AND IS_VISIBLE = '1' Order by Parent_Order, bar_Order"
else
sql_obj="Select bar_id,bar_title,parent_id,parent_title,0 as ParentPerm,0 as barPerm from dbo.bar_organizations_table('"&OrgID&"') "&_
" WHERE PARENT_ID IS NOT NULL AND IS_VISIBLE = '1' Order by Parent_Order, bar_Order"

end if

'response.Write sql_obj
set rs_obj=con.getRecordSet(sql_obj)
While not rs_obj.eof
	barID = trim(rs_obj(0))
	barTitle = trim(rs_obj(1))
	barParent = trim(rs_obj(2))
	parentTitle = trim(rs_obj(3))
	ParentPerm=trim(rs_obj(4))
	barPerm=trim(rs_obj(5))
'	response.Write "i=" & i &"--parentTitle="&parentTitle &":old_parent="& old_parent &"<BR>"
	if parentTitle<>old_parent and i>1 then%>
	</tr></table></td></tr><tr><td height=1 bgcolor=#ffffff width=100%></td></tr><tr><td width=100%><table border="0" cellpadding=0 cellspacing=0><tr>
	<%end if
	If trim(barID) = "43" Then
		barTitle = barTitle & " <font color=red>(הרשאות טפסים)</font>"
	ElseIf trim(barID) = "10" Then
		barTitle = barTitle & " <font color=red></font>"	
	End If
	dep_id= trim(Request.QueryString("dep_id")) 

	JobId = trim(Request.QueryString("JobId")) 

	If Len(USER_ID) > 0 Or (JobId <> "") then	   	
	
		If Len(USER_ID) > 0 Then
			sql_obj="Select is_visible From dbo.bar_users_table('" & OrgID & "','" & USER_ID & "') WHERE bar_id = " & barID
		Else
			sql_obj="SELECT is_visible FROM dbo.bar_jobs WHERE (organization_id = '" & OrgID & "') AND (job_id = '" & JobId & "') " & _
			" AND (bar_id = " & barID & ")"
		End If
	'	Response.Write sql_obj
		set rs_visible=con.getRecordSet(sql_obj)
		If not rs_visible.eof Then		
			barVisible = rs_visible("is_visible")
		Else
			barVisible = 0
		End If
		set rs_visible = Nothing
	Else
		barVisible = 1
	End If

    If isNumeric(barParent) Then		
	If old_parent <> parentTitle Then	

	If Len(USER_ID) > 0 Or (JobId <> "")  then
		If Len(USER_ID) > 0 Then
			sql_obj="SELECT is_visible FROM dbo.bar_users_table('" & OrgID & "','" & USER_ID & "') WHERE bar_id = " & barParent
		Else
		sql_obj="SELECT is_visible FROM dbo.bar_jobs WHERE (organization_id = '" & OrgID & "') AND (job_id = '" & job_id & "') " & _
		" AND (bar_id = " & barParent & ")"			
		End If	
		'Response.Write sql_obj
		'Response.End
		set rs_parent=con.getRecordSet(sql_obj)
		If not rs_parent.eof Then					
			parentVisible = trim(rs_parent(0))			
		End If
		set rs_parent = Nothing		
	Else
		parentVisible = 1	
	End If	
%>
	<td class="form_titleNormal" align="right" colspan=2 ><input type=checkbox dir="ltr" name="is_visible<%=barParent%>" <%If trim(parentVisible) = "1" Then%> checked <%End If%> <%if ParentPerm=0 then%>disabled <%end if%> ID="is_visible<%=barParent%>" onclick="return check_all_bars(this,'<%=barParent%>')"></td>
	<td align="right" <%if ParentPerm=1 then%>class="form_titleNormal"<%else%>class="title_disabled_form"<%end if%> width=120  nowrap dir="<%=dir_obj_var%>" bgcolor=#bbbad6><%=parentTitle%><%'=ParentPerm%>&nbsp;</td>
<%
	End If	
	End If
%>
	<td align="<%=align_var%>" colspan=2><input type=checkbox dir="ltr" name="is_visible<%=barID%>" id="<%=barID%>!<%=barParent%>"  <%If trim(barVisible) = "1" Then%> checked <%End If%> <%if barPerm=0 then%>disabled<%end if%>></td>
	<td align="<%=align_var%>" class="title_<%if barPerm=1 then%>show<%else%>disabled<%end if%>_form"  nowrap dir="<%=dir_obj_var%>"><%=barTitle%><%'=barPerm%>&nbsp;</td>
<% 

   old_parent = parentTitle
   rs_obj.moveNext
i=i+1
   Wend
   set  rs_obj = Nothing
%>
</table></td></tr>
<tr><td height=5 nowrap colspan=3></td></tr>
<tr><td height=1 nowrap colspan=3 bgcolor=#808080></td></tr>
<tr><td height=10 nowrap colspan=3></td></tr>
<tr>
<td colspan=3>
<table cellpadding=0 cellspacing=0 width=500 align=center>
<tr>
<td width=50% align="center"><input type=button class="but_menu" style="width:90px" onclick="document.location.href='default.asp';" value="<%=arrButtons(2)%>" id=Button2 name=Button2></td>
<td width=50 nowrap></td>
<td width=50% align=center><input type=button class="but_menu" style="width:90px" onclick="javascript:CheckFields(); return false;" value="<%=arrButtons(1)%>" id=Button1 name=Button1></td></tr>
</table></td></tr>
<tr><td colspan=3 height=10 nowrap></td></tr>
</table></form>
</td></tr></table>
   	<SCRIPT LANGUAGE="JavaScript">document.write(getCalendarStyles());</SCRIPT>
			<SCRIPT LANGUAGE="JavaScript" type="text/javascript">
            <!--
            var cal1xx = new CalendarPopup('CalendarDiv');
                cal1xx.showNavigationDropdowns();
                cal1xx.yearSelectStart
                cal1xx.offsetX = -50;
                cal1xx.offsetY = 0;
            //-->
			</SCRIPT>
					<DIV ID='CalendarDiv' name='CalendarDiv' STYLE='VISIBILITY:hidden;POSITION:absolute;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>

</body>
</html>
<%set con = nothing%>
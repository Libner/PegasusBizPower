<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
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
    	if len(BIRTHDAY) then
			else
			 BIRTHDAY="null"
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
     
     'response.Write "IsTypingClient="&IsTypingClient
      
 
 
	 
     
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
				"COMPANIES,TASKS,CASH_FLOW,CHIEF,Archive_appeal,IP_login,Add_Insurance, Add_GroupsTours,Sms_Write,EmailGroupSend,User_Screen_TourVisible,User_ScreenTourStatus,User_ScreenSendMail,User_ScreenTourOrder,User_SendSms_GeneralScreen,User_SendSms_ContactScreen,EDIT_APPEAL,Order_Price,User_OperationScreenView,ReportCloseProcView,ReportMaakavRishumView,DashBoardView,Edit_DepartmentAppeal,salesControl_Edit,salesControl_Points_Edit,EditFlyCompanies,EditAirports,EditApprovalStatus,EditCurrency,EditFlightStatus,ExportDataToExcel,EditMainData,InsertRowMainData,DeleteRowMainData,Edit_IsNewLetters,IsTypingClient,EMAIL_PRIVATE,MOBILE_PRIVATE,BIRTHDAY) " &_
				" values ('" & FIRSTNAME &"','"& LASTNAME &"','"& FIRSTNAME_ENG &"','"& LASTNAME_ENG &"','"& LOGINNAME &"','"& PASSWORD &"',1,"& OrgID &","&_
				job_id &"," & dep_Id &"," & ImportanceId &",'"& phone &"','"& mobile &"','"& fax &"','"& email &"','" & Month_Min_Order & "','" &_
				WORK_PRICING& "','" &SURVEYS& "','" &EMAILS& "','" &COMPANIES& "','" & TASKS & "','" &_
				CASH_FLOW & "','" & CHIEF  & "','" & arch_app & "'," & IP_login  & "," & Add_Insurance  & "," & Add_GroupsTours &","& Sms_Write &","& EmailGroupSend & ",'" & Screen_TourVisible & "','" & ScreenTourStatus & "','" & ScreenSendMail & "','" & ScreenTourOrder & "','" & ScreenGeneralSendSms & "','" & ScreenContactSendSms & "','" & EDIT_APPEAL & "','" & Order_Price & "','" & ReportMaakavRishumView  & "','" & ReportCloseProcView  & "','" & OperationScreenView &"','" & DashBoardView &"','" &Edit_DepartmentAppeal & "','" & salesControl_Edit &"' "& ",'" & salesControl_Points_Edit &"','"& EditFlyCompanies &"','"& EditAirports &"','"& ExportDataToExcel  &"','"& EditMainData &"','"& InsertRowMainData &"','"& DeleteRowMainData &"','"& EditApprovalStatus &"','"& EditCurrency &"','"& EditFlightStatus &"','"& Edit_IsNewLetters &"','"& IsTypingClient &"','"& emailPrivate  &"','"& mobilePrivate &"'," & birthday &"); SELECT @@IDENTITY AS NewID"
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
					"', Order_Price = '" & Order_Price & "',User_OperationScreenView='"& OperationScreenView  & "',ReportCloseProcView='"& ReportCloseProcView & "',ReportMaakavRishumView='"& ReportMaakavRishumView & "',DashBoardView='"& DashBoardView &"',Edit_DepartmentAppeal='"& Edit_DepartmentAppeal &"',salesControl_Edit='"& salesControl_Edit &"',salesControl_Points_Edit='"& salesControl_Points_Edit & _
					"', EditFlyCompanies='" & EditFlyCompanies & "', EditAirports='" & EditAirports & "', EditApprovalStatus='" & EditApprovalStatus &_
					"', UndoPNR='" & UndoPNR &_
					"', IsSendFromTL='" & IsSendFromTL &_
    		        "', EditCurrency='" & EditCurrency & "', EditFlightStatus='" & EditFlightStatus &"',ExportDataToExcel='"& ExportDataToExcel  &"',EditMainData='"& EditMainData  &"',InsertRowMainData='"& InsertRowMainData  &"',DeleteRowMainData='"& DeleteRowMainData  & "', Edit_IsNewLetters='" & Edit_IsNewLetters & "',IsTypingClient='" &IsTypingClient  &"',MOBILE_PRIVATE='"& mobilePrivate  &"',EMAIL_PRIVATE='"& EMAILPrivate &"',BIRTHDAY="& BIRTHDAY &" WHERE USER_ID=" & USER_ID
					 
	'Response.Write sqlStr
	'Response.End
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
	" job_id,WORK_PRICING,SURVEYS,EMAILS,COMPANIES,TASKS,CASH_FLOW,CHIEF,ISNULL(EDIT_APPEAL,0),"&_
	" isNULL(Voice_Operator,0),IsNUll(Order_Price,0),Add_Insurance, Add_GroupsTours ,Sms_Write,EmailGroupSend,IP_login,IsNULL(User_Screen_TourVisible,0),IsNULL(User_ScreenTourStatus,0),IsNULL(User_ScreenSendMail,0),ImportanceId,IsNULL(User_SendSms_GeneralScreen,0),IsNULL(User_SendSms_ContactScreen,0),IsNULL(User_ScreenTourOrder,0),IsNULL(User_OperationScreenView,0)," & _
	"IsNULL(ReportCloseProcView,0),IsNULL(ReportMaakavRishumView ,0),IsNULL(DashBoardView ,0),Department_Id, Edit_DepartmentAppeal,salesControl_Edit,salesControl_Points_Edit,EditFlyCompanies,EditAirports,EditApprovalStatus,EditCurrency,EditFlightStatus,Edit_IsNewLetters,IsNULL(ExportDataToExcel ,0) ,IsNULL(EditMainData ,0),IsNULL(InsertRowMainData ,0),IsNULL(DeleteRowMainData ,0),IsNULL(Archive_appeal ,0),IsNULL(UndoPNR ,0) ,IsNULL(IsSendFromTL,0),FIRSTNAME_ENG,LASTNAME_ENG,IsNULL(IsTypingClient,0), " & _
	"EMAIL_PRIVATE,MOBILE_PRIVATE,BIRTHDAY	 FROM dbo.USERS WHERE USER_ID=" & USER_ID  
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
<FORM name="frmMain" ACTION="addWorker.asp?USER_ID=<%=USER_ID%>" METHOD="post" onSubmit="return CheckFields()" ID="Form1">
<table align=center border="0" width="600" cellpadding="2" cellspacing="0">
<tr><td height=5 nowrap colspan=2></td></tr>
<tr><td colspan=2 align=left><table border=0 cellpadding=0 cellspacing=0  align=left ID="Table1">
<tr>
<td width=50% align="left"><input type=button class="but_menu" style="width:90px" onclick="document.location.href='default.asp';" value="<%=arrButtons(2)%>" id="Button3" name=Button2></td>
<td width=50 nowrap></td>
<td width=50% align=left><input type=button class="but_menu" style="width:90px" onclick="javascript:CheckFields(); return false;" value="<%=arrButtons(1)%>" id="Button4" name=Button1></td></tr>
</table></td></tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap>
	<select dir="<%=dir_obj_var%>" name="JOB_ID" class="norm" style="width:200px" ID="JOB_ID">
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
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--עיסוק--><%=arrTitles(7)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap>
	<select dir="<%=dir_obj_var%>" name="Dep_ID" class="norm" style="width:200px" ID="Dep_ID">
	<option value="0">שיוך למחלקה</option>
<%	sqlstr = "SELECT departmentId,departmentName FROM Departments ORDER BY PriorityLevel desc"
		set rs_dep = con.getRecordSet(sqlstr)
		while not rs_dep.eof	%>
	<option value="<%=trim(rs_dep(0))%>" <%If trim(dep_id) = trim(rs_dep(0)) Then%> selected <%End If%>><%=rs_dep(1)%></option>
<%	rs_dep.moveNext
		wend
		set rs_dep = nothing %>
	</select>
	</td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;שיוך למחלקה&nbsp;</td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="<%=dir_obj_var%>" name="FIRSTNAME" value="<%=vFix(FIRSTNAME)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--שם פרטי--><%=arrTitles(3)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="<%=dir_obj_var%>" name="LASTNAME" value="<%=vFix(LASTNAME)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--שם משפחה--><%=arrTitles(4)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="FIRSTNAME_ENG" value="<%=vFix(FIRSTNAME_ENG)%>"  style="width:200" maxlength=20 style="font-family:arial" ></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp; שם פרטי באנגלית<%'=arrTitles(3)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr>" name="LASTNAME_ENG" value="<%=vFix(LASTNAME_ENG)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;שם משפחה באנגלית<%'=arrTitles(4)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="LOGINNAME" value="<%=vFix(LOGINNAME)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--שם משתמש--><%=arrTitles(5)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="PASSWORD" value="<%=vFix(PASSWORD_U)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--סיסמה--><%=arrTitles(6)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>

<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="phone" value="<%=vfix(phone)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--טלפון--><%=arrTitles(8)%>&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="mobile" value="<%=vfix(mobile)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" align="<%=align_var%>" class="title_show_form" nowrap width="150"><!--נייד-->נייד עבודה<%'=arrTitles(9)%>&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="mobilePrivate" value="<%=vfix(mobilePrivate)%>"  style="width:200" maxlength=20 style="font-family:arial" ID="mobilePrivate"></td>
	<th align="<%=align_var%>" align="<%=align_var%>" class="title_show_form" nowrap width="150"><!--נייד-->נייד פרטי<%'=arrTitles(9)%>&nbsp;</th>
</tr>

<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="fax" value="<%=vfix(fax)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--פקס--><%=arrTitles(10)%>&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="email" value="<%=vfix(email)%>"  style="width:200" maxlength=100 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--דואר אלקטרוני-->דואר אלקטרוני עבודה<%'=arrTitles(11)%>&nbsp;<span style="color:#FF0000">*</span></th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="emailPrivate" value="<%=vfix(emailPrivate)%>"  style="width:200" maxlength=100 style="font-family:arial" ID="emailPrivate"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--דואר אלקטרוני-->דואר אלקטרוני פרטי<%'=arrTitles(11)%>&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap>
	 		<a href='' onclick='callCalendar(document.Form1.birthday,"Asbirthday");return false;' id="Asbirthday">
			<image src='../../images/calend.gif' border=0 ></a>&nbsp;<input dir="ltr" type="text" class="passw" size=10 id="birthday" name="birthday" value="<%=birthday%>" maxlength=10 readonly>&nbsp;
</td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;תאריך לידה&nbsp;</th>
</tr>

<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="Month_Min_Order" value="<%=vfix(Month_Min_Order)%>" size="3" maxlength="3" style="font-family:arial" onkeypress="GetNumbers();"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;יעד הזמנות מינימלי&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="text" size="3" maxlength="3" dir="ltr" style="font-family:arial" name="Order_Price" value="<%=nFix(Order_Price)%>" ID="Order_Price"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>&nbsp;שווי כספי של הזמנה&nbsp;&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="IP_login" <%If trim(IP_login) = "1" Then%> checked <%End If%> ID="IP_login"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>&nbsp;מורשה כניסה רק מי-IP&nbsp;&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="chief" <%If trim(chief) = "1" Then%> checked <%End If%> ID="chief"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>&nbsp;מורשה מחיקה&nbsp;&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="arch_app" <%If trim(arch_app) = "1" Then%> checked <%End If%> ID="arch_app"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>מורשה העברה<br> טופס לארכיון /לטיפול חוזר</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="Add_Insurance" <%If trim(Add_Insurance) = "1" Then%> checked <%End If%> ID="Add_Insurance"></td>
	<th  align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>"ביטוח מול "דיוויד שילד&nbsp;&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="Add_GroupsTours" <%If trim(Add_GroupsTours) = "1" Then%> checked <%End If%> ID="Add_GroupsTours"></td>
	<th  align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>רישום לטיולים של הקבוצות הסגורות&nbsp;&nbsp;</th>
</tr>

<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="Sms_Write" <%If trim(Sms_Write) = "1" Then%> checked <%End If%> ID="Sms_Write"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>מורשה לשינוי תוכן<br>הודעת sms &nbsp;&nbsp;</th>

</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="EmailGroupSend" <%If trim(EmailGroupSend) = "1" Then%> checked <%End If%> ID="EmailGroupSend"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>מורשה לשליחת <br>אימייל קבוצתי &nbsp;&nbsp;</th>

</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="Edit_IsNewLetters" <%If trim(Edit_IsNewLetters) = "1" Then%> checked <%End If%> ID="Edit_IsNewLetters"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>לעדכון שדה מעוניין בקבלת דיוור&nbsp;&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="IsTypingClient" <%If trim(IsTypingClient) = "1" Then%> checked <%End If%> ID="IsTypingClient"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>הרשאת הקלדה אינטרנטית&nbsp;&nbsp;</th>
</tr>

<tr><td colspan=2 height=20  align="<%=align_var%>" class="title"><font size=4>דש בורד מכירות</font></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="Edit_DepartmentAppeal" <%If trim(Edit_DepartmentAppeal) = "1" Then%> checked <%End If%> ID="Edit_DepartmentAppeal"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>מורשה שינוי שדה שיוך למחלקה בטופס רישום חתום&nbsp;&nbsp;</th>

</tr>

<tr><td colspan=2 height=10></td></tr>

<tr><td colspan=2 height=20  align="<%=align_var%>" class="title"><font size=4>מסך מיקודי טיולים</font></td></tr>

<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="Screen_TourVisible" <%If trim(Screen_TourVisible) = "1" Then%> checked <%End If%> ID="Screen_TourVisible"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>להציג טיול בתצוגה&nbsp;&nbsp;</th>

</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="ScreenTourStatus" <%If trim(ScreenTourStatus) = "1" Then%> checked <%End If%> ID="ScreenTourStatus"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>שינוי סטטוס &nbsp;&nbsp;</th>

</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="ScreenSendMail" <%If trim(ScreenSendMail) = "1" Then%> checked <%End If%> ID="ScreenSendMail"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>מורשה לשליחת <br>אימייל ממסך ההגדרות &nbsp;&nbsp;</th>
</tr>

<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="ScreenTourOrder" <%If trim(ScreenTourOrder) = "1" Then%> checked <%End If%> ID="ScreenTourOrder"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>מורשה לשינוי מיקום שורות במסך הצפיה &nbsp;&nbsp;</th>

</tr>

<tr><td colspan=2><br><br></td></tr>
<tr  ><td colspan=2 height=20  align="<%=align_var%>" class="title"><font size=4>מסך משובים</font></td></tr>

<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="ScreenGeneralSendSms" <%If trim(ScreenGeneralSendSms) = "1" Then%> checked <%End If%> ID="ScreenGeneralSendSms"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>מורשה לשליחת <br> סמס ממסך מרכז משובים</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="ScreenContactSendSms" <%If trim(ScreenContactSendSms) = "1" Then%> checked <%End If%> ID="ScreenContactSendSms"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl>מורשה לשליחת <br>סמס ממסך איש קשר &nbsp;&nbsp;</th>

</tr>
<tr><td colspan=2><br><br></td></tr>
<tr  ><td colspan=2 height=20  align="<%=align_var%>" class="title"><font size=4>מסך אופרציה</font></td></tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="OperationScreenView" <%If trim(OperationScreenView) = "1" Then%> checked <%End If%> ID="OperationScreenView"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>מורשה לצפיה בכל הטיולים  &nbsp;&nbsp;</th>

</tr>
<tr><td colspan=2 height=20  align="<%=align_var%>" class="title"><font size=4>מסך טפסים וסקרים - דוח מעקב רישום ואחוזי סגירה </font></td></tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="ReportMaakavRishumView" <%If trim(ReportMaakavRishumView) = "1" Then%> checked <%End If%> ID="ReportMaakavRishumView"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>הפקת דוח מעקב רישום&nbsp;&nbsp;</th>

</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="ReportCloseProcView" <%If trim(ReportCloseProcView) = "1" Then%> checked <%End If%> ID="ReportCloseProcView"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>הפקת דוח  מגמות ביקושים ואחוזי סגירה&nbsp;&nbsp;</th>

</tr>
<tr><td colspan=2><br><br></td></tr>

<tr style="background-color:FFCE42">
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="DashBoardView" <%If trim(DashBoardView) = "1" Then%> checked <%End If%> ID="DashBoardView"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>מורשה לצפיה בדש בורד איש קשר&nbsp;&nbsp;</th>

</tr>
<tr><td colspan=2><br><br></td></tr>

<tr><td colspan=2><br><br></td></tr>
<tr  ><td colspan=2 height=20  align="<%=align_var%>" class="title"><font size=4>מסך מכירות</font></td></tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="salesControl_Edit" <%If trim(salesControl_Edit) = "1" Then%> checked <%End If%> ID="salesControl_Edit"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top> מורשה לעדכון שדות שליטה חודשי / שנתי  &nbsp;&nbsp;</th>

</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="salesControl_Points_Edit" <%If trim(salesControl_Points_Edit) = "1" Then%> checked <%End If%> ID="salesControl_Points_Edit"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top> מורשה לעדכון שדות  מסך הגדרות יעדים  &nbsp;&nbsp;</th>

</tr>


<tr><td colspan=2><br><br></td></tr>
<tr  ><td colspan=2 height=20  align="<%=align_var%>" class="title"><font size=4>טיים לימיטים</font></td></tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="EditFlyCompanies" <%If trim(EditFlyCompanies) = "1" Then%> checked <%End If%> ID="EditFlyCompanies"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>חברת תעופה&nbsp;&nbsp;</th>

</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="EditAirports" <%If trim(EditAirports) = "1" Then%> checked <%End If%> ID="EditAirports"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>נמל טיסה&nbsp;&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="EditApprovalStatus" <%If trim(EditApprovalStatus) = "1" Then%> checked <%End If%> ID="EditApprovalStatus"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>סטטוס אישור מקומות טיסה&nbsp;&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="EditCurrency" <%If trim(EditCurrency) = "1" Then%> checked <%End If%> ID="EditCurrency"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>מטבע עלות טיסה&nbsp;&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="EditFlightStatus" <%If trim(EditFlightStatus) = "1" Then%> checked <%End If%> ID="EditFlightStatus"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>סטטוס טיסות&nbsp;&nbsp;</th>
</tr>

<tr><td colspan=2><br><br></td></tr>
<tr><td colspan=2 align=right class="form_title"><u><b>מסך שליטה ראשי של מערכת הטיסות</b></u></td></tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="EditMainData" <%If trim(EditMainData) = "1" Then%> checked <%End If%> ID="EditMainData"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>הרשאות שינוי שדות &nbsp;&nbsp;</th>
</tr>

<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="InsertRowMainData" <%If trim(InsertRowMainData) = "1" Then%> checked <%End If%> ID="InsertRowMainData"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>הרשאת פתיחת שדות חדשים &nbsp;&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="DeleteRowMainData" <%If trim(DeleteRowMainData) = "1" Then%> checked <%End If%> ID="DeleteRowMainData"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>הרשאת מחיקת שורה &nbsp;&nbsp;</th>
</tr>
 

<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="ExportDataToExcel" <%If trim(ExportDataToExcel) = "1" Then%> checked <%End If%> ID="ExportDataToExcel"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>יצוא קובץ אקסל למסך תעופה&nbsp;&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="UndoPNR" <%If trim(UndoPNR) = "1" Then%> checked <%End If%> ID="UndoPNR"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>ביטול נעילה PNR &nbsp;&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="IsSendFromTL" <%If trim(IsSendFromTL) = "1" Then%> checked <%End If%> ID="IsSendFromTL"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>מורשה לשליחת מיילים של טיים לימיטים בשם משתמש אחר&nbsp;&nbsp;</th>
</tr>


<tr><td colspan=2><br><br></td></tr>


<%if false then%><tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="OperationScreenView" <%If trim(OperationScreenView) = "1" Then%> checked <%End If%> ID="OperationScreenView"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150" dir=rtl valign=top>מורשה לצפיה בכל הטיולים  &nbsp;&nbsp;</th>
</tr>
<%end if%>

<tr><td colspan=2><br><br></td></tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap>
	<select dir="<%=dir_obj_var%>" name="ImportanceId" class="norm" style="width:200px" ID="ImportanceId">
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
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150"><!--עיסוק-->דרגת חשיבות ידיעת המשוב<%'=arrTitles(7)%></td>
</tr>




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
	<th align="<%=align_var%>" class="title_show_form" nowrap width="100" valign=top><!--קבוצה--><%=arrTitles(14)%></td>
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
	<th align="<%=align_var%>" class="title_show_form" nowrap width="100" valign=top><!--אחראי קבוצות--><%=arrTitles(17)%></td>
</tr>
<tbody name="edit_appeal_tbody" id="edit_appeal_tbody" <%If trim(ResInGroupsID) = "" Then%> style="display:none" <%End If%>>
<tr>
	<th colspan=3 align="<%=align_var%>" dir="<%=dir_var%>" class="title_show_form">&nbsp;<!--אחראי מורשה צפיה בלבד--><%=arrTitles(19)%><input type=radio name="EDIT_APPEAL" value="0" <%If trim(EDIT_APPEAL) = "0" Then%> checked <%End If%> ID="Radio1" style="position:relative;top:2px"></th>
</tr>
<tr>
	<th colspan=3 align="<%=align_var%>" dir="<%=dir_var%>" class="title_show_form">&nbsp;<!--אחראי מורשה עדכון--><%=arrTitles(20)%><input type=radio name="EDIT_APPEAL" value="1" <%If trim(EDIT_APPEAL) = "1" Then%> checked <%End If%> ID="Radio2" style="position:relative;top:2px"></th>
</tr>
</tbody>
<tr><td height=5 nowrap colspan=3></td></tr>
<tr><td height=1 nowrap colspan=3 bgcolor=#808080></td></tr>
</table></td></tr>
<%End If%>
<tr><td height=5 nowrap colspan=3></td></tr>
<tr><td class="title" align="<%=align_var%>" colspan=3><!--הרשאות--><%=arrTitles(13)%>&nbsp;&nbsp;<span style="font-weight:500"><%=arrTitles(22)%></span></td></tr>
<tr><td colspan=3><table cellpadding=0 cellspacing=0 width="100%">
<% 
sql_obj="Select bar_id,bar_title,parent_id,parent_title from dbo.bar_organizations_table('"&OrgID&"') "&_
" WHERE PARENT_ID IS NOT NULL AND IS_VISIBLE = '1' Order by Parent_Order, bar_Order"
set rs_obj=con.getRecordSet(sql_obj)
While not rs_obj.eof
	barID = trim(rs_obj(0))
	barTitle = trim(rs_obj(1))
	barParent = trim(rs_obj(2))
	parentTitle = trim(rs_obj(3))
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
<tr>
	<th class="form_title" align="<%=align_var%>" colspan=2><input type=checkbox dir="ltr" name="is_visible<%=barParent%>" <%If trim(parentVisible) = "1" Then%> checked <%End If%> ID="is_visible<%=barParent%>" onclick="return check_all_bars(this,'<%=barParent%>')"></th>
	<th align="<%=align_var%>" class="form_title" nowrap width="250" dir="<%=dir_obj_var%>"><%=parentTitle%>&nbsp;</th>
</tr>
<%
	End If	
	End If
%>
<tr>
	<td align="<%=align_var%>" colspan=2><input type=checkbox dir="ltr" name="is_visible<%=barID%>" id="<%=barID%>!<%=barParent%>"  <%If trim(barVisible) = "1" Then%> checked <%End If%>></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="250" dir="<%=dir_obj_var%>"><%=barTitle%>&nbsp;</th>
</tr>
<% 
   old_parent = parentTitle
   rs_obj.moveNext
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
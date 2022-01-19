<%
  Response.CharSet = "windows-1255"
  Response.Buffer = false
  Server.ScriptTimeout = 600

	
%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
<head>
    <link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body style="margin: 0px; background-color: #E6E6E6">
    <div id="div_save" name="div_save" style="position: absolute; left: 0px; top: 0px;
        width: 100%; height: 100%;">
        <table height="100%" width="100%" cellspacing="2" cellpadding="2" border="0">
            <tr>
                <td align="center" valign="middle">
                    <table dir="rtl" border="0" height="100" width="400" cellspacing="1" cellpadding="1">
                        <tr>
                            <td align="center" class="div_title" valign="middle">
                                מתבצעת טעינת נתונים אנא המתן .....
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>
</body>
<%COMPANIES = trim(Request.Cookies("bizpegasus")("COMPANIES"))
     UserID=trim(Request.Cookies("bizpegasus")("UserID"))
   'sqlstr="select Department_Id from Users where User_Id="&UserIdOrderOwner
    ' 	set rs_D = con.GetRecordSet(sqlstr)
	'	If not rs_D.eof Then
     '     DepartmentId=rs_D("Department_Id")
     '     else
     '     DepartmentId=0
     '  end if
    '      set rs_D=Nothing
          
     OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
     lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))

	 If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	 End If
	 If lang_id = 2 Then
		dir_var = "rtl" :  align_var = "left"  :  dir_obj_var = "ltr"
	Else
		dir_var = "ltr"  :   align_var = "right"  :   dir_obj_var = "rtl"
	End If	    

	 projectID=trim(Request.QueryString("project_ID"))
	 mechanismID=trim(Request.QueryString("mechanism_ID"))
	 from=trim(Request.QueryString("from"))
	 set upl=Server.CreateObject("SoftArtisans.FileUp")		
 ' response.Write  "55"
'fileName= upl.UserFileName
'response.end
	 quest_id = trim(upl.Form("quest_id"))
	 	''update file
		
	
	 
	 
	If not IsNumeric(quest_id) Then
		quest_id = 0
	End If
			if trim(upl.Form("Is_GroupsTour"))="1" then
	Is_GroupsTour=trim(upl.Form("Is_GroupsTour"))
	else
	Is_GroupsTour=0
	end if	
	
	appID = trim(upl.Form("appID"))
	If not IsNumeric(appID) Then
		appID = 0
		
	else
	
		'-------
		if quest_id="17857"  or quest_id="17001"  or   quest_id= "17057" or  quest_id = "16735" then
	
			insT=upl.form("Tour_id")
			if insT="" then insT=0

			insD=upl.form("ToursDate_Id")
			if insD ="" then insD=0

			insG=upl.form("Guide_Id")
			DepartmentOwner=upl.form("DepartmentOwner")
			'response.Write "DepartmentOwner="& DepartmentOwner
			'response.end
			if insG ="" then insG=0


			sqlstr="UPDATE APPEALS SET Tour_Id="& insT &",Departure_Id="& insD &",Guide_Id=" & insG  &" where appeal_id="& appID
			'response.Write sqlstr
			'response.end
			con.ExecuteQuery(sqlstr)

		
		
			'response.Write "111"
			'response.Write upl.UserFileName
			'		response.Write "222"
			'response.end
			if upl.UserFileName<>"" then
			'response.Write(upl.UserFileName)
			'response.end
					str_mappath="../../../download/thanksLetter"
					File_Name=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
					extend = LCase(Mid(File_Name,InstrRev(File_Name,".")+1))
					name_without_extend = LCase(Mid(File_Name,1,Len(File_Name)-Len(extend)-1))
					NewFileName =  "thanksLetter_" & appID
					new_name = NewFileName
					upload = true
					i = 0
					set fs = server.CreateObject("Scripting.FileSystemObject") 
					do while fs.FileExists(Server.MapPath(str_mappath & "/"& new_name & "." & extend ))	
						i =  i + 1
						new_name = NewFileName & "_" & i
					loop
					newFileName = new_name & "." & extend
					set fs = Nothing
					if quest_id="17857"  then	
							upl.Form("f1").SaveAs Server.Mappath("../../../download/thanksLetter/") & "/" & newFileName
									sqlstr="UPDATE APPEALS SET File_ThankLetter='"& NewFileName &"' where appeal_id="& appID
							'	response.write sqlstr
							'response.end
										con.ExecuteQuery(sqlstr)
					end if 

			end if

		end if

	 
		'--
		
		
		
		
		
	End If
	companyId=trim(upl.Form("companyId"))
	If not IsNumeric(companyId) Then
		companyId = 0
	End If	
	contactID=trim(upl.Form("contactID"))
	If not IsNumeric(contactID) Then
		contactID = 0
	End If		
	ReservationId=trim(upl.Form("ReservationId"))
	If not IsNumeric(ReservationId) Then
		ReservationId = 0
	End If		
    RelationId=trim(upl.Form("RelationId"))
	CRM_Country=trim(upl.Form("CRM_Country"))
	If not IsNumeric(CRM_Country) Then
		CRM_Country = 0
	End If
	if not IsNumeric(RelationId) then
	RelationId=0
	end if
	
	UserIdOrderOwner=trim(upl.Form("UserIdOrderOwner"))
	'response.Write UserIdOrderOwner
	
	If not IsNumeric(UserIdOrderOwner) Then
		UserIdOrderOwner = 0
		DepartmentId=0
	else
		sqlstr="select Department_Id from Users where User_Id="&UserIdOrderOwner
		'response.Write sqlstr
    		set rs_D = con.GetRecordSet(sqlstr)
		If not rs_D.eof Then
	      
			if rs_D("Department_Id")<>"null" then
			DepartmentId=rs_D("Department_Id")
	        
			else
			DepartmentId=0
			end if
		else
			DepartmentId=0
		end if
          set rs_D=Nothing
	End If
'	response.Write "DepartmentId="& DepartmentId
	'response.end
	
	''''----------------------------------
	''''----------------------------------
	
	appeal_date = trim(upl.Form("appeal_date"))
	appeal_status=trim(upl.Form("appeal_status"))
					
	if appID = nil and quest_id <> nil then	   
			
		If trim(quest_id) <> "" Then
			sqlstr = "Select ADD_CLIENT From Products WHERE Product_ID=" & quest_id				
			set rsd = con.getRecordSet(sqlstr)
			If not rsd.eof Then				
				addClient = trim(rsd(0))			
			End If
			set rsd = Nothing	
		End If					
	
		If trim(from) = "appeal" And (trim(addClient) = "1" Or  trim(addClient) = "2") Then
	        private_flag = "0"
			if upl.Form("company_name") <> nil then  

				company_name=trim(sFix(upl.Form("company_name")))
				address=trim(sFix(upl.Form("address")))
				company_email=trim(sFix(upl.Form("company_email")))
				cityName=trim(sFix(upl.Form("cityName")))
				zip_code = trim(sFix(upl.Form("zip_code")))
				url=trim(sFix(upl.Form("url")))
				phone1=trim(sFix(upl.Form("phone1")))
				phone2=trim(sFix(upl.Form("phone2")))
				fax1=trim(sFix(upl.Form("fax1")))
				fax2=trim(sFix(upl.Form("fax2")))					
				status=trim(upl.Form("status"))
				If trim(status) = "" Then
					status = "4"
				End If	
				company_desc = sFix(Left(trim(upl.Form("company_desc")),500))
				
				if IsNumeric(companyID) then
					sqlUpd="UPDATE companies SET company_name='" & company_name &_
					"', date_update=getDate() , company_desc='" & company_desc & "'" &_
					", address='"& address & "', city_Name='" & cityName & "', zip_code='" & zip_code &_  
					"', url='" & url & "'" & ", prefix_phone1='"& prefix_phone1 & "', prefix_phone2='"& prefix_phone2 &_
					"', prefix_fax1='"& prefix_fax1 &"', prefix_fax2='"& prefix_fax2 &"'" &_
					", phone1='"& phone1 &"', phone2='"& phone2 &"', fax1='"& fax1 &_
					"', fax2='"& fax2 &"', email='"& company_email &"'"&_   
					"  WHERE company_ID="& companyID
					'Response.Write(sqlUpd)
					'Response.End
					con.ExecuteQuery(sqlUpd)
				else      
					sqlIns="SET NOCOUNT ON; INSERT INTO COMPANIES(company_name,date_update,private_flag,company_desc"&_      
					", address,url,prefix_phone1,prefix_phone2,prefix_fax1,prefix_fax2"&_
					", phone1,phone2,fax1,fax2,zip_code,city_Name,email,status,Organization_ID "&_      
					")  VALUES ( '"& company_name &"',getDate(), '"& private_flag &"','"&company_desc & "','" &_      
					address &"','"& url &"','"& prefix_phone1 &"','"& prefix_phone2 &"','"&_
					prefix_fax1 &"','"& prefix_fax2 &"','"& phone1 &"','"& phone2 &"','"&_
					fax1 &"','"& fax2 &"','"& zip_code &"','"& cityName &"','" & company_email & "','" & status & "',"&_
					OrgID & "); SELECT @@IDENTITY AS NewID"
					'Response.Write(sqlIns)
					'Response.End     
					set rs_tmp = con.getRecordSet(sqlIns)
						companyID = rs_tmp.Fields("NewID").value
					set rs_tmp = Nothing	  
				end if			
				'הוספת איש קשר ללא חברה - בחורים מספר של חברה ללא חברה במערכת
			ElseIf upl.Form("companyID")=nil And upl.Form("company_name") = nil and upl.Form("CONTACT_NAME")<>nil Then 
					sqlstr = "Select company_id From companies WHERE Organization_ID = " & orgID & " And private_flag = 1"
					set rs_private = con.getRecordSet(sqlstr)
					if not rs_private.eof then		
						companyID = rs_private(0)	
					end if
					set rs_private = Nothing	
			End If		
			' REM: add User ====================================================
			if upl.Form("CONTACT_NAME")<>nil then     
			    
				contact_name=sFix(trim(upl.Form("contact_name")))     
				phone=sFix(trim(upl.Form("phone")))
				fax=sFix(trim(upl.Form("fax")))
				cellular=sFix(trim(upl.Form("cellular")))
				email=sFix(trim(upl.Form("email")))   
				messangerName=sFix(trim(upl.Form("messangerName")))
				contact_address=trim(sFix(upl.Form("contact_address")))
				contact_city_name=trim(sFix(upl.Form("contact_city_name")))
				contact_zip_code=trim(sFix(upl.Form("contact_zip_code")))
				contact_desc = sFix(Left(trim(upl.Form("contact_desc")),500))
				contact_types = trim(upl.Form("contact_types"))
				responsible_id = trim(upl.Form("responsible_id"))
				   
				If trim(addClient) = "1" Then 'הוספת איש קשר לחברה לקוחות פרטיים
				
					found_contact_id = 0
					sqlstr = "EXEC dbo.contacts_chk_phone	@OrgId='" & OrgID & "', @EditContactId='" & ContactId & "', " & _
					" @cp='" & cellular & "', @pn='" & phone & "'"	
					set rs_tmp = con.getRecordSet(sqlstr)	
					If rs_tmp.eof = False Then
						found_contact_id = trim(rs_tmp(0))
					End If					
					set rs_tmp = Nothing
					
					If found_contact_id > 0 Then
						contactID = cLng(found_contact_id)
					End If
				
					If IsNumeric(contactID) Then
						sqlUpd="UPDATE contacts SET CONTACT_NAME='" & contact_name &_
						"', email='"& email  &"', phone='"& phone &_
						"', cellular='"& cellular &"',"&_
						" date_update=getDate() WHERE contact_ID="& contactID 
						'Response.Write(sqlUpd)
						'Response.End
						con.ExecuteQuery(sqlUpd)
						
						'--insert into changes table
						sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
						" SELECT 'איש קשר', ' שם:'  + IsNULL(CONTACT_NAME, '') + ' נייד: ' + IsNULL(cellular, ''), CONTACT_ID, 'עדכון', getDate()," & UserID & _
						" FROM dbo.CONTACTS WHERE (CONTACT_ID = "& contactID &")"
						con.executeQuery(sqlstr)  
	     								
					ElseIf IsNumeric(companyID) Then
						sqlIns="SET DATEFORMAT MDY; SET NOCOUNT ON; "&_
						" INSERT INTO CONTACTS(company_id,contact_name,email,phone,cellular,date_update,Organization_ID)" &_				  
						" VALUES ("& companyID & ",'" & contact_name & "','" & email & "','"&_
						phone &"','" & cellular & "'" &	", getDate(), " & OrgID & "); SELECT @@IDENTITY AS NewID"				
						'Response.Write(sqlIns)
						'Response.End
						set rs_tmp = con.getRecordSet(sqlIns)
							contactID = rs_tmp.Fields("NewID").value
						set rs_tmp = Nothing	 
						
						'--insert into changes table
						sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
						" SELECT 'איש קשר', ' שם:'  + IsNULL(CONTACT_NAME, '') + ' נייד: ' + IsNULL(cellular, ''), CONTACT_ID, 'הוספה', getDate()," & UserID & _
						" FROM dbo.CONTACTS WHERE (CONTACT_ID = "& contactID &")"
						con.executeQuery(sqlstr)  					
							
					End If						
				Else
					if trim(contactID)<>"" then   		
						sqlUpd="UPDATE contacts SET CONTACT_NAME='"& CONTACT_NAME &_
						"', email='"& email  &"',phone='"& phone &"', fax='"& fax &_
						"', cellular='"& cellular &"', date_update=getDate()" &_
						", messanger_name='"& messangerName &"', contact_address = '" & contact_address &_
						"', contact_city_name = '" & contact_city_name & "', contact_zip_code = '" & contact_zip_code &_    
						"', contact_desc = '" & contact_desc & "', responsible_id = '" & responsible_id &_   
						"' WHERE contact_ID="& contactID 
						'Response.Write(sqlUpd)
						'Response.End
						con.ExecuteQuery(sqlUpd)
						
						'--insert into changes table
						sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
						" SELECT 'איש קשר', ' שם:'  + IsNULL(CONTACT_NAME, '') + ' נייד: ' + IsNULL(cellular, ''), CONTACT_ID, 'עדכון', getDate()," & UserID & _
						" FROM dbo.CONTACTS WHERE (CONTACT_ID = "& contactID &")"
						con.executeQuery(sqlstr)  					
										
					else               
						sqlIns="SET DATEFORMAT MDY; SET NOCOUNT ON; INSERT INTO contacts(company_ID,CONTACT_NAME,"&_
						" email,phone,fax,cellular,date_update,messanger_name, contact_address,contact_city_name,"&_
						" contact_zip_code,contact_desc,responsible_id,Organization_ID) VALUES ("&_
						companyID & ", '"& CONTACT_NAME & "','" & email & "','" & phone &_
						"','"& fax & "','" & cellular & "', getDate(),'" & messangerName & "','" & contact_address & "','" &_
						contact_city_name & "','" & contact_zip_code & "','" & contact_desc & "','" & responsible_id & "'," & OrgID & ");"&_
						" SELECT @@IDENTITY AS NewID"
						'Response.Write(sqlIns)
						'Response.End
						set rs_tmp = con.getRecordSet(sqlIns)
							contactID = rs_tmp.Fields("NewID").value
						set rs_tmp = Nothing	
						
						'--insert into changes table
						sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
						" SELECT 'איש קשר', ' שם:'  + IsNULL(CONTACT_NAME, '') + ' נייד: ' + IsNULL(cellular, ''), CONTACT_ID, 'הוספה', getDate()," & UserID & _
						" FROM dbo.CONTACTS WHERE (CONTACT_ID = "& contactID &")"
						con.executeQuery(sqlstr)
					End if     
				End If
			    
				' עדכון סיווגי אנשי הקשר
				sqlstr="Delete FROM contact_to_types WHERE contact_ID = " & contactID
				con.executeQuery(sqlstr)	
				
				If  Len(contact_types) > 0 Then
					arrTypes = Split(contact_types & ",", ",")
					numOfTypes = Ubound(arrTypes)
				End If	
				
				If IsArray(arrTypes) And numOfTypes > 0 Then
					For i=0 To numOfTypes		
						If IsNumeric(arrTypes(i)) Then	
						sqlstr="Insert Into contact_to_types (contact_ID,type_id) values ("&contactID&","&arrTypes(i)&" )"			
						con.executeQuery(sqlstr)	
						End If		
					Next
				End If
			End If
		End If	
		
		If trim(companyID) = "" Or IsNull(companyID) Then
			companyID = "NULL"
		End If	
		If trim(contactID) = "" Or IsNull(contactID) Then
			contactID = "NULL"
		End If	
		If trim(projectID) = "" Or IsNull(projectID) Then
			projectID = "NULL"
		End If
		If trim(mechanismID) = "" Or IsNull(mechanismID) Then
			mechanismID = "NULL"
		End If		
					
		if RelationId>0 then
			sqlstr =  "select DateDiff(d,Appeal_Date,Getdate()) as NDays ,appeal_CountryId from appeals where Appeal_Id="&RelationId
			'response.Write sqlstr
			set NN=con.GetRecordSet(sqlStr)
		If not NN.eof Then
			NDays = NN("NDays")
			CRM_Country= NN("appeal_CountryId")
			'response.write "NDays="&NDays &"<bR>Country_CRM="&Country_CRM
			'response.end
		End If
		Set NN = Nothing
		
		end if
		 If Is_GroupsTour = 1  then
		 	sqlstring="SET DATEFORMAT DMY; SET NOCOUNT ON; " &_		
		" INSERT INTO APPEALS (APPEAL_DATE,USER_ID,CONTACT_ID,COMPANY_ID,PROJECT_ID,MECHANISM_ID, "&_
		" QUESTIONS_ID, appeal_status, appeal_deleted, type_id, User_Id_Order_Owner, ORGANIZATION_ID) " & _
		" VALUES ('" & appeal_date & "'," & UserID & "," & contactID & "," & companyID & "," & projectID & "," &_
		mechanismID & "," & quest_id & ",'" & appeal_status & "',0,2," & UserIdOrderOwner & "," & OrgID   &");"&_
		"SELECT @@IDENTITY AS NewID"

		 else
		sqlstring="SET DATEFORMAT DMY; SET NOCOUNT ON; " &_		
		" INSERT INTO APPEALS (APPEAL_DATE,USER_ID,CONTACT_ID,COMPANY_ID,PROJECT_ID,MECHANISM_ID, "&_
		" QUESTIONS_ID, appeal_status, appeal_deleted, type_id, User_Id_Order_Owner, ORGANIZATION_ID,appeal_CountryId,RelationId,NumberDaysBetweenForms) " & _
		" VALUES ('" & appeal_date & "'," & UserID & "," & contactID & "," & companyID & "," & projectID & "," &_
		mechanismID & "," & quest_id & ",'" & appeal_status & "',0,2," & UserIdOrderOwner & "," & OrgID  & "," & CRM_Country &","& RelationId &",'"& NDays &"');"&_
		"SELECT @@IDENTITY AS NewID"
		'Response.Write(sqlstring)
		end if
		set rs_tmp = con.getRecordSet(sqlstring)
			appID = rs_tmp.Fields("NewID").value
			if quest_id=16735 then
			sqlstring="Update APPEALS set Department_Id="& DepartmentId  &" where Appeal_id="& appID
			con.ExecuteQuery(sqlstring)
			end if
			
		'	response.Write "appID="&appID
	    set rs_tmp = Nothing	
	    if RelationId>0 then
	 		sqlstr="UPDATE APPEALS SET appeal_status=3 where contact_id="& contactID  &" and  appeal_status<>5 "
		con.ExecuteQuery(sqlstr)
		pDate=day(Now()) &"/" & month(Now())&"/" & year(Now())  
		'response.Write pDate
		'response.end
		sqlstr="UPDATE APPEALS SET appeal_status=5 where appeal_id="& RelationId  
		'response.Write sqlstr
		'response.end
		con.ExecuteQuery(sqlstr)
		sqlstr="SET DATEFORMAT DMY;UPDATE APPEALS SET APPEAL_Update_DATE='"& pDate &"' where appeal_id="& RelationId  
	con.ExecuteQuery(sqlstr)
		sqlstr="UPDATE APPEALS SET appeal_status=1 where appeal_id="& appID
			
		con.ExecuteQuery(sqlstr)
		end if
			  If Is_GroupsTour = 1 Then 'הוסף טופס רישום חתום
		'  response.Write "Is_GroupsTour="&Is_GroupsTour
		' response.end
		  		sqlstr="UPDATE APPEALS SET appeal_status=3 where contact_id="& contactID  &" and  appeal_status<>5 "
				con.ExecuteQuery(sqlstr)
				sqlstr="UPDATE APPEALS SET appeal_status=1 where appeal_id="& appID
				con.ExecuteQuery(sqlstr)
		  End If
		
		'--insert into changes table
		if quest_id="17857"  or quest_id="17001"  or   quest_id= "17057" or  quest_id ="16735" then


			'response.Write upl.UserFileName

			if upl.UserFileName<>"" then
			'response.Write(upl.UserFileName)
			'response.end
				str_mappath="../../../download/thanksLetter"
				File_Name=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
					extend = LCase(Mid(File_Name,InstrRev(File_Name,".")+1))
					name_without_extend = LCase(Mid(File_Name,1,Len(File_Name)-Len(extend)-1))
					NewFileName =  "thanksLetter_" & appID
					new_name = NewFileName
					upload = true
					i = 0
					set fs = server.CreateObject("Scripting.FileSystemObject") 
					do while fs.FileExists(Server.MapPath(str_mappath & "/"& new_name & "." & extend ))	
						i =  i + 1
						new_name = NewFileName & "_" & i
					loop
					newFileName = new_name & "." & extend
					'response.Write  newFileName
					'response.end
					set fs = Nothing
				if quest_id="17857"  then	
				
					upl.Form("f1").SaveAs Server.Mappath("../../../download/thanksLetter/") & "/" & newFileName
							sqlstr="UPDATE APPEALS SET File_ThankLetter='"& NewFileName &"' where appeal_id="& appID
					'	response.write sqlstr
					'response.end
								con.ExecuteQuery(sqlstr)

				end if
			end if



			insT=upl.form("Tour_id")
			if insT="" then insT=0

			insD=upl.form("ToursDate_Id")
			if insD ="" then insD=0

			insG=upl.form("Guide_Id")
			if insG ="" then insG=0


				sqlstr="UPDATE APPEALS SET Tour_Id="& insT &",Departure_Id="& insD &",Guide_Id=" & insG  &" where appeal_id="& appID
		'response.Write sqlstr
		'response.end
					con.ExecuteQuery(sqlstr)

		
		end if
		
		
		
		
		sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
		"SELECT 'טופס', 'שם: ' + P.Product_Name + ', איש קשר:'  + IsNULL(CONTACT_NAME, ''), Appeal_Id, 'הוספה', getDate(), " & UserId & _
		" FROM dbo.Appeals A LEFT OUTER JOIN dbo.Products P ON P.Product_Id = A.Questions_id " & _
		" LEFT OUTER JOIN dbo.CONTACTS C ON C.Contact_Id = A.Contact_Id WHERE (Appeal_Id = " & appID & ")	"
		con.executeQuery(sqlstr)	    
	    
	    If ReservationId > 0 And contactID > 0 And  quest_id = 16735 Then 'הוסף טופס רישום חתום
			sqlstr="DELETE FROM contact_to_forms WHERE Reservation_Id = " & ReservationId
			con.executeQuery(sqlstr)	
			
			sqlstr="INSERT INTO contact_to_forms (Contact_Id, Reservation_Id, Appeal_Id) VALUES (" &_
			contactID & "," & ReservationId & "," & appID & ")"			
			con.executeQuery(sqlstr)				
	    End If
	     
	    
	    
		sqlstr="SELECT Field_Id,FIELD_TYPE FROM FORM_FIELD WHERE product_id=" & quest_id & " ORDER BY Field_Order"
		'Response.Write sqlstr
		'Response.End
		set fields=con.GetRecordSet(sqlstr)
		If Not fields.eof Then
			arr_fields = fields.getRows()
		End If
		set fields=nothing
		
		If isArray(arr_fields) Then
		For ff=0 To Ubound(arr_fields,2)
			Field_Id = trim(arr_fields(0,ff))
			Field_Value = trim(upl.Form("field" & Field_Id))		
			'Response.Write Field_Id & " " & Field_Value
			'Response.End						
			sqlstr="INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (" & appID & "," & Field_Id & ",'"& sFix(Field_Value) &"')"
			'Response.Write(sqlstr) & "<br><br>"
			'Response.End
			con.ExecuteQuery sqlstr				
		Next
		End If	
		
		new_app = 1
		
If trim(appID) <> "" Then
		
	''update file
	'''''''''''''''''''
		
		
    sqlstr = "Exec dbo.get_appeals '','','','','" & OrgID & "','','','','','','','" & appID & "'"
	'Response.Write("sqlstr=" & sqlstr)
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
		mechanismId = app("mechanism_id")
		appeal_status = app("appeal_status")
		appeal_status_name = trim(app("appeal_status_name"))
		appeal_status_color = trim(app("appeal_status_color"))
		
		If trim(mechanismId) <> "" Then
			sqlstr = "Select mechanism_Name from mechanism WHERE mechanism_Id = " & mechanismId
			set rs_name = con.getRecordSet(sqlstr)
			If not rs_name.eof Then
				mechanismName = trim(rs_name.Fields(0))
			End If
			set rs_name = Nothing
		End If		
		
		sqlstr =  "Select Langu, Product_Name, PRODUCT_DESCRIPTION, RESPONSIBLE From Products WHERE PRODUCT_ID = " & quest_id
		'Response.Write sqlstr
		'Response.End
		set rsq = con.getRecordSet(sqlstr)
		If not rsq.eof Then		
			Langu = trim(rsq(0))
			prodName = trim(rsq(1))
			PRODUCT_DESCRIPTION = trim(rsq(2))
			RESPONSIBLE = trim(rsq(3))
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
		End If		
		
		If isNumeric(RESPONSIBLE) And Len(RESPONSIBLE) > 0 Then
			sqlstr = "Select EMAIL From USERS Where USER_ID = " & RESPONSIBLE
			set rswrk = con.getRecordSet(sqlstr)
			If not rswrk.eof Then
			  toMail = trim(rswrk("EMAIL"))
			End If
			set rswrk = Nothing
			
			sqlstr = "Select EMAIL From USERS Where USER_ID = " & UserID
			set rswrk = con.getRecordSet(sqlstr)
			If not rswrk.eof Then
			  fromEmail = trim(rswrk("EMAIL"))
			End If
			set rswrk = Nothing		
			
			sqlstr =  "Select Langu, PRODUCT_DESCRIPTION, PRODUCT_NAME From Products WHERE PRODUCT_ID = " & quest_id
			'Response.Write sqlstr
			'Response.End
			set rsq = con.getRecordSet(sqlstr)
			If not rsq.eof Then		
				Langu = trim(rsq(0))
				PRODUCT_DESCRIPTION = trim(rsq(1))
				PRODUCT_NAME = trim(rsq(2))
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
			
			xmlFilePath = "../../../download/xml_forms/bizpower_forms.xml"		
			'----start adding to XML file ------
			set objDOM = Server.CreateObject("Microsoft.XMLDOM")
			objDom.async = false
		
			if objDOM.load(server.MapPath(xmlFilePath)) then
				Set objRoot = objDOM.documentElement
		
				Set objChield = objDOM.createElement("FORM")		
 				set objNewField=objRoot.appendChild(objChield) 		  			
		  
				Set objNewAttribute = objDOM.createAttribute("APPEAL_ID")
				objNewAttribute.text = appID		
 				objNewField.attributes.setNamedItem(objNewAttribute) 
				Set objNewAttribute=nothing
				
				Set objNewAttribute = objDOM.createAttribute("TYPE_ID")
				objNewAttribute.text = "2"
 				objNewField.attributes.setNamedItem(objNewAttribute) 
				Set objNewAttribute=nothing
				
				Set objNewAttribute = objDOM.createAttribute("PRODUCT_NAME")
				objNewAttribute.text = PRODUCT_NAME
 				objNewField.attributes.setNamedItem(objNewAttribute)  	
 				Set objNewAttribute=nothing
 				
 				Set objNewAttribute = objDOM.createAttribute("RESPONSIBLE")
				objNewAttribute.text = RESPONSIBLE
 				objNewField.attributes.setNamedItem(objNewAttribute)  	
 				Set objNewAttribute=nothing				
 		  
				objDom.save server.MapPath(xmlFilePath)
				    
				set objNewField=nothing
		        Set objChield = nothing
	        end if ' objDOM.load	
	        set objDOM = nothing
	
            '----end adding to XML file ------		  
	        If trim(Langu) = "heb" Then		
	            strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & vbCrLf &_
	            "<link href=" & strLocal & "netcom/IE4.css rel=STYLESHEET type=text/css></head><body>" & vbCrLf  &_			
	            "<table border=0 width=450 cellspacing=0 cellpadding=0 align=center bgcolor=#e6e6e6><tr>" & vbCrLf  &_	
                "<td class=title_form width=100% style=""font-size:14pt;line-height:26px"" align=center>" & PRODUCT_NAME &_
                "</td></tr><tr>" & vbCrLf &_	
	            "<td width=100% align=right valign=top  bgcolor=#C9C9C9 style=""border-bottom: 1px solid #808080;border-top: 1px solid #808080;padding-right:15px"">" & vbCrLf  &_
	            "<TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=0><tr>" & vbCrLf &_
	            "<td align=right width=""100%""><span style=""width:110;"" class=""Form_R"" dir=""ltr"">" & appeal_date & "</span></td>" & vbCrLf &_
	            "<td align=right width=130 nowrap class=""form_title"">תאריך הזנה</td></tr>" & vbCrLf &_
	            "<tr><td align=right width=100% ><span style=""width:150;""  class=""Form_R"" dir=""rtl"">" & user_name & "</span></td>" & vbCrLf &_
	            "<td align=right width=130 nowrap class=""form_title"">הזנה ע''י עובד</td></tr>" & vbCrLf &_
	            "<tr><td align=right width=100% ><span class=""status_num"" style=""background-color:" & trim(appeal_status_color) & "; text-align:center"" >" & appeal_status_name & "</span></td>" & vbCrLf &_
	            "<td align=right width=130 nowrap class=""form_title"">סטטוס</td></tr>" & vbCrLf
	            If isNumeric(companyId) Then
		            strBody = strBody & "<tr><td align=right width=100% >" & companyName & "</td>" & vbCrLf &_
		            "<td align=right width=130 nowrap class=""form_title"">קישור ל" & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & "</td></tr>"	& vbCrLf	
	            End If
	            If isNumeric(contactId) Then
		            strBody = strBody & "<tr><td align=right width=100% >" & contactName & "</td>" & vbCrLf &_
	                "<td align=right width=130 nowrap class=""form_title"">קישור ל" & trim(Request.Cookies("bizpegasus")("ContactsOne")) & "</td></tr>" & vbCrLf
	            End If
	            If isNumeric(projectID) And trim(projectName) <> "" Then
		            strBody = strBody & "<tr><td align=right width=100% >" & projectName & "</td>" & vbCrLf &_
	                "<td align=right width=130 nowrap class=""form_title"">קישור ל" & trim(Request.Cookies("bizpegasus")("Projectone")) & "</td></tr>" & vbCrLf	
	            End If
	            If isNumeric(mechanismID) And trim(mechanismName) <> "" Then
		            strBody = strBody & "<tr><td align=right width=100% >" & mechanismName & "</td>" & vbCrLf &_
	                "<td align=right width=130 nowrap class=""form_title"">קישור למנגנון</td></tr>" & vbCrLf
	            End If	
	            strBody = strBody & "<tr><td height=10 colspan=2 nowrap></td></tr></table></td></tr>" & vbCrLf
                strBody = strBody & "<tr><td width=100% colspan=2 align=right><TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=3 bgcolor=White>" & vbCrLf
	
	            Set arr_fields = Nothing
	            sqlstr = "Exec dbo.get_field_value '" & OrgID & "','','"& appId & "','"& quest_id &"',''"
	            set fields=con.GetRecordSet(sqlstr)
	            If Not fields.eof Then
		            arr_fields = fields.getRows()
	            End If
	            set fields=nothing
	
	            If isArray(arr_fields) Then
	                For ff=0 To Ubound(arr_fields,2)
		                Field_Align = trim(arr_fields(0,ff))
		                Field_Title = trim(breaks(arr_fields(5,ff)))
		                Field_Type = trim(arr_fields(6,ff))
		                Field_Value = trim(arr_fields(8,ff))		
			
		                If pr_language = "eng" Then
			                strBody = strBody & "<tr><td dir=rtl "
			                If Field_Type <> "10" Then  'question
				                strBody = strBody & " class=""form"""
			                Else
				                strBody = strBody & " class=""title_form"""
			                End If
			                strBody = strBody & " width=""100%"" align=left bgcolor=#DADADA valign=""top"" style=""padding-left:15px""><b> " & Field_Title & "</b></td></tr>" & vbCrLf
		                ElseIf pr_language = "heb" Then
			                strBody = strBody & "<tr><td dir=rtl "
			                If Field_Type <> "10" then  'question
				                strBody = strBody & " class=""form"""
			                Else
				                strBody = strBody & " class=""title_form"""
			                End If
			                strBody = strBody & " width=""100%"" align=right bgcolor=#DADADA valign=""top"" style=""padding-right:15px""><b> " & Field_Title & "</b></td></tr>" & vbCrLf
		                End If
						
		                If Field_Type <> "10" Then  'question
			                strBody = strBody & "<tr><td width=""85%"" class=""form"" align=" & td_align & " bgcolor=#f0f0f0 style=""padding-right:15px"""
			                If pr_language = "eng" Then
			                    strBody = strBody & " dir=ltr"
			                Else
			                    strBody = strBody & " dir=rtl"
			                End if
			                strBody = strBody & " >&nbsp;"
			                If Field_Type = "1" Or Field_Type = "2" Or Field_Type = "3" Or  Field_Type = "6" Or Field_Type = "7" Or Field_Type = "8" Or Field_Type = "9" Or Field_Type = "11" Or Field_Type = "12" Then
				                strBody = strBody & breaks(Field_Value)				
			                Elseif Field_Type = "5" Then  'check
				                If Field_Value="on" And Field_Align="rtl" Then
					                strBody = strBody & "כן"
				                ElseIf Field_Value="on" And Field_Align="ltr" Then
					                strBody = strBody & "yes"
				                ElseIf Field_Align="rtl" Then
					                strBody = strBody & "לא"
				                Else
					                strBody = strBody & "no"
				                End If					
			                End If
			                strBody = strBody & "&nbsp;</td></tr>" & vbCrLf
		                End If
	                Next
	            End If
	            strBody = strBody & "</TABLE></td></tr>" & vbCrLf  &_
	            "<tr><td align=center colspan=2 bgcolor=white><br><a href='" & strLocal & "netCom/members/appeals/appeal_card.asp?appID="&appID&"&UserID="&RESPONSIBLE&"' target=_blank style='FONT-SIZE:16px;COLOR:#0000FF'><B>לחץ כאן לפרטים</B></a><br></td></tr>" & vbCrLf  &_
	            "</table></td></tr></table>" & vbCrLf
	        Else
	    	    strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & vbCrLf &_
			    "<link href=" & strLocal & "netcom/IE4.css rel=STYLESHEET type=text/css></head><body>" & vbCrLf  &_			
			    "<table border=0 width=450 cellspacing=0 cellpadding=0 align=center bgcolor=#e6e6e6 dir="&dir_align&"><tr>" & vbCrLf  &_	
			    "<td class=title_form width=100% style=""font-size:14pt;line-height:26px"" align=center>" & PRODUCT_NAME &_
			    "</td></tr><tr>" & vbCrLf &_	
			    "<td width=100% align=left valign=top  bgcolor=#C9C9C9 style=""border-bottom: 1px solid #808080;border-top: 1px solid #808080;padding-left:15px"">" & vbCrLf  &_
			    "<TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=0 dir=rtl><tr>" & vbCrLf &_
			    "<td align=left width=100% ><span style=""width:110;"" class=""Form_R"" dir=""ltr"">" & appeal_date & "</span></td>" & vbCrLf &_
			    "<td align=left width=130 nowrap class=""form_title"">Enter date</td></tr>" & vbCrLf &_			
			    "<tr><td align=left width=100% ><span class=""status_num"" style=""background-color:" & trim(appeal_status_color) & "; text-align:center"">" & appeal_status_name & "</span></td>" & vbCrLf &_
			    "<td align=left width=130 nowrap class=""form_title"">Status</td></tr>" & vbCrLf
			    If isNumeric(companyId) Then
				    strBody = strBody & "<tr><td align=left width=100% >" & companyName & "</td>" & vbCrLf &_
				    "<td align=left width=130 nowrap class=""form_title"">Link to company</td></tr>" & vbCrLf		
			    End If
			    If isNumeric(contactId) Then
				    strBody = strBody & "<tr><td align=left width=100% >" & contactName & "</td>" & vbCrLf &_
				    "<td align=left width=130 nowrap class=""form_title"">Link to contact</td></tr>" & vbCrLf
			    End If
			    If isNumeric(projectID) And trim(projectName) <> "" Then
				    strBody = strBody & "<tr><td align=left width=100% >" & projectName & "</td>" & vbCrLf &_
				    "<td align=left width=130 nowrap class=""form_title"">Link to project</td></tr>" & vbCrLf	
			    End If
			    If isNumeric(mechanismID) And trim(mechanismName) <> "" Then
				    strBody = strBody & "<tr><td align=left width=100% >" & mechanismName & "</td>" & vbCrLf &_
				    "<td align=left width=130 nowrap class=""form_title"">Link to Sub-Project</td></tr>" & vbCrLf	
			    End If			
			    strBody = strBody & "<tr><td height=10 colspan=2 nowrap></td></tr></table></td></tr>" & vbCrLf	
			    strBody = strBody & "<tr><td width=100% colspan=2 align=left><TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=3 bgcolor=White>" & vbCrLf
			    sqlstr = "Exec dbo.get_field_value '" & OrgID & "','','"& appId & "','"&quest_id&"',''"
			    set fields=con.GetRecordSet(sqlstr)
			    If Not fields.eof Then
				    arr_fields = fields.getRows()
			    End If
			    set fields=nothing
			
			    If isArray(arr_fields) Then
			        For ff=0 To Ubound(arr_fields,2)
				        Field_Align = trim(arr_fields(0,ff))
				        Field_Title = trim(breaks(arr_fields(5,ff)))
				        Field_Type = trim(arr_fields(6,ff))
				        Field_Value = trim(arr_fields(8,ff))
					
				        If pr_language = "eng" Then
					        strBody = strBody & "<tr><td dir=ltr "
					        If Field_Type <> "10" Then  'question
						        strBody = strBody & " class=""form"""
					        Else
						        strBody = strBody & " class=""title_form"""
					        End If
					        strBody = strBody & " width=""100%"" align=left bgcolor=#DADADA valign=""top"" style=""padding-left:15px""><b> " & Field_Title & "</b></td></tr>" & vbCrLf
				        ElseIf pr_language = "heb" Then
					        strBody = strBody & "<tr><td dir=ltr "
					        If Field_Type <> "10" then  'question
						        strBody = strBody & " class=""form"""
					        Else
						        strBody = strBody & " class=""title_form"""
					        End If
					        strBody = strBody & " width=""100%"" align=left bgcolor=#DADADA valign=""top"" style=""padding-left:15px""><b> " & Field_Title & "</b></td></tr>" & vbCrLf
				        End If
								
				        If Field_Type <> "10" Then  'question
					        strBody = strBody & "<tr><td width=""85%"" class=""form"" align=" & td_align & " bgcolor=#f0f0f0 style=""padding-left:15px"""
					        strBody = strBody & " dir=ltr"
					        strBody = strBody & " >"
					        If Field_Type = "1" Or Field_Type = "2" Or Field_Type = "3" Or  Field_Type = "6" Or Field_Type = "7" Or Field_Type = "8" Or Field_Type = "9" Or Field_Type = "11" Or Field_Type = "12" Then
						        strBody = strBody & breaks(Field_Value)				
					        Elseif Field_Type = "5" Then  'check
						        If Field_Value="on" Then
							        strBody = strBody & "yes"						
						        Else
							        strBody = strBody & "no"
						        End If	
					        End If					
					        strBody = strBody & "</td></tr>" & vbCrLf
				        End If		
		            Next
		        End If
		        strBody = strBody & "</TABLE></td></tr>" & vbCrLf  &_
		        "<tr><td align=center colspan=2 bgcolor=white><br><a href='" & strLocal & "netCom/members/appeals/appeal_card.asp?appID="&appID&"&UserID="&RESPONSIBLE&"' target=_blank style='FONT-SIZE:16px;COLOR:#0000FF'><B>Click here for details</B></a><br></td></tr>" & vbCrLf  &_
		        "</table></td></tr></table>" & vbCrLf
	         End If
	
	        Dim Msg
	        Set Msg = Server.CreateObject("CDO.Message")
		        Msg.BodyPart.Charset = "windows-1255"
		        Msg.From = "support@pegasusisrael.co.il"
		        Msg.MimeFormatted = true
		        Msg.Subject = PRODUCT_NAME & " - BIZPOWER"
		        Msg.To = toMail		
		        Msg.HTMLBody = strBody	
		        Msg.HTMLBodyPart.ContentTransferEncoding = "quoted-printable"		
		        Msg.Send()						
	        Set Msg = Nothing
	        End If
         End If	
	elseif appID <> nil and quest_id <> nil then	        
	    If trim(appeal_status) = "3" Then
			appeal_close_date = " getDate()"
		Else
			appeal_close_date = " NULL"
		End If	
		If trim(companyID) = "" Or IsNull(companyID) Then
			companyID = "NULL"
		End If	
		If trim(contactID) = "" Or IsNull(contactID) Then
			contactID = "NULL"
		End If	
		If trim(projectID) = "" Or IsNull(projectID) Then
			projectID = "NULL"
		End If
		If trim(mechanismID) = "" Or IsNull(mechanismID) Then
			mechanismID = "NULL"
		End If
	    sqlstr="UPDATE APPEALS SET appeal_status = '" & appeal_status & "', appeal_close_date = " & appeal_close_date & ", " & _
	    " COMPANY_ID = " & companyID & ", CONTACT_ID = " & contactID & ", PROJECT_ID = " & projectID & _
	    ", MECHANISM_ID = " & mechanismID & ", User_Id_Update = " & UserId & ", User_Id_Order_Owner = " & UserIdOrderOwner &",Department_Id= "& DepartmentId  & _
	    ", APPEAL_Update_DATE = getDate() " &_
	     " WHERE (ORGANIZATION_ID = " & OrgID & ") AND (appeal_id = " & appID & ")"
		'Response.Write(sqlstr)
		'Response.End 
		  ' ",appeal_CountryId=" & CRM_Country &_
	   ' ",RelationId="& RelationId &_
	
		con.ExecuteQuery(sqlstr)
		'if RelationId<>"" then
		'sqlstr="UPDATE APPEALS SET appeal_status=3 where contact_id="& contactID and  appeal_status<>5 
		'response.Write sqlstr &"<BR>"
		'con.ExecuteQuery(sqlstr)
		'sqlstr="UPDATE APPEALS SET appeal_status=5 where appeal_id="& RelationId
		'response.Write sqlstr &"<BR>"
		'
		'con.ExecuteQuery(sqlstr)
		'sqlstr="UPDATE APPEALS SET appeal_status=1 where appeal_id="& appID
		'response.Write sqlstr &"<BR>"
		'
		'con.ExecuteQuery(sqlstr)
		'end if
		'response.end
		
		'--insert into changes table
		sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
		"SELECT 'טופס', 'שם: ' + P.Product_Name + ', איש קשר:'  + IsNULL(CONTACT_NAME, ''), Appeal_Id, 'עדכון', getDate(), " & UserId & _
		" FROM dbo.Appeals A LEFT OUTER JOIN dbo.Products P ON P.Product_Id = A.Questions_id " & _
		" LEFT OUTER JOIN dbo.CONTACTS C ON C.Contact_Id = A.Contact_Id WHERE (Appeal_Id = " & appID & ")	"
		con.executeQuery(sqlstr)		
		
		sqlstr="SELECT Field_Id,FIELD_TYPE FROM FORM_FIELD Where Product_ID = " & quest_id & " Order by Field_Order"
		set fields=con.GetRecordSet(sqlstr)
		If Not fields.eof Then
			arr_fields = fields.getRows()
		End If
		set fields=nothing
		
		If isArray(arr_fields) Then
		For ff=0 To Ubound(arr_fields,2)
			Field_Id = trim(arr_fields(0,ff))
			Field_Value = trim(upl.Form("field" & Field_Id))
				
			sqlstr="UPDATE FORM_VALUE Set Field_Value='"& Left(sFix(Field_Value),2000) &"' Where Appeal_Id=" & appID & " and Field_Id=" & Field_Id
			'Response.Write(sqlstr)
			con.ExecuteQuery sqlstr				
		Next
		End If
		
		new_app = 0			
End If 'upl.Form("updseker") <> nil 

  	str_mappath="../../../download/documents"
  	
  	If upl.Form("document_name") <> nil And upl.Form("document_file") <> nil  Then
	
	File_Name=Mid(upl.UserFileName,InstrRev(upl.UserFilename,"\")+1)
	extend = LCase(Mid(File_Name,InstrRev(File_Name,".")+1))
	NewFileName =  "appeal_doc_" & appID
	new_name = NewFileName
	upload = true
	i = 0
	set fs = server.CreateObject("Scripting.FileSystemObject") 
	do while fs.FileExists(Server.MapPath(str_mappath & "/"& new_name & "." & extend ))	
		i =  i + 1
		new_name = NewFileName & "_" & i
	loop
	newFileName = new_name & "." & extend
	set fs = Nothing
	if quest_id<>"17857"  then	
	
	upl.Form("document_file").SaveAs Server.Mappath(str_mappath) & "/" & NewFileName
	
	sqlstr = "SET NOCOUNT ON; Insert INTO DOCUMENTS (document_name, document_file, document_desc) VALUES ('"&_
	sFix(trim(upl.Form("document_name")))&"','"&sFix(trim(NewFileName))&"','"&_
	sFix(trim(upl.Form("document_desc")))&"'); SELECT @@IDENTITY AS NewID"
	set rs_tmp = con.getRecordSet(sqlstr)
		document_id = rs_tmp.Fields("NewID").value
	set rs_tmp = Nothing	

	sqlstr = "INSERT INTO appeals_documents (appeal_id, document_id) values (" & appID & "," & document_id & ")"
	con.ExecuteQuery(sqlstr)	
	End If	
	End If	
	Set con = Nothing	%>
<script language="javascript">
<!--	
 		<%	If trim(lang_id) = "1" Then
				str_alert = "הטופס נקלט במערכת"
			Else
				str_alert = "The form was accepted"
			End If
			
			sent_task = trim(upl.Form("send_task")) 
			Set upl = Nothing 	%>
		<%If trim(from) <> "pop_up" Then%>
			<%If new_app = 1 And sent_task = "0" Then%>
			window.alert("<%=str_alert%>");	
			document.location.href="../appeals/appeals.asp?prodID=<%=quest_id%>";
			<%ElseIf new_app = 1 And sent_task = "1" Then%>
			window.open("../tasks/addtask.asp?appealID=<%=appid%>", "AddTask" ,"scrollbars=1,toolbar=0,top=20,left=120,width=430,height=530,align=center,resizable=0");
			document.location.href="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appID=<%=appID%>&companyId=<%=companyId%>&contactID=<%=contactID%>&projectID=<%=projectID%>";
			<%ElseIf  sent_task = "1" Then%>
			window.open("../tasks/addtask.asp?appealID=<%=appid%>", "AddTask" ,"scrollbars=1,toolbar=0,top=20,left=120,width=430,height=530,align=center,resizable=0");
			document.location.href="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appID=<%=appID%>&companyId=<%=companyId%>&contactID=<%=contactID%>&projectID=<%=projectID%>";
			<%Else%>
			window.alert("<%=str_alert%>");	
			
		<%	if contactID<> "" then
	
		%>
document.location.href="../companies/contact.asp?contactID=<%=contactID%>";

			<%else%>

			document.location.href="../appeals/appeals.asp?prodID=<%=quest_id%>";
	<%		end if%>
	//		document.location.href="../appeals/appeals.asp?prodID=<%=quest_id%>";
			<%End If%>
		<%Else%>
			window.close();
				window.opener.document.location.reload(true);
		<%End If%>
//-->
</script>

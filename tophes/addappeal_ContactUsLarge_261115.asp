<!--#INCLUDE FILE="../Netcom/reverse.asp"-->
<!--#include file="../Netcom/connect.asp"-->

<%
Response.Expiresabsolute = Now() - 1 
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache" 

quest_id = trim(Request("quest_id"))
OrgID = trim(Request("OrgID"))
appID = trim(Request("appID"))

UseCaptcha = 0	
If trim(OrgID) <> "" Then
	sqlstring = "SELECT UseCaptcha FROM organizations WHERE ORGANIZATION_ID = " & OrgID
	set rs_tmp = con.GetRecordSet(sqlstring)
	if not rs_tmp.EOF  then				
		UseCaptcha = trim(rs_tmp("UseCaptcha"))
	else
	    UseCaptcha = 0	
	end if
	set rs_tmp = Nothing
	If isNULL(UseCaptcha) Then
		UseCaptcha = 0
	End if		
End If

If (Request.Form("txt_captcha") <> nil) Then
	txt_captcha = trim(Request.Form("txt_captcha"))
Else
	txt_captcha = ""
End If

If ((trim(txt_captcha) = trim(Session("AuthenticationStr")) ) And (trim(txt_captcha) <> "")) Or trim(cStr(UseCaptcha)) = "0" Then

	sizeLogoPr = 0
	If isNumeric(trim(quest_id)) Then
			sqlpg="SELECT PRODUCT_LOGO FROM PRODUCTS WHERE PRODUCT_ID="&quest_id&""
			'Response.Write sqlpg
			'Response.End
			set pg=con.getRecordSet(sqlpg)	
			sizeLogoPr=pg.Fields("PRODUCT_LOGO").ActualSize
			set pg = Nothing
	End If

	sizeLogoOrg = 0
	If trim(OrgID) <> "" And sizeLogoPr = 0 Then
		sqlstring = "SELECT ORGANIZATION_LOGO FROM organizations WHERE ORGANIZATION_ID = " & OrgID
		'Response.Write sqlstring
		'Response.End
		set logo=con.GetRecordSet(sqlstring)
		if not logo.EOF  then				
			sizeLogoOrg = logo("ORGANIZATION_LOGO").ActualSize
		else
			sizeLogoOrg = 0	
		end if
		set logo = Nothing	
	End If		
	If trim(quest_id) <> "" Then
			sqlstr = "Select Product_Name,PRODUCT_DESCRIPTION,PRODUCT_THANKS,Langu, ADD_CLIENT From Products WHERE Product_ID=" & quest_id				
			set rsd = con.getRecordSet(sqlstr)
			If not rsd.eof Then
				Product_Name = rsd(0)
				product_description = rsd(1)
				product_thanks = rsd(2)
				Langu = rsd(3)
				addClient = trim(rsd(4))
				If trim(PRODUCT_THANKS) = "" Then
					if Langu = "eng" then
						PRODUCT_THANKS = "The form has been received"
					else
						PRODUCT_THANKS = "הטופס נקלט במערכת"
					end if
				End If
				
				if Langu = "eng" then
					dir_align = "ltr"
					td_align = "left"
					Langu = "eng"
				else
					dir_align = "rtl"
					td_align = "right"
					Langu = "heb"
				end if		
			End If
			set rsd = Nothing	
	End If	

	pageId = trim(Request.Form("pageId"))
	fromID = trim(Request.Form("fromID"))
	referer = trim(Request.Form("referer"))
	MYreferer = Request.ServerVariables("HTTP_HOST")
						
	if appID = nil and quest_id <> nil then
		companyID = "NULL"
		contactID = "NULL"		
		
		If (trim(addClient) = "1" Or  trim(addClient) = "2") Then
			private_flag = "0"
			If trim(addClient) = "1" Then 'הוספת איש קשר לחברה לקוחות פרטיים
			
				contact_name =  sFix(trim(Request.Form("contact_name")))
				contact_email=trim(sFix(Request.Form("contact_email")))	
				contact_phone=sFix(trim(Request.Form("contact_phone")))								
				contact_cellular=sFix(trim(Request.Form("contact_cellular")))			
				
				sqlstr = "Select company_id From companies WHERE Organization_ID = " & trim(OrgID) &_
				" AND private_flag = 1"
				set rs_comp_pr = con.getRecordSet(sqlstr)	
				If not rs_comp_pr.eof Then
					companyID = rs_comp_pr(0)
				End If
				set rs_comp_pr = Nothing
				
				found_contact_id = 0
				sqlstr = "EXEC dbo.contacts_chk_phone	@OrgId='" & OrgID & "', @EditContactId='0', " & _
				" @cp='" & contact_cellular & "',	@pn='" & contact_phone & "'"	
				set rs_tmp = con.getRecordSet(sqlstr)	
				If rs_tmp.eof = False Then
					found_contact_id = trim(rs_tmp(0))
				End If					
				set rs_tmp = Nothing
				
				If found_contact_id > 0 Then
					contactID = cLng(found_contact_id)
				End If
				
				If IsNumeric(companyID) And IsNumeric(contactID) Then
				'NO UPDATE if the member exists!!!!
					'sqlUpd="UPDATE contacts SET CONTACT_NAME='" & contact_name &_
					'"', email='"& contact_email  &"', phone='"& contact_phone &_
					'"', cellular='"& contact_cellular &"',"&_
					'" date_update=getDate() WHERE contact_ID="& contactID 
					'Response.Write(sqlUpd)
					'Response.End
					'con.ExecuteQuery(sqlUpd)
				ElseIf IsNumeric(companyID) Then
					sqlIns="SET DATEFORMAT MDY; SET NOCOUNT ON; "&_
					" INSERT INTO CONTACTS(company_id,contact_name,email,phone,cellular,date_update,Organization_ID)" &_				  
					" VALUES ("& companyID & ",'" & contact_name & "','" & contact_email & "','"&_
					contact_phone &"','" & contact_cellular & "'" &_
					", getDate(), " & OrgID & "); SELECT @@IDENTITY AS NewID"				
					'Response.Write(sqlIns)
					'Response.End
					set rs_tmp = con.getRecordSet(sqlIns)
						contactID = rs_tmp.Fields("NewID").value
					set rs_tmp = Nothing	 
				End If		
			
			Else  ' הוספת איש קשר וחברה
					company_name=trim(sFix(Request.Form("company_name")))				
					address=trim(sFix(Request.Form("address")))					
					email=trim(sFix(Request.Form("email")))
					cityName=trim(sFix(Request.Form("cityName")))
					zip_code = trim(sFix(Request.Form("zip_code")))
					url=trim(sFix(Request.Form("url")))
					phone1=trim(sFix(Request.Form("phone1")))
					phone2=trim(sFix(Request.Form("phone2")))
					fax1=trim(sFix(Request.Form("fax1")))
					fax2=trim(sFix(Request.Form("fax2")))	
					email=trim(sFix(Request.Form("email")))
					status=trim(Request.Form("status"))
					If trim(status) = "" Then
						status = "4"
					End If	
					company_desc = sFix(Left(trim(Request.Form("company_desc")),500))
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
						fax1 &"','"& fax2 &"','"& zip_code &"','"& cityName &"','" & email & "','" & status & "',"&_
						OrgID & "); SELECT @@IDENTITY AS NewID"
						'Response.Write(sqlIns)
						'Response.End     
						set rs_tmp = con.getRecordSet(sqlIns)
							companyID = rs_tmp.Fields("NewID").value
						set rs_tmp = Nothing	  
					end if					
					
					contact_name=sFix(trim(Request.Form("contact_name")))     				
					contact_phone=sFix(trim(Request.Form("contact_phone")))				
					contact_fax=sFix(trim(Request.Form("contact_fax")))				
					contact_cellular=sFix(trim(Request.Form("contact_cellular")))
					contact_email=sFix(trim(Request.Form("contact_email")))   
					messangerName=sFix(trim(Request.Form("messangerName")))				
					contact_address=trim(sFix(Request.Form("contact_address")))
					contact_city_name=trim(sFix(Request.Form("contact_city_name")))
					contact_zip_code=trim(sFix(Request.Form("contact_zip_code")))
					contact_desc = sFix(Left(trim(Request.Form("contact_desc")),500))
					contact_types = trim(Request.Form("contact_types"))
					
					if IsNumeric(contactID) then 
					'NO UPDATE if the member exists!!!!   		
						'sqlUpd="UPDATE contacts SET CONTACT_NAME='"& CONTACT_NAME &_
						'"', email='"& contact_email  &"',phone='"& contact_phone &"', fax='"& contact_fax &_
						'"', cellular='"& contact_cellular &"', date_update=getDate()" &_
						'", messanger_name='"& messangerName &"', contact_address = '" & contact_address &_
						'"', contact_city_name = '" & contact_city_name & "', contact_zip_code = '" & contact_zip_code &_    
						'"', contact_desc = '" & contact_desc &_    
						'"' WHERE contact_ID="& contactID 
						'Response.Write(sqlUpd)
						'Response.End
						'con.ExecuteQuery(sqlUpd)				
					else               
						sqlIns="SET DATEFORMAT MDY; SET NOCOUNT ON; INSERT INTO contacts("&_
						" company_ID,CONTACT_NAME,email,phone,fax,cellular,date_update,messanger_name,"&_
						" contact_address,contact_city_name,contact_zip_code,contact_desc,Organization_ID) VALUES ("&_
						companyID & ", '"& CONTACT_NAME & "','" & contact_email & "','" & contact_phone &_
						"','"& contact_fax & "','" & contact_cellular & "', getDate(),'" & messangerName & "','" & contact_address & "','" &_
						contact_city_name & "','" & contact_zip_code & "','" & contact_desc & "'," & OrgID & ");"&_
						" SELECT @@IDENTITY AS NewID"
						'Response.Write(sqlIns)
						'Response.End
						set rs_tmp = con.getRecordSet(sqlIns)
							contactID = rs_tmp.Fields("NewID").value
						set rs_tmp = Nothing	
						End if     
					    
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
				
			con.executeQuery("SET DATEFORMAT dmy")
			sqlstring="SET NOCOUNT ON;" &_		
			"INSERT INTO appeals (appeal_date,questions_id,appeal_status,appeal_deleted,type_id,organization_id,company_ID,contact_ID) " &_
			"VALUES (getDate()," & quest_id & ",'1',0,2," & OrgID & "," & companyID & "," & contactID & ")" &_
			"SELECT @@IDENTITY AS NewID"
			'Response.Write(sqlstring)
			'Response.End 
			set rs_tmp = con.getRecordSet(sqlstring)
				appID = rs_tmp.Fields("NewID").value
			set rs_tmp = Nothing			
			
			sqlstr="SELECT Field_Id,FIELD_TYPE  FROM FORM_FIELD Where product_id=" & quest_id & " Order by Field_Order"
			'Response.Write sqlstr
			'Response.End
			set fields=con.GetRecordSet(sqlstr)
				do while not fields.EOF
					Field_Id = fields("Field_Id")				
					Field_Value = trim(Request("field" & Field_Id))		
					'Response.Write Field_Id & " " & Field_Value & "<br>"
					'Response.End						
					sqlstr="INSERT INTO FORM_VALUE (Appeal_Id,Field_Id,Field_Value) VALUES (" & appID & "," & Field_Id & ",'"& sFix(Field_Value) &"')"
					'Response.Write(sqlstr) & "<br><br>"
					'Response.End
					con.ExecuteQuery sqlstr				
				fields.moveNext()
				loop
			set fields=nothing
			'Response.End
			new_app = 1
			
			if Lcase(trim(referer))<>trim(MYreferer) Or isNumeric(pageId) then    
				sqlCount="INSERT INTO Statistic (Page_ID,Product_ID,Category_ID,Appeal_Id,Date,referer) VALUES ('" &_
				pageId & "'," & quest_id & ",'" & fromID & "'," & appID & ", getDate(), '" & sFix(referer) & "')"
				'Response.Write sqlCount
				'Response.End 
				con.getRecordSet(sqlCount)       
			end if			
		
			If trim(appID) <> "" Then
			sqlstr = "EXECUTE get_appeals '','','','','" & OrgID & "','','','','','','','" & appID & "'"
			'Response.Write("sqlstr=" & sqlstr)
			set app = con.GetRecordSet(sqlstr)
			if not app.eof then
			
			appeal_date = Day(app("appeal_date"))& "/" & Month(app("appeal_date")) & "/" & Year(app("appeal_date"))		
			companyName = app("company_Name")
			contactName = app("contact_Name")
			projectName = app("project_Name")
			userName = app("user_Name")
			productName = app("product_Name")
			quest_id = app("questions_id")
			companyId = app("company_id")
			contactId = app("contact_id")
			projectId = app("project_id")
			appeal_status = trim(app("appeal_status"))
			appeal_status_color = trim(app("appeal_status_color"))
			appeal_status_name = trim(app("appeal_status_name"))			
			
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
		
		If isNumeric(RESPONSIBLE) Then
			sqlstr = "Select EMAIL From USERS Where USER_ID = " & RESPONSIBLE
			set rswrk = con.getRecordSet(sqlstr)
			If not rswrk.eof Then
			  toMail = trim(rswrk("EMAIL"))
			  fromEmail = trim(rswrk("EMAIL"))
			End If
			set rswrk = Nothing						
			
			xmlFilePath = "../download/xml_forms/bizpower_forms.xml"		
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
	"<link href=""" & strLocal & "netcom/IE4.css"" rel=""STYLESHEET"" type=""text/css""></head><body>" & vbCrLf  &_			
	"<table border=0 width=450 cellspacing=0 cellpadding=0 align=center bgcolor=#e6e6e6><tr>" & vbCrLf  &_	
    "<td class=title_form width=100% style=""font-size:14pt;line-height:26px"" align=center>" & PRODUCT_NAME &_
    "</td></tr><tr>" & vbCrLf &_	
	"<td width=100% align=right valign=top  bgcolor=#C9C9C9 style=""border-bottom: 1px solid #808080;border-top: 1px solid #808080;padding-right:15px"">" & vbCrLf  &_
	"<TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=0><tr>" & vbCrLf &_
	"<td align=right width=100% ><span style=""width:110;"" class=""Form_R"" dir=""ltr"">" & appeal_date & "</span></td>" & vbCrLf &_
	"<td align=right width=130 nowrap class=""form_title"">תאריך קליטה</td></tr>" & vbCrLf &_
	"<tr><td align=right width=100% ><span class=""status_num"" style=""background-color:" & trim(appeal_status_color) & "; text-align:center"">" & appeal_status_name & "</span></td>" & vbCrLf &_
	"<td align=right width=130 nowrap class=""form_title"">סטטוס</td></tr>" & vbCrLf
	If trim(companyId) <> "" Then
	strBody = strBody & "<tr><td align=right width=100% >" & companyName & "</td>"&_
	"<td align=right width=130 nowrap class=""form_title"">קישור לחברה</td></tr>" & vbCrLf
	End If
	If isNumeric(contactId) Then
	strBody = strBody & "<tr><td align=right width=100% >" & contactName & "</td>" & vbCrLf &_
	"<td align=right width=130 nowrap class=""form_title"">קישור לאיש קשר</td></tr>" & vbCrLf
	End If	
	strBody = strBody & "<tr><td height=10 colspan=2 nowrap></td></tr></table></td></tr>" & vbCrLf
	
    strBody = strBody & "<tr><td width=100% colspan=2 align=right><TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=3 bgcolor=White>" & vbCrLf
	sqlstr = "EXECUTE get_field_value '" & OrgID & "','','"& appId & "','" & quest_id & "',''"
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
		"<tr><td align=center colspan=2 bgcolor=white><br><a href='" & strLocal & "netCom/members/appeals/appeal_card.asp?appID="&appID&"&UserID="&RESPONSIBLE&"' target=_blank style='FONT-SIZE:16px;COLOR:#0000FF'>לחץ כאן לפרטים</a><br></td></tr>" & vbCrLf  &_
		"</table></td></tr></table>" & vbCrLf
	Else
	    	strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & vbCrLf &_
			"<link href=""" & strLocal & "netcom/IE4.css"" rel=""STYLESHEET"" type=""text/css""></head><body>" & vbCrLf  &_			
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
			strBody = strBody & "<tr><td height=10 colspan=2 nowrap></td></tr></table></td></tr>" & vbCrLf	
			strBody = strBody & "<tr><td width=100% colspan=2 align=left><TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=3 bgcolor=White>" & vbCrLf
			sqlstr = "EXECUTE get_field_value '" & OrgID & "','','"& appId & "','"&quest_id&"',''"
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
	    
	    sqlstring="UPDATE appeals set CONTACT_ID = " & contactID & ", COMPANY_ID = " & companyID &_
	    ", PROJECT_ID = " & projectID & " WHERE ORGANIZATION_ID = " & OrgID & " AND appeal_id = " & appID 	
		'Response.Write(sqlstring)
		'Response.End 
		con.ExecuteQuery(sqlstring) 
		
		sqlstr="SELECT Field_Id,FIELD_TYPE  FROM FORM_FIELD Where product_id=" & quest_id & " Order by Field_Order"
		set fields=con.GetRecordSet(sqlstr)
			do while not fields.EOF
				Field_Id = fields("Field_Id")				
				Field_Value = trim(Request("field" & Field_Id))	
				
				sqlstr="UPDATE FORM_VALUE Set Field_Value='"& Left(sFix(Field_Value),2000) &"' Where Appeal_Id=" & appID & " and Field_Id=" & Field_Id
				'Response.Write(sqlstr)
				con.ExecuteQuery sqlstr	
				
			fields.moveNext()
			loop
		set fields=nothing
		
		new_app = 0
			
	end if 'Request.Form("updseker") <> nil

%>
<html>
<head>
<title><%=Product_Name%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link href="../Netcom/IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body  marginwidth="0" marginheight="0" hspace="0" vspace="0" topmargin="0" leftmargin="0" rightmargin="0" style="BACKGROUND-COLOR: transparent;">
<table cellpadding="0" cellspacing="0" width="100%" ID="Table1">
<tr><td  align=center>
<table WIDTH="90%" ALIGN="center" BORDER="0" CELLPADDING="0" cellspacing="1">	
	 <tr><td nowrap height="10"></td></tr>	
	<tr>		
		<td width=100%  valign=top align="right" dir="rtl" style="FONT-FAMILY: Arial; FONT-SIZE: 18pt; FONT-WEIGHT: normal; color: #00437a; "><%=Server.HTMLEncode("אנו מודים לך על פנייתך לחברת פגסוס.") %>
		</td>
	</tr>	
	 <tr><td nowrap height="3"></td></tr>		
	<tr>		
		<td width=100%  valign=top align="right" dir="rtl" style="FONT-FAMILY: Arial; FONT-SIZE: 18pt; FONT-WEIGHT: normal; color: #00437a; "><%=Server.HTMLEncode("נציגנו יצרו עמך קשר בהקדם האפשרי.") %>
		</td>
	</tr>	
	 <tr><td nowrap height="10"></td></tr>			
</table>
</td>
</tr>
</table>
</body></html>
<%Else%>
<script language=javascript>
<!--
			window.alert('טעית בהקלדת התוים מהתמונה, אנא נסה שוב');
			window.history.back();
//-->
</script>
<%End If%>
<%set con=Nothing%>
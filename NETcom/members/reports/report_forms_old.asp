<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
</head>
<body><br><br><br><br>
<H4  align=center> Please wait ... </H4>
<%  
    Function enterFix(word)
		word=replace("" & word,chr(13),Space(1))
		word=replace("" & word,chr(10),Space(1))	 
		enterFix=word
    End Function   
  
  	OrgID	 = trim(Request.Cookies("bizpegasus")("OrgID"))
  	User_ID	 = trim(Request.Cookies("bizpegasus")("UserID"))
  	lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))
  	is_groups = trim(Request.Cookies("bizpegasus")("ISGROUPS"))
		
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"
		align_var = "left"
		dir_obj_var = "ltr"
	Else
		dir_var = "ltr"
		align_var = "right"
		dir_obj_var = "rtl"
	End If
	
	set fs = createobject("scripting.filesystemobject")
  
	set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
	filestring="../../../download/reports/"
	'Response.Write server.mappath(filestring)
	'Response.End
	fs.DeleteFile server.mappath(filestring) & "/*.*" ,False
  
	quest_id = trim(Request("quest_id"))
	start_date = trim(Request("dateStart"))		
	end_date = trim(Request("dateEnd"))		
	UserID = trim(Request("user_id"))
	CompanyID = trim(Request("company_id"))
	ProjectID = trim(Request("project_id"))
	mechanismID = trim(Request("mechanism_id")) 	
		
	If trim(UserID) <> "" Then
		sqlstr = "Select FIRSTNAME + CHAR(32) + LASTNAME FROM USERS WHERE USER_ID = " & UserID
		set rs_user = con.getRecordSet(sqlstr)
		If not rs_user.eof Then
			userName = trim(rs_user(0))
		End If
		set rs_user = Nothing
	End If	 
	
	strTitle = ""
	If trim(CompanyID) <> "" Then
		sqlstr = "Select Company_Name From Companies WHERE Company_ID = " & CompanyID
		set rs_com = con.getRecordSet(sqlstr)
		if not rs_com.eof then
			strTitle = strTitle & "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & ":&nbsp;<font color=""#FF9900"">" & trim(rs_com(0)) & "</font>"
		end if
		set rs_com = nothing
	End If

	If trim(ProjectID) <> "" Then
		sqlstr = "Select Project_Name From Projects WHERE Project_ID = " & ProjectID
		set rs_project = con.getRecordSet(sqlstr)
		if not rs_project.eof then
			strTitle = strTitle & "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("Projectone")) & ":&nbsp;<font color=""#FF9900"">" & trim(rs_project(0)) & "</font>"
		end if
		set rs_project = nothing	
	End If

	If trim(mechanismID) <> "" Then
		sqlstr = "Select mechanism_name, project_name From mechanism Inner Join Projects On mechanism.project_id = projects.project_id WHERE mechanism_id = " & mechanismID
		set rs_mech = con.getRecordSet(sqlstr)
		if not rs_mech.eof then
			If lang_id = 1 Then
				strTitle = strTitle & "<br>&nbsp;מנגנון:&nbsp;<font color=""#FF9900"">" & trim(rs_mech(1)) & " - " & trim(rs_mech(0)) & "</font>"
			Else
				strTitle = strTitle & "<br>&nbsp;Sub-Project:&nbsp;<font color=""#FF9900"">" & trim(rs_mech(1)) & " - " & trim(rs_mech(0)) & "</font>"
			End If			
		end if
		set rs_mech = nothing
	End If				  
  
  filestring="../../../download/reports/export_answers_" & trim(quest_id) & ".xls"
  'Response.Write server.mappath(filestring)
  'Response.End
  set XLSfile=fs.CreateTextFile(server.mappath(filestring)) ' Create text file as excel file
    
  If isNumeric(quest_id) Then
	sqlstr="Select PRODUCT_NAME,ADD_CLIENT From PRODUCTS WHERE PRODUCT_ID = " & quest_id
	set rsn = con.getRecordSet(sqlstr)
	If not rsn.eof Then
		prodName=trim(rsn(0))
		addCl=trim(rsn(1))		
	End If
	set rsn = Nothing
  End If
  
    XLSfile.WriteLine "<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
	XLSfile.WriteLine "<head>"
	XLSfile.WriteLine "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1255"">"
	XLSfile.WriteLine "<meta name=ProgId content=Excel.Sheet>"
	XLSfile.WriteLine "<meta name=Generator content=""Microsoft Excel 9"">"
	XLSfile.WriteLine "<style>"
	XLSfile.WriteLine "<!--table"
	XLSfile.WriteLine "	{mso-displayed-decimal-separator:""\."";"
	XLSfile.WriteLine "	mso-displayed-thousand-separator:""\,"";}"
	XLSfile.WriteLine "@page"
	XLSfile.WriteLine "{ margin:1.0in .75in 1.0in .75in;"
	XLSfile.WriteLine "	mso-header-margin:.5in;"
	XLSfile.WriteLine "	mso-footer-margin:.5in;"
	XLSfile.WriteLine " mso-page-orientation:landscape;}"
	XLSfile.WriteLine "tr"
	XLSfile.WriteLine "	{mso-height-source:auto;}"
	XLSfile.WriteLine "br"
	XLSfile.WriteLine "	{mso-data-placement:same-cell;}"
	XLSfile.WriteLine ".style0"
	XLSfile.WriteLine "	{mso-number-format:General;"
	XLSfile.WriteLine "	vertical-align:bottom;"
	XLSfile.WriteLine "	white-space:nowrap;"
	XLSfile.WriteLine "	mso-rotate:0;"
	XLSfile.WriteLine "	mso-background-source:auto;"
	XLSfile.WriteLine "	mso-pattern:auto;"
	XLSfile.WriteLine "	color:windowtext;"
	XLSfile.WriteLine "	font-size:10.0pt;"
	XLSfile.WriteLine "	font-weight:400;"
	XLSfile.WriteLine "	font-style:normal;"
	XLSfile.WriteLine "	text-decoration:none;"
	XLSfile.WriteLine "	font-family:Arial;"
	XLSfile.WriteLine "	mso-generic-font-family:auto;"
	XLSfile.WriteLine "	mso-font-charset:0;"
	XLSfile.WriteLine "	border:none;"
	XLSfile.WriteLine "	mso-protection:locked visible;"
	XLSfile.WriteLine "	text-align:center;"
	XLSfile.WriteLine "	mso-style-name:Normal;"
	XLSfile.WriteLine "	mso-height-source:auto;"
	XLSfile.WriteLine "	height:20pt;}"
	XLSfile.WriteLine ".xl24"  ' title
	XLSfile.WriteLine "	{mso-style-parent:style0;"
	XLSfile.WriteLine "	font-family:Arial, sans-serif;"
	XLSfile.WriteLine "	font-size:12.0pt;"
	XLSfile.WriteLine "	font-weight:700;"
	XLSfile.WriteLine "	text-align:center;"
	XLSfile.WriteLine "	direction: " & dir_obj_var & ";"
	XLSfile.WriteLine "	white-space:normal;}"
	XLSfile.WriteLine ".xl25" ' title cell
	XLSfile.WriteLine "	{mso-style-parent:style0;"
	XLSfile.WriteLine "	color:#000000;"
	XLSfile.WriteLine "	font-size:10.0pt;"
	XLSfile.WriteLine "	margin-right:5pt;"
	XLSfile.WriteLine "	margin-left:5pt;"
	XLSfile.WriteLine "	font-weight:700;"
	XLSfile.WriteLine "	font-family:Arial, sans-serif;"
	XLSfile.WriteLine "	text-align:center;"
	XLSfile.WriteLine "	background:#FF9900;"
	XLSfile.WriteLine "	border: 0.5pt solid black;"
	XLSfile.WriteLine "	mso-pattern:auto none;"
	XLSfile.WriteLine "	mso-height-source: auto;"	
	XLSfile.WriteLine " white-space:normal;}"
	XLSfile.WriteLine ".xl26"  ' data cells
	XLSfile.WriteLine "	{mso-style-parent:style0;"
	XLSfile.WriteLine "	font-weight:400;"
	XLSfile.WriteLine "	font-size:10.0pt;"
	XLSfile.WriteLine "	margin-right:5pt;"
	XLSfile.WriteLine "	margin-left:5pt;"
	XLSfile.WriteLine "	font-family:Arial, sans-serif;"
	XLSfile.WriteLine "	border: 0.5pt solid black;"
	XLSfile.WriteLine "	mso-height-source: auto;"	
	XLSfile.WriteLine "	vertical-align:top;"
	XLSfile.WriteLine "	text-align:"&align_var&";}"
	XLSfile.WriteLine "-->"
	XLSfile.WriteLine "</style>"
	XLSfile.WriteLine "<!--[if gte mso 9]><xml>"
	XLSfile.WriteLine "<x:ExcelWorkbook>"
	XLSfile.WriteLine "<x:ExcelWorksheets>"
	XLSfile.WriteLine "<x:ExcelWorksheet>"
	XLSfile.WriteLine "<x:Name>" & prodName & "</x:Name>"
	XLSfile.WriteLine "<x:WorksheetOptions>"
	XLSfile.WriteLine "<x:Print>"
	XLSfile.WriteLine "      <x:ValidPrinterInfo/>"
	XLSfile.WriteLine "      <x:PaperSizeIndex>9</x:PaperSizeIndex>"
	XLSfile.WriteLine "      <x:HorizontalResolution>600</x:HorizontalResolution>"
	XLSfile.WriteLine "      <x:VerticalResolution>600</x:VerticalResolution>"
	XLSfile.WriteLine "</x:Print>"
	XLSfile.WriteLine "<x:Selected/>"
	XLSfile.WriteLine "<x:ProtectContents>False</x:ProtectContents>"
	XLSfile.WriteLine "<x:ProtectObjects>False</x:ProtectObjects>"
	XLSfile.WriteLine "<x:ProtectScenarios>False</x:ProtectScenarios>"
	XLSfile.WriteLine "<x:WorksheetOptions>"
	XLSfile.WriteLine "<x:Selected/>"
	If trim(lang_id) = "1" Then
	XLSfile.WriteLine "<x:DisplayRightToLeft/>"
	End If
	XLSfile.WriteLine "<x:Panes>"
	XLSfile.WriteLine "<x:Pane>"
	XLSfile.WriteLine "<x:Number>1</x:Number>"
	XLSfile.WriteLine "<x:ActiveRow>1</x:ActiveRow>"
	XLSfile.WriteLine "</x:Pane>"
	XLSfile.WriteLine "</x:Panes>"
	XLSfile.WriteLine "<x:ProtectContents>False</x:ProtectContents>"
	XLSfile.WriteLine "<x:ProtectObjects>False</x:ProtectObjects>"
	XLSfile.WriteLine "<x:ProtectScenarios>False</x:ProtectScenarios>"
	XLSfile.WriteLine "</x:WorksheetOptions>"
	XLSfile.WriteLine "</x:ExcelWorksheet>"
	XLSfile.WriteLine "</x:ExcelWorksheets>"
	XLSfile.WriteLine "<x:WindowTopX>360</x:WindowTopX>"
	XLSfile.WriteLine "<x:WindowTopY>60</x:WindowTopY>"
	XLSfile.WriteLine "</x:ExcelWorkbook>"
	XLSfile.WriteLine "</xml><![endif]-->"
	XLSfile.WriteLine "</head>"
	XLSfile.WriteLine "<body>"
	XLSfile.WriteLine "<table style='border-collapse:collapse;table-layout:fixed' dir="&dir_var&">"
  
    strLine="<tr style='background:#808080;color:#ffffff'><td colspan=7 class=xl24 style='background:#808080;color:#ffffff'> " &  trim(prodName) & "</td></tr>"
    strLine=strLine&"<tr style='background:#808080;color:#ffffff'><td colspan=7 class=xl24 style='background:#808080;color:#ffffff'> " & start_date & " - " & end_date & strTitle & "</td></tr>"
    XLSfile.writeline strLine 
  
    If trim(addCl) = "" Then
		If trim(lang_id) = "1" Then
		strLine="<tr style='background:#ff9900;mso-height-source:auto;'>"&_
		"<td class=xl25>התקבל</td><td class=xl25>" & trim(Request.Cookies("bizpegasus")("CompaniesOne")) &_
		"</td><td class=xl25>" & trim(Request.Cookies("bizpegasus")("ContactsOne")) & "</td><td class=xl25> " &_
		trim(Request.Cookies("bizpegasus")("Projectone")) & "</td><td class=xl25>מנגנון</td><td class=xl25>עובד</td>"&_
		"<td class=xl25>מספר טופס מלא</td><td class=xl25>נסגר</td><td class=xl25>תוכן סגירה</td>"
		Else
		strLine="<tr style='background:#ff9900;mso-height-source:auto;'><td class=xl25>Date</td>"&_
		"<td class=xl25>" & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & "</td><td class=xl25>" &_
		trim(Request.Cookies("bizpegasus")("ContactsOne")) & "</td><td class=xl25> " & trim(Request.Cookies("bizpegasus")("Projectone")) &_
		"</td><td class=xl25>Sub-Project</td><td class=xl25>Employee</td><td class=xl25>Form ID</td><td class=xl25>Closing date</td>"&_
		"<td class=xl25>Closing content</td>"
		End If
	Else
		If trim(lang_id) = "1" Then
		strLine="<tr style='background:#ff9900;mso-height-source:auto;'><td class=xl25>התקבל</td><td class=xl25>" & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & "</td><td class=xl25>" & trim(Request.Cookies("bizpegasus")("ContactsOne")) & "</td><td class=xl25> " & trim(Request.Cookies("bizpegasus")("Projectone")) & "</td>" &_
		"<td class=xl25>מנגנון</td><td class=xl25>עיסוק</td><td class=xl25>טלפון</td><td class=xl25>טלפון נייד</td><td class=xl25>פקס</td><td class=xl25>Email</td>"&_		
		"<td class=xl25>עובד</td><td class=xl25>מספר טופס מלא</td><td class=xl25>נסגר</td><td class=xl25>תוכן סגירה</td>"
		Else
		strLine="<tr style='background:#ff9900;mso-height-source:auto;'><td class=xl25>Date</td><td class=xl25>" & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & "</td><td class=xl25>" & trim(Request.Cookies("bizpegasus")("ContactsOne")) & "</td><td class=xl25> " & trim(Request.Cookies("bizpegasus")("Projectone")) & "</td>"&_
		"<td class=xl25>Sub-Project</td><td class=xl25>Position</td><td class=xl25>Phone</td><td class=xl25>Cell phone</td><td class=xl25>Fax</td><td class=xl25>Email</td>"&_		
		"<td class=xl25>Employee</td><td class=xl25>Form ID</td><td class=xl25>Closing date</td><td class=xl25>Closing content</td>"
		End If
	End If
		
    
	sqlStr="SELECT FIELD_TITLE FROM FORM_FIELD WHERE product_id = " & quest_id & " AND FIELD_TYPE <> '10' Order By FIELD_ORDER"
	set fields=con.GetRecordSet(sqlStr)
	If not fields.eof Then
		arr_fields = fields.getRows()
	End If
	Set fields = Nothing
	If isArray(arr_fields) Then
	For ff=0 To Ubound(arr_fields,2)			
	strLine=strLine&"<td class=xl25>"&enterFix(trim(arr_fields(0,ff))) & "</td>"
	Next
	End If
	strLine = strLine & "</tr>"
	XLSfile.writeline strLine	'write titles	
	strLine=""

	start_date_ =  Month(start_date) & "/" & Day(start_date) & "/" & Year(start_date)
	end_date_ =  Month(end_date) & "/" & Day(end_date) & "/" & Year(end_date) 

	sql = "Exec dbo.get_appeals '','','','','" & OrgID & "',' APPEAL_DATE DESC','"&start_date_&"','"&end_date_&"','"&CompanyID&"','','"&ProjectID&"','','" & quest_id & "','" & User_ID & "','','"&sFix(userName) & "','" & is_groups & "','" & mechanismID & "'"	
	'Response.Write sql
	'Response.End
	set rsa = con.getRecordSet(sql)
	while not rsa.eof
		appId = trim(rsa("appeal_id"))
		appDate=trim(rsa("appeal_date"))
		If IsDate(appDate) Then
			appDate = Day(appDate) & "/" & Month(appDate) & "/" & Year(appDate)
		Else
			appDate = ""
		End If
		project=trim(rsa("project_name"))
		contact_name=trim(rsa("contact_name"))
		company=trim(rsa("company_name"))
		companyID=trim(rsa("company_id"))
		contactID=trim(rsa("contact_id")) 
		mechanismID=trim(rsa("mechanism_ID"))
		phone = ""
		fax = ""
		email = ""
		If trim(companyID) <> "" Then
			sqlstr = "Select phone1, fax1, email FROM Companies WHERE Company_ID = " & companyID
			set rs_comp = con.getRecordSet(sqlstr)
			If not rs_comp.Eof Then
				phone = trim(rs_comp(0))
				If Len(phone) > 0 Then phone = cStr(phone)
				fax = trim(rs_comp(1))
				If Len(fax) > 0 Then fax = cStr(fax)
				email = trim(rs_comp(2))
				If Len(email) > 0 Then email = cStr(email)
			End If
			set rs_comp = Nothing
		End If
		position = ""
		cellular = ""
		If trim(contactID) <> "" Then
			sqlstr = "Select messanger_name, cellular FROM Contacts WHERE Contact_ID = " & contactID
			set rs_comp = con.getRecordSet(sqlstr)
			If not rs_comp.Eof Then
				position = rs_comp(0)	
				cellular = rs_comp(1)
			End If
			set rs_comp = Nothing
		End If    
		user=trim(rsa("user_name"))
		appeal_status=trim(rsa("appeal_status"))
		appeal_close_date = trim(rsa("appeal_close_date"))
		If IsDate(appeal_close_date) Then
			appeal_close_date = Day(appeal_close_date) & "/" & Month(appeal_close_date) & "/" & Year(appeal_close_date)
		Else
			appeal_close_date = ""
		End If    
		appeal_close_text = trim(rsa("appeal_close_text"))
		If trim(mechanismId) <> "" Then
			sqlstr = "Select mechanism_Name from mechanism WHERE mechanism_Id = " & mechanismId
			set rs_name = con.getRecordSet(sqlstr)
			If not rs_name.eof Then
				mechanismName = trim(rs_name.Fields(0))
			End If
			set rs_name = Nothing
		Else
			mechanismName = ""	
		End If		
   
		If isNumeric(appId) Then
		If trim(addCl) = "" Then
		strLine = "<tr><td class=xl26 x:str>" & appDate & "</td><td class=xl26 x:str>" & company & "</td><td class=xl26 x:str>" & contact_name & "</td><td class=xl26 x:str>" & project & "</td><td class=xl26 x:str>" & mechanismName & "</td><td class=xl26 x:str>" & user & "</td><td class=xl26>" & appId & "</td><td class=xl26 x:str>" & appeal_close_date & "</td><td class=xl26 x:str>" & appeal_close_text & "</td>"
		Else
		strLine = "<tr><td class=xl26 x:str>" & appDate & "</td><td class=xl26 x:str>" & company & "</td><td class=xl26 x:str>" & enterFix(contact_name) & "</td><td class=xl26 x:str>" & enterFix(project) & "</td><td class=xl26 x:str>" & enterFix(mechanismName) & "</td>" &_
		"<td class=xl26 x:str>" & enterFix(position) & "</td><td  class=xl26 x:str>" & trim(phone) & "</td><td  class=xl26 x:str>" & trim(cellular) & "</td><td  class=xl26 x:str>" & trim(fax) & "</td><td class=xl26 x:str>" & trim(email) & "</td>" &_    
		"<td class=xl26 x:str>" & enterFix(user) & "</td><td class=xl26>" & appId & "</td><td class=xl26 x:str>" & enterFix(appeal_close_date) & "</td><td  class=xl26>" & enterFix(appeal_close_text) & "</td>"
		End If
		sqlStr="SELECT FIELD_ID FROM FORM_FIELD WHERE product_id = " & quest_id & " AND FIELD_TYPE <> '10' Order By FIELD_ORDER"
		set fields=con.GetRecordSet(sqlStr)
		If not fields.eof Then
			arr_fields = fields.getRows()
		Else
			arr_fields = Nothing	
		End If
		Set fields = Nothing
		If isArray(arr_fields) Then
		For ff=0 To Ubound(arr_fields,2)
			Field_Id = arr_fields(0,ff)		
			sql = "Exec dbo.get_field_value '" & OrgID & "','"& Field_Id &"','"& appId & "','" & quest_id & "',''"
			'Response.Write sql
			'Response.End
			set rs_value=con.GetRecordSet(sql)
			If not rs_value.EOF Then          
				FIELD_TYPE=rs_value.Fields(6)
				FIELD_TITLE=rs_value.Fields(5)		
				FIELD_VALUE=rs_value.Fields(8)
				If trim(FIELD_TYPE) <> "5" Then
				strLine = strLine & "<td class=xl26 x:str>" & enterFix(trim(FIELD_VALUE)) & "</td>"       
				ElseIf trim(FIELD_VALUE) = "on" Then
				strLine = strLine & "<td class=xl26 x:str>" & enterFix(trim(FIELD_TITLE)) & "</td>"
				Else
				strLine = strLine & "<td class=xl26>&nbsp;</td>"
				End If
			Else
			strLine = strLine & "<td class=xl26>&nbsp;</td>"       
			End If
			set rs_value = Nothing 	    
		Next
		End If
		strLine = strLine & "</tr>"
		XLSfile.writeline strLine  
		strLine=""
		End If
		rsa.moveNext
	Wend
	set rsa = Nothing  
 XLSfile.WriteLine "</TABLE></body></html>"    
 XLSfile.close
 set con=Nothing
 Response.Redirect(filestring) 'Open the excel file in the browser
 set con= Nothing	 
 %>
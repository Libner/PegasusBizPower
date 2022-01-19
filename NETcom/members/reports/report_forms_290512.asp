<% Response.CharSet = "windows-1255"
       Response.Buffer = True 
       Server.ScriptTimeout = 1000  %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
<head></head>
<body>
<div id="div_save" name="div_save" style="position:absolute; left:0px; top:0px; width:100%; height:100%;">  												
  <table height="100%" width="100%" cellspacing="2" cellpadding="2" border="0">
     <tr><td align="center" valign=middle>
     <table dir=ltr border="0" height="100" width="400" cellspacing="1" cellpadding="1">
     <tr><td align="center" valign=middle><font size="5" color=#0066ff face="Arial">אנא המתן בעת טעינת הדוח</font></td></tr>
      </table>
    </td></tr></table>
</div>
</body>
</html>
<!--#include file="code.asp" -->
<%Function enterFix(word)
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
		dir_var = "rtl" :  align_var = "left" : dir_obj_var = "ltr"
	Else
		dir_var = "ltr" : align_var = "right" : dir_obj_var = "rtl"
	End If	
  
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
	
	Dim strTitle
	Set strTitle = New StrConCatArray
	
	If trim(CompanyID) <> "" Then
		sqlstr = "Select Company_Name From Companies WHERE Company_ID = " & CompanyID
		set rs_com = con.getRecordSet(sqlstr)
		if not rs_com.eof then
			strTitle.Add "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & ":&nbsp;<font color=""#FF9900"">" & trim(rs_com(0)) & "</font>"
		end if
		set rs_com = nothing
	End If

	If trim(ProjectID) <> "" Then
		sqlstr = "Select Project_Name From Projects WHERE Project_ID = " & ProjectID
		set rs_project = con.getRecordSet(sqlstr)
		if not rs_project.eof then
			strTitle.Add "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("Projectone")) & ":&nbsp;<font color=""#FF9900"">" & trim(rs_project(0)) & "</font>"
		end if
		set rs_project = nothing	
	End If

	If trim(mechanismID) <> "" Then
		sqlstr = "Select mechanism_name, project_name From mechanism Inner Join Projects On mechanism.project_id = projects.project_id WHERE mechanism_id = " & mechanismID
		set rs_mech = con.getRecordSet(sqlstr)
		if not rs_mech.eof then
			If lang_id = 1 Then
				strTitle.Add "<br>&nbsp;מנגנון:&nbsp;<font color=""#FF9900"">" & trim(rs_mech(1)) & " - " & trim(rs_mech(0)) & "</font>"
			Else
				strTitle.Add "<br>&nbsp;Sub-Project:&nbsp;<font color=""#FF9900"">" & trim(rs_mech(1)) & " - " & trim(rs_mech(0)) & "</font>"
			End If			
		end if
		set rs_mech = nothing
	End If				  
    
	If isNumeric(quest_id) Then
		sqlstr="Select PRODUCT_NAME,ADD_CLIENT From PRODUCTS WHERE PRODUCT_ID = " & quest_id
		set rsn = con.getRecordSet(sqlstr)
		If not rsn.eof Then
			prodName=trim(rsn(0))
			addCl=trim(rsn(1))		
		End If
		set rsn = Nothing
	End If
  
    Dim strXLS
	Set strXLS = New StrConCatArray
  
    strXLS.Add  "<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
	strXLS.Add  "<head>"
	strXLS.Add  "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1255"">"
	strXLS.Add  "<meta name=ProgId content=Excel.Sheet>"
	strXLS.Add  "<meta name=Generator content=""Microsoft Excel 9"">"
	strXLS.Add  "<style>"
	strXLS.Add  "<!--table"
	strXLS.Add  "	{mso-displayed-decimal-separator:""\."";"
	strXLS.Add  "	mso-displayed-thousand-separator:""\,"";}"
	strXLS.Add  "@page"
	strXLS.Add  "{ margin:1.0in .75in 1.0in .75in;"
	strXLS.Add  "	mso-header-margin:.5in;"
	strXLS.Add  "	mso-footer-margin:.5in;"
	strXLS.Add  " mso-page-orientation:landscape;}"
	strXLS.Add  "tr"
	strXLS.Add  "	{mso-height-source:auto;}"
	strXLS.Add  "br"
	strXLS.Add  "	{mso-data-placement:same-cell;}"
	strXLS.Add  ".style0"
	strXLS.Add  "	{mso-number-format:General;"
	strXLS.Add  "	vertical-align:bottom;"
	strXLS.Add  "	white-space:nowrap;"
	strXLS.Add  "	mso-rotate:0;"
	strXLS.Add  "	mso-background-source:auto;"
	strXLS.Add  "	mso-pattern:auto;"
	strXLS.Add  "	color:windowtext;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	font-weight:400;"
	strXLS.Add  "	font-style:normal;"
	strXLS.Add  "	text-decoration:none;"
	strXLS.Add  "	font-family:Arial;"
	strXLS.Add  "	mso-generic-font-family:auto;"
	strXLS.Add  "	mso-font-charset:0;"
	strXLS.Add  "	border:none;"
	strXLS.Add  "	mso-protection:locked visible;"
	strXLS.Add  "	text-align:center;"
	strXLS.Add  "	mso-style-name:Normal;"
	strXLS.Add  "	mso-height-source:auto;"
	strXLS.Add  "	height:20pt;}"
	strXLS.Add  ".xl24"  ' title
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	font-size:12.0pt;"
	strXLS.Add  "	font-weight:700;"
	strXLS.Add  "	text-align:center;"
	strXLS.Add  "	direction: " & dir_obj_var & ";"
	strXLS.Add  "	white-space:normal;}"
	strXLS.Add  ".xl25" ' title cell
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	color:#000000;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	font-weight:700;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	text-align:center;"
	strXLS.Add  "	background:#FF9900;"
	strXLS.Add  "	border: 0.5pt solid black;"
	strXLS.Add  "	mso-pattern:auto none;"
	strXLS.Add  "	mso-height-source: auto;"	
	strXLS.Add  " white-space:normal;}"
	strXLS.Add  ".xl26"  ' data cells
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	font-weight:400;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	border: 0.5pt solid black;"
	strXLS.Add  "	mso-height-source: auto;"	
	strXLS.Add  "	vertical-align:top;"
	strXLS.Add  "	text-align:"&align_var&";}"
	strXLS.Add  "-->"
	strXLS.Add  "</style>"
	strXLS.Add  "<!--[if gte mso 9]><xml>"
	strXLS.Add  "<x:ExcelWorkbook>"
	strXLS.Add  "<x:ExcelWorksheets>"
	strXLS.Add  "<x:ExcelWorksheet>"
	strXLS.Add  "<x:Name>" & prodName & "</x:Name>"
	strXLS.Add  "<x:WorksheetOptions>"
	strXLS.Add  "<x:Print>"
	strXLS.Add  "      <x:ValidPrinterInfo/>"
	strXLS.Add  "      <x:PaperSizeIndex>9</x:PaperSizeIndex>"
	strXLS.Add  "      <x:HorizontalResolution>600</x:HorizontalResolution>"
	strXLS.Add  "      <x:VerticalResolution>600</x:VerticalResolution>"
	strXLS.Add  "</x:Print>"
	strXLS.Add  "<x:Selected/>"
	strXLS.Add  "<x:ProtectContents>False</x:ProtectContents>"
	strXLS.Add  "<x:ProtectObjects>False</x:ProtectObjects>"
	strXLS.Add  "<x:ProtectScenarios>False</x:ProtectScenarios>"
	strXLS.Add  "<x:WorksheetOptions>"
	strXLS.Add  "<x:Selected/>"
	If trim(lang_id) = "1" Then
	strXLS.Add  "<x:DisplayRightToLeft/>"
	End If
	strXLS.Add  "<x:Panes>"
	strXLS.Add  "<x:Pane>"
	strXLS.Add  "<x:Number>1</x:Number>"
	strXLS.Add  "<x:ActiveRow>1</x:ActiveRow>"
	strXLS.Add  "</x:Pane>"
	strXLS.Add  "</x:Panes>"
	strXLS.Add  "<x:ProtectContents>False</x:ProtectContents>"
	strXLS.Add  "<x:ProtectObjects>False</x:ProtectObjects>"
	strXLS.Add  "<x:ProtectScenarios>False</x:ProtectScenarios>"
	strXLS.Add  "</x:WorksheetOptions>"
	strXLS.Add  "</x:ExcelWorksheet>"
	strXLS.Add  "</x:ExcelWorksheets>"
	strXLS.Add  "<x:WindowTopX>360</x:WindowTopX>"
	strXLS.Add  "<x:WindowTopY>60</x:WindowTopY>"
	strXLS.Add  "</x:ExcelWorkbook>"
	strXLS.Add  "</xml><![endif]-->"
	strXLS.Add  "</head>"
	strXLS.Add  "<body>"
	strXLS.Add  "<table style='border-collapse:collapse;table-layout:fixed' dir=" : strXLS.Add dir_var  : strXLS.Add ">"  
    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=7 class=xl24 style='background:#808080;color:#ffffff'> " &  trim(prodName) & "</td></tr>"
    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=7 class=xl24 style='background:#808080;color:#ffffff'> " & start_date & " - " & end_date & strTitle.value & "</td></tr>"
  
    If trim(addCl) = "" Then
		If trim(lang_id) = "1" Then
		strXLS.Add  "<tr style='background:#ff9900;mso-height-source:auto;'><td class=xl25>התקבל</td><td class=xl25>" 
		strXLS.Add trim(Request.Cookies("bizpegasus")("CompaniesOne"))  : strXLS.Add "</td><td class=xl25>"
		strXLS.Add trim(Request.Cookies("bizpegasus")("ContactsOne"))  : strXLS.Add "</td><td class=xl25> "
		strXLS.Add	trim(Request.Cookies("bizpegasus")("Projectone"))  : strXLS.Add "</td><td class=xl25>מנגנון</td>"
		strXLS.Add	"<td class=xl25>הוזן ע''י</td>"		
		If trim(quest_id) = "16735" Then
			strXLS.Add	"<td class=xl25>למי שייך הרישום</td>"		
		End If
		strXLS.Add	"<td class=xl25>מספר טופס מלא</td><td class=xl25>נסגר</td><td class=xl25>תוכן סגירה</td>"
		Else
		strXLS.Add "<tr style='background:#ff9900;mso-height-source:auto;'><td class=xl25>Date</td><td class=xl25>"
		strXLS.Add  trim(Request.Cookies("bizpegasus")("CompaniesOne"))  : strXLS.Add "</td><td class=xl25>"
		strXLS.Add trim(Request.Cookies("bizpegasus")("ContactsOne"))   : strXLS.Add "</td><td class=xl25> " 
		strXLS.Add trim(Request.Cookies("bizpegasus")("Projectone")) : strXLS.Add	"</td><td class=xl25>Sub-Project</td>"
		strXLS.Add "<td class=xl25>Employee</td><td class=xl25>Form ID</td><td class=xl25>Closing date</td>"
		strXLS.Add "<td class=xl25>Closing content</td>"
		End If
	Else
		If trim(lang_id) = "1" Then
		strXLS.Add  "<tr style='background:#ff9900;mso-height-source:auto;'><td class=xl25>התקבל</td><td class=xl25>" 
		strXLS.Add trim(Request.Cookies("bizpegasus")("CompaniesOne"))  : strXLS.Add "</td><td class=xl25>"
		strXLS.Add trim(Request.Cookies("bizpegasus")("ContactsOne"))  : strXLS.Add "</td><td class=xl25> "
		strXLS.Add trim(Request.Cookies("bizpegasus")("Projectone"))  : strXLS.Add "</td><td class=xl25>מנגנון</td>"
		strXLS.Add  "<td class=xl25>עיסוק</td><td class=xl25>כתובת</td><td class=xl25>עיר</td><td class=xl25>מיקוד</td>"
		strXLS.Add  "<td class=xl25>טלפון</td><td class=xl25>טלפון נייד</td><td class=xl25>פקס</td><td class=xl25>Email</td>"
		strXLS.Add  "<td class=xl25>סיווג</td><td class=xl25>פרטים נוספים</td><td class=xl25>אחראי לקוח</td>"	
		strXLS.Add  "<td class=xl25>הוזן ע''י</td>"		
		If trim(quest_id) = "16735" Then
			strXLS.Add "<td class=xl25>למי שייך הרישום</td>"		
		End If
		strXLS.Add  "<td class=xl25>מספר טופס מלא</td><td class=xl25>נסגר</td><td class=xl25>תוכן סגירה</td>"
		Else
		strXLS.Add  "<tr style='background:#ff9900;mso-height-source:auto;'><td class=xl25>Date</td><td class=xl25>"
		strXLS.Add trim(Request.Cookies("bizpegasus")("CompaniesOne"))  : strXLS.Add "</td><td class=xl25>"
		strXLS.Add trim(Request.Cookies("bizpegasus")("ContactsOne"))  : strXLS.Add "</td><td class=xl25>"
		strXLS.Add trim(Request.Cookies("bizpegasus")("Projectone"))  : strXLS.Add "</td><td class=xl25>Sub-Project</td>"
		strXLS.Add  "<td class=xl25>Position</td><td class=xl25>Address</td><td class=xl25>City</td><td class=xl25>Zip</td>"
		strXLS.Add  "<td class=xl25>Phone</td><td class=xl25>Cell phone</td><td class=xl25>Fax</td><td class=xl25>Email</td>"
		strXLS.Add  "<td class=xl25>Groups</td><td class=xl25>Description</td><td class=xl25>Responsible</td>"
		strXLS.Add  "<td class=xl25>Employee</td><td class=xl25>Form ID</td><td class=xl25>Closing date</td><td class=xl25>Closing content</td>"
		End If
	End If
		
    Dim strTable
    strTable = ""
	sqlStr="SELECT FIELD_TITLE FROM FORM_FIELD WHERE product_id = " & quest_id & " AND FIELD_TYPE <> '10' Order By FIELD_ORDER"
	set rs_fields=con.GetRecordSet(sqlStr)
	If not rs_fields.eof Then		
		strTable = rs_fields.GetString(,,,"</td><td class=xl25>","&nbsp;")
	End If
	Set rs_fields = Nothing		
	
	strXLS.Add  "<td class=xl25>"  : strXLS.Add  strTable  : strXLS.Add "</td></tr>"  : strXLS.Add vbCrlf	

	sql = "Exec dbo.get_appeals_report @OrgID=" & nFix(OrgID) & ", @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @companyID=" & nFix(CompanyID) & ", @projectID=" & nFix(ProjectID) & ", @productID=" & nFix(quest_id) & "," & _
	" @UserID=" & nFix(User_ID) & ", @user_name='" & sFix(userName) & "', @IsGroups='" & is_groups & "', @mechanismID=" & nFix(mechanismID)
	'Response.Write sql
	'Response.End
	set rsa = con.getRecordSet(sql)
	If Not rsa.Eof Then
		arr_app = rsa.getRows()
	End If
	Set rsa = Nothing	
	
	If isArray(arr_app) Then
	For aa=0 To Ubound(arr_app,2)
		appId = trim(arr_app(0,aa))
		appDate=trim(arr_app(1,aa))
		If IsDate(appDate) Then
			appDate = Day(appDate) & "/" & Month(appDate) & "/" & Year(appDate)
		Else
			appDate = ""
		End If
		project=trim(arr_app(2,aa))		
		company=trim(arr_app(3,aa))
		contact_name=trim(arr_app(4,aa))
		phone=trim(arr_app(5,aa))
		fax = trim(arr_app(6,aa))	
		cellular = trim(arr_app(7,aa))
		email = trim(arr_app(8,aa))		
		position = trim(arr_app(9,aa))
		contact_address = trim(arr_app(10,aa))
		contact_city = trim(arr_app(11,aa))
		contact_zip_code = trim(arr_app(12,aa))
		contact_types = trim(arr_app(13,aa))
		contact_responsible = trim(arr_app(14,aa))
		contact_desc = trim(arr_app(15,aa))
		user=trim(arr_app(16,aa))
		appeal_status=trim(arr_app(17,aa))
		appeal_close_date = trim(arr_app(18,aa))
		If IsDate(appeal_close_date) Then
			appeal_close_date = Day(appeal_close_date) & "/" & Month(appeal_close_date) & "/" & Year(appeal_close_date)
		Else
			appeal_close_date = ""
		End If    
		appeal_close_text = trim(arr_app(19,aa))
		mechanismName = trim(arr_app(20,aa))		
		OrderOwner = trim(arr_app(21,aa))		
   
		If isNumeric(appId) Then
		If trim(addCl) = "" Then
		strXLS.Add  "<tr><td class=xl26 x:str>"  : strXLS.Add appDate  : strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add company 
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add contact_name  : strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add project 
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add mechanismName  : strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add user : strXLS.Add "</td>"
		If trim(quest_id) = "16735" Then
			strXLS.Add "<td class=xl26 x:str>"  :  strXLS.Add OrderOwner   : strXLS.Add "</td>"
		End If		
		strXLS.Add "<td class=xl26>"  : strXLS.Add appId  : strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add appeal_close_date 
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add appeal_close_text  : strXLS.Add "</td>"
		Else
		strXLS.Add "<tr><td class=xl26 x:str>"  : strXLS.Add appDate  : strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add company  : strXLS.Add "</td>"
		strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add enterFix(contact_name)  : strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add enterFix(project)  : strXLS.Add "</td>"
		strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add enterFix(mechanismName)  : strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add enterFix(position)  : strXLS.Add "</td>"
		strXLS.Add "<td  class=xl26 x:str>"  : strXLS.Add contact_address  : strXLS.Add "</td><td  class=xl26 x:str>"  : strXLS.Add contact_city  : strXLS.Add "</td><td  class=xl26 x:str>"  : strXLS.Add contact_zip_code  : strXLS.Add "</td>"
		strXLS.Add "<td  class=xl26 x:str>"  : strXLS.Add phone  : strXLS.Add "</td><td  class=xl26 x:str>"  : strXLS.Add cellular  : strXLS.Add "</td><td  class=xl26 x:str>"  : strXLS.Add fax  : strXLS.Add "</td>"
		strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add email  : strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add contact_types  : strXLS.Add "</td>"
		strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add contact_desc  : strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add contact_responsible  : strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add enterFix(user)  : strXLS.Add "</td>"
		If trim(quest_id) = "16735" Then
			strXLS.Add "<td class=xl26 x:str>"  :  strXLS.Add OrderOwner   : strXLS.Add "</td>"
		End If
		strXLS.Add "<td class=xl26>"  : strXLS.Add appId  : strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add enterFix(appeal_close_date)  : strXLS.Add "</td>"
		strXLS.Add "<td  class=xl26>"  : strXLS.Add enterFix(appeal_close_text)  : strXLS.Add "</td>"
		End If
		sqlStr="Exec dbo.get_appeal_value '" & OrgID & "','" & appId & "','" & quest_id & "'"
		set rs_fields=con.GetRecordSet(sqlStr)
		If not rs_fields.eof Then		
			strTable = rs_fields.GetString(,,,"</td><td class=xl26 x:str>","&nbsp;")
		End If
		Set rs_fields = Nothing
		strXLS.Add  "<td class=xl26 x:str>" :  strXLS.Add strTable  : strXLS.Add "</td></tr>"  : strXLS.Add vbCrlf	
		End If
	Next 
	End If
    strXLS.Add  "</TABLE></body></html>"    
	Set con = Nothing	 
	
	filestring = "export_forms_" & quest_id & ".xls"
    Response.Clear()
    Response.AddHeader "content-disposition", "inline; filename=" & filestring
    Response.ContentType = "application/vnd.ms-excel"
    Response.Write(strXLS.value)  
    Set strXLS = Nothing %>
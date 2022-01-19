<% Response.CharSet = "windows-1255"
       Response.Buffer = True 
       Server.ScriptTimeout = 1000  %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
<head></head>
<body>
<div id="div_save" name="div_save" style="position:absolute; left:0px; top:0px; width:100%; height:100%;">  												
  <table height="100%" width="100%" cellspacing="2" cellpadding="2" border="0" ID="Table1">
     <tr><td align="center" valign=middle>
     <table dir=ltr border="0" height="100" width="400" cellspacing="1" cellpadding="1" ID="Table2">
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
  
	quest_id = 16735
	start_date = trim(Request("dateStart"))		
	end_date = trim(Request("dateEnd"))	
'	response.Write 	start_date &":"& end_date
'	response.end
'	UserID = trim(Request("user_id"))
'	CompanyID = trim(Request("company_id"))
'	ProjectID = trim(Request("project_id"))
'	mechanismID = trim(Request("mechanism_id")) 	
		
'	If trim(UserID) <> "" Then
'		sqlstr = "Select FIRSTNAME + CHAR(32) + LASTNAME FROM USERS WHERE USER_ID = " & UserID
'		set rs_user = con.getRecordSet(sqlstr)
'		If not rs_user.eof Then
'			userName = trim(rs_user(0))
'		End If
'		set rs_user = Nothing
'	End If	 
	
	Dim strTitle
	Set strTitle = New StrConCatArray
	
	'If trim(CompanyID) <> "" Then
	'	sqlstr = "Select Company_Name From Companies WHERE Company_ID = " & CompanyID
	'	set rs_com = con.getRecordSet(sqlstr)
	'	if not rs_com.eof then
	'		strTitle.Add "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & ":&nbsp;<font color=""#FF9900"">" & trim(rs_com(0)) & "</font>"
	'	end if
	'	set rs_com = nothing
	'End If

'	If trim(ProjectID) <> "" Then
'		sqlstr = "Select Project_Name From Projects WHERE Project_ID = " & ProjectID
'		set rs_project = con.getRecordSet(sqlstr)
'		if not rs_project.eof then
'			strTitle.Add "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("Projectone")) & ":&nbsp;<font color=""#FF9900"">" & trim(rs_project(0)) & "</font>"
'		end if
'		set rs_project = nothing	
'	End If

'	If trim(mechanismID) <> "" Then
'		sqlstr = "Select mechanism_name, project_name From mechanism Inner Join Projects On mechanism.project_id = projects.project_id WHERE mechanism_id = " & mechanismID
'		set rs_mech = con.getRecordSet(sqlstr)
'		if not rs_mech.eof then
'			If lang_id = 1 Then
'				strTitle.Add "<br>&nbsp;מנגנון:&nbsp;<font color=""#FF9900"">" & trim(rs_mech(1)) & " - " & trim(rs_mech(0)) & "</font>"
'			Else
'				strTitle.Add "<br>&nbsp;Sub-Project:&nbsp;<font color=""#FF9900"">" & trim(rs_mech(1)) & " - " & trim(rs_mech(0)) & "</font>"
'			End If			
'		end if
'		set rs_mech = nothing
'	End If				  
    
	If isNumeric(quest_id) Then
		sqlstr="Select PRODUCT_NAME,ADD_CLIENT From PRODUCTS WHERE PRODUCT_ID = " & quest_id
		set rsn = con.getRecordSet(sqlstr)
		If not rsn.eof Then
			prodName=trim(rsn(0))
			addCl=trim(rsn(1))		
		End If
		set rsn = Nothing
	End If
 ' response.Write prodName
'  response.end
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
    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=7 class=xl24 style='background:#808080;color:#ffffff'>  דוח מסכם - " &  trim(prodName) & "</td></tr>"
    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=7 class=xl24 style='background:#808080;color:#ffffff'> " & start_date & " - " & end_date & strTitle.value & "</td></tr>"

		If trim(lang_id) = "1" Then
		strXLS.Add  "<tr style='background:#ff9900;mso-height-source:auto;'><td class=xl25>שם מלא</td><td class=xl25>" 
		strXLS.Add "טלפון1 </td><td class=xl25>"
		strXLS.Add "טלפון2</td><td class=xl25> "
		strXLS.Add	"טלפון3</td><td class=xl25>האם נשלח ביטוח </td>"
		strXLS.Add	"<td class=xl25>נשלח ע''י</td>"		
		strXLS.Add	"<td class=xl25>תאריך שנשלח</td>"		
		strXLS.Add	"<td class=xl25>שעה שנשלח</td></tr>"		
		
		End If
	
   ' Dim strTable
   ' strTable = ""
	'sqlStr="SELECT FIELD_TITLE FROM FORM_FIELD WHERE product_id = " & quest_id & " AND FIELD_TYPE <> '10' Order By FIELD_ORDER"
	'set rs_fields=con.GetRecordSet(sqlStr)
	'If not rs_fields.eof Then		
'		strTable = rs_fields.GetString(,,,"</td><td class=xl25>","&nbsp;")
'	End If
'	Set rs_fields = Nothing		
'	
'	strXLS.Add  "<td class=xl25>"  : strXLS.Add  strTable  : strXLS.Add "</td></tr>"  : strXLS.Add vbCrlf	
'
sumInsYes=0
sumInsNo=0
	sql = "Exec dbo.get_Insurance_reportExcel  @start_date='" & start_date & "', @end_date='" & end_date &"'"
	'Response.Write sql
	'Response.End
	set rsa = con.getRecordSet(sql)
	If Not rsa.Eof Then
		arr_app = rsa.getRows()
	End If
	Set rsa = Nothing	
	If isArray(arr_app) Then
	For aa=0 To Ubound(arr_app,2)
	contactName=trim(arr_app(2,aa))
	phone1=trim(arr_app(3,aa))
	phone2=trim(arr_app(4,aa))
	phone3=trim(arr_app(5,aa))
	if IsDate(trim(arr_app(6,aa))) then
		Insurancedate=FormatDateTime(trim(arr_app(6,aa)),2) 
		Insurancetime=FormatDateTime(trim(arr_app(6,aa)),4) 
		else
Insurancedate=""
Insurancetime=""
	end if
	
	if trim(arr_app(7,aa))="1" then
InsuranceStatus="כן"
sumInsYes=sumInsYes+1
	else
	InsuranceStatus="לא"
	sumInsNo=sumInsNo+1
	end if
	InsuranceWorkerName=trim(arr_app(9,aa))
	strXLS.Add "<TR><td class=xl26>"  : strXLS.Add contactName  : strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add phone1 
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add phone2  : strXLS.Add "</td>"
		strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add phone3  : strXLS.Add "</td>"
	strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add InsuranceStatus  : strXLS.Add "</td>"
strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add InsuranceWorkerName  : strXLS.Add "</td>"

strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add Insurancedate  : strXLS.Add "</td>"
strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add Insurancetime  : strXLS.Add "</td></tr>"
	
		
	Next 
	End If
	strXLS.Add "<TR><Td colspan=8 align=right class=xl25 >סה``כ נשלח ביטוח " : strXLS.Add sumInsYes :strXLS.Add "</td></tr>"
	strXLS.Add "<TR><Td colspan=8 align=right class=xl25 >סה``כ לא נשלח ביטוח " : strXLS.Add sumInsNo :strXLS.Add "</td></tr>"
 
    strXLS.Add  "</TABLE></body></html>"    
	Set con = Nothing	 

	
	filestring = "export_forms_" & quest_id & ".xls"
    Response.Clear()
    Response.AddHeader "content-disposition", "inline; filename=" & filestring
    Response.ContentType = "application/vnd.ms-excel"
    Response.Write(strXLS.value)  
    Set strXLS = Nothing %>
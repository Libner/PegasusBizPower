<% Response.CharSet = "windows-1255"
       Response.Buffer = True 
       Server.ScriptTimeout = 1000  %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
	<head>
	</head>
	<body>
		<div id="div_save" name="div_save" style="position:absolute; left:0px; top:0px; width:100%; height:100%;">
			<table height="100%" width="100%" cellspacing="2" cellpadding="2" border="0" ID="Table1">
				<tr>
					<td align="center" valign="middle">
						<table dir="ltr" border="0" height="100" width="400" cellspacing="1" cellpadding="1" ID="Table2">
							<tr>
								<td align="center" valign="middle"><font size="5" color="#0066ff" face="Arial">אנא המתן 
										בעת טעינת הדוח</font></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</div>
	</body>
</html>
<!--#include file="code.asp" -->
<%  Session.LCID = 1037
OrgID=264
	quest_id = trim(Request("quest_id"))
	
	
if quest_id=16504 then ProductName="טופס מתעניין בטיול"
if quest_id=16735 then ProductName="טופס רישום חתום"
if quest_id=17001 then  ProductName="טופס פניות הציבור"
if quest_id=17057 then ProductName="טופס תלונת לקוח"

	start_date = DateValue(trim(Request("dateStart")))
    end_date = DateValue(trim(Request("dateEnd")))
 ' response.Write quest_id &":" &start_date &":" &end_date 
  'response.end
'///////////////////////////
dir_var = "rtl" :  align_var = "center" : dir_obj_var = "ltr"
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
	strXLS.Add  "   width:150px;"	
	strXLS.Add  "	border: 0.5pt solid black;"
	strXLS.Add  "	mso-pattern:auto none;"
	strXLS.Add  "	mso-height-source: auto;"	
	strXLS.Add  " white-space:normal;}"
	
	strXLS.Add  ".xl25_blue" ' title cell
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	color:#000000;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	font-weight:700;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	text-align:center;"
	strXLS.Add  "	background:#D4E3F2;"
	strXLS.Add  "   width:150px;"	
	strXLS.Add  "	border: 0.5pt solid black;"
	strXLS.Add  "	mso-pattern:auto none;"
	strXLS.Add  "	mso-height-source: auto;"	
	strXLS.Add  " white-space:normal;}"
	
	strXLS.Add  ".xl25_blueRight" ' title cell
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	color:#000000;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	font-weight:700;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	text-align:right;"
	strXLS.Add  "	background:#D4E3F2;"
	strXLS.Add  "   width:150px;"	
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
	strXLS.Add  ".xl26_yellow"  ' data cells
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	font-weight:400;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	background:yellow;"
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
	strXLS.Add  "<x:Name>דוח טפסים לפי טיולים</x:Name>"
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
	strXLS.Add  "<x:DisplayRightToLeft/>"
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
    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=10 class=xl24 style='background:#808080;color:#ffffff'>דוח "& ProductName&" </td></tr>"

    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=10 class=xl24 style='background:#808080;color:#ffffff'> " & start_date & " - " & end_date & "  בין תאריכים  </td></tr>"
select  case quest_id
case 17057
		strXLS.Add  "<tr style='background:#ff9900;mso-height-source:auto;'>" 
		strXLS.Add "<td class=xl25>"
		strXLS.Add trim(Request.Cookies("bizpegasus")("ContactsOne"))  : strXLS.Add "</td> "
		strXLS.Add "<td  class=xl25>טלפון נייד</td>"
        strXLS.Add "<td  class=xl25>Email</td>"
         strXLS.Add "<td  class=xl25>תאריך פתיחת טופס  ""תלונה"" אחרון</td>"
         strXLS.Add "<td  class=xl25>יעד הטיול האחרון שאליו מתייחס טופס ""תלונה""</td>"
         strXLS.Add "<td  class=xl25>קוד הטיול האחרון שאליו מתייחס טופס ""תלונה""</td>"
         strXLS.Add "<td  class=xl25>תאריך יציאת הטיול שאליו מתייחס טופס ""תלונה""</td>"
        strXLS.Add "<td  class=xl25>שם מדריך</td>"
         strXLS.Add "<td  class=xl25>תאריך פתיחת טופס ""מתעניין בטיול""</td>"
         strXLS.Add "<td  class=xl25>מי פתח את טופס ""מתעניין בטיול""</td>"
         strXLS.Add "<td  class=xl25>כמה טפסי תיעוד שיחה קיימים בין פתיחת טופס תלונה עד לפתיחת טופס מתעניין</td>"
        strXLS.Add "<td  class=xl25>תאריך פתיחת טופס ""רישום חתום""</td>"
         strXLS.Add	"</tr>"  : strXLS.Add vbCrlf	

 	sql = "Exec dbo.get_appealsLoyalty17057 @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @productID='" & nFix(quest_id)  &"'"
'	response.Write sql 
'	response.end
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
		contact_name=trim(arr_app(3,aa))
		cellular = trim(arr_app(4,aa))
		email = trim(arr_app(5,aa))		
		GuideName= trim(arr_app(6,aa))
		user=trim(arr_app(7,aa))
		TourName = trim(arr_app(8,aa))
		DepartureName = trim(arr_app(9,aa))
    	Country = trim(arr_app(10,aa))
		KodTiul= trim(arr_app(11,aa))
		
		ContactId= trim(arr_app(12,aa))
	
			If isNumeric(appId) Then
			
			
		
			
		strXLS.Add  "<tr>"
		strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add contact_name
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add cellular : strXLS.Add "</td>"
	    strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add email : strXLS.Add "</td>"
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add appDate : strXLS.Add "</td>"
		
		strXLS.Add "<td class=xl26>" :strXLS.Add Country  : strXLS.Add "</td>"
		strXLS.Add "<td class=xl26>" : strXLS.Add KodTiul :strXLS.Add "</td>"
  		strXLS.Add "<td class=xl26 x:str>" : strXLS.Add DepartureName :strXLS.Add "</td>"
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add GuideName : strXLS.Add "</td>"
			
		'''
		sqlStr1="Exec dbo.get_Data16504 " & ContactId & ",'" & appDate &"'"
				'תאריך פתיחת טופס מתעניין בטיול,
				'מי פתח את טופס מתעניין בטיול,
				'כמה טפסי תיעוד שיחה קיימים בין פתיחת טופס תלונה עד לפתיחת טופס מתעניין,תאריך
				' פתיחת טופס רישום חתום
'response.Write sqlStr1
		set rs_fields=con.GetRecordSet(sqlStr1)
		If not rs_fields.eof Then	
		User16504=rs_fields.Fields.Item(1)	
		date16504=rs_fields.Fields.Item(0)	
		else
		User16504=""
		date16504=""
		End If

		Set rs_fields = Nothing
'		response.Write "User16504="& User16504 &":"& date16504
'		response.end
strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add  date16504  : strXLS.Add "</td>"
strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add  User16504  : strXLS.Add "</td>"

	sqlStr1="Exec dbo.get_Data16470 " & ContactId & ",'" & appDate & "','" & date16504 &"'"
'	response.Write sqlStr1
'	response.end
	set rs_fields=con.GetRecordSet(sqlStr1)
		If not rs_fields.eof Then	
			count=rs_fields.Fields.Item(0)	
		else
		count=0
		End If

		Set rs_fields = Nothing

			strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add  count  : strXLS.Add "</td>"
			
			
			
			
			
	sqlStr1="Exec dbo.get_Data16735 " & ContactId & ",'" & appDate &"'"
				'תאריך פתיחת טופס מתעניין בטיול,
				'מי פתח את טופס מתעניין בטיול,
				'כמה טפסי תיעוד שיחה קיימים בין פתיחת טופס תלונה עד לפתיחת טופס מתעניין,תאריך
				' פתיחת טופס רישום חתום
'response.Write sqlStr1
		set rs_fields=con.GetRecordSet(sqlStr1)
		If not rs_fields.eof Then	

		date16735=rs_fields.Fields.Item(0)	
		else
		date16735=""
		End If

		Set rs_fields = Nothing

		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add date16735 : strXLS.Add "</td>"
strXLS.Add "</tr>"  : strXLS.Add vbCrlf	

			'''
		
		End If
			strXLS.Add "</tr>"
	Next 
	End If
	
case 17001 'טופס פניות הציבור
	strXLS.Add  "<tr style='background:#ff9900;mso-height-source:auto;'>" 
		strXLS.Add "<td class=xl25>"
		strXLS.Add trim(Request.Cookies("bizpegasus")("ContactsOne"))  : strXLS.Add "</td> "
		strXLS.Add "<td  class=xl25>טלפון נייד</td>"
        strXLS.Add "<td  class=xl25>Email</td>"
         strXLS.Add "<td  class=xl25>תאריך פתיחת טופס</td>"
         
         strXLS.Add "<td  class=xl25>יעד הטיול</td>"
         strXLS.Add "<td  class=xl25>קוד הטיול</td>"
         strXLS.Add "<td  class=xl25>תאריך יציאת הטיול</td>"
        strXLS.Add "<td  class=xl25>שם מדריך</td>"
         strXLS.Add "<td  class=xl25>תאריך פתיחת טופס ""מתעניין בטיול""</td>"
         strXLS.Add "<td  class=xl25>מי פתח את טופס ""מתעניין בטיול""</td>"
         strXLS.Add "<td  class=xl25>תאריך פתיחת טופס ""רישום חתום""</td>"
         strXLS.Add	"</tr>"  : strXLS.Add vbCrlf	
	sql = "Exec dbo.get_appealsLoyalty17001 @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @productID='" & nFix(quest_id)  &"'"
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
		contact_name=trim(arr_app(3,aa))
		cellular = trim(arr_app(4,aa))
		email = trim(arr_app(5,aa))		
		GuideName= trim(arr_app(6,aa))
		user=trim(arr_app(7,aa))
		TourName = trim(arr_app(8,aa))
		DepartureName = trim(arr_app(9,aa))
    	Country = trim(arr_app(10,aa))
		KodTiul= trim(arr_app(11,aa))
		ContactId= trim(arr_app(12,aa))
		CountryCode = trim(arr_app(13,aa))
			If isNumeric(appId) Then
		strXLS.Add  "<tr>"
		strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add contact_name
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add cellular : strXLS.Add "</td>"
	    strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add email : strXLS.Add "</td>"
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add appDate : strXLS.Add "</td>"
		
		strXLS.Add "<td class=xl26>" :strXLS.Add Country  : strXLS.Add "</td>"
		strXLS.Add "<td class=xl26>" : strXLS.Add KodTiul :strXLS.Add "</td>"
    	strXLS.Add "<td class=xl26 x:str>" : strXLS.Add DepartureName :strXLS.Add "</td>"
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add GuideName : strXLS.Add "</td>"

			'sqlStr1="Exec dbo.get_Data16504ForCountryCRM " & ContactId & ",'" & CountryCode &"'"
				sqlStr1="Exec dbo.get_Data16504 " & ContactId & ",'" & appDate &"'"
				'תאריך פתיחת טופס מתעניין בטיול,
				'מי פתח את טופס מתעניין בטיול,
				'כמה טפסי תיעוד שיחה קיימים בין פתיחת טופס תלונה עד לפתיחת טופס מתעניין,תאריך
				' פתיחת טופס רישום חתום
'response.Write sqlStr1
'response.end
		set rs_fields=con.GetRecordSet(sqlStr1)
		If not rs_fields.eof Then	
		User16504=rs_fields.Fields.Item(1)	
		date16504=rs_fields.Fields.Item(0)	
		else
		User16504=""
		date16504=""
		End If
	Set rs_fields = Nothing

'		response.Write "User16504="& User16504 &":"& date16504
'		response.end
strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add  date16504  : strXLS.Add "</td>"
strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add  User16504  : strXLS.Add "</td>"
	'if IsNumeric(CountryCode) then
	'	sqlStr1="Exec dbo.get_Data16735ForCountryCRM " & ContactId  & ",'" & CountryCode &"'"
			sqlStr1="Exec dbo.get_Data16735 " & ContactId & ",'" & appDate &"'"
				'תאריך פתיחת טופס מתעניין בטיול,
				'מי פתח את טופס מתעניין בטיול,
				'כמה טפסי תיעוד שיחה קיימים בין פתיחת טופס תלונה עד לפתיחת טופס מתעניין,תאריך
				' פתיחת טופס רישום חתום
'response.Write sqlStr1
'response.end
		set rs_fields=con.GetRecordSet(sqlStr1)
		If not rs_fields.eof Then	

		date16735=rs_fields.Fields.Item(0)	
		else
		date16735=""
		End If

		Set rs_fields = Nothing
'
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add date16735 : strXLS.Add "</td>"
strXLS.Add "</tr>"  : strXLS.Add vbCrlf	
	
		
		End If
'			strXLS.Add "</tr>"
	Next 
	End If
	
'end case 17001
'----------------------------------------------------------------------------------------
case 16504 'טופס מתעניין בטיול
	strXLS.Add  "<tr style='background:#ff9900;mso-height-source:auto;'>" 
		strXLS.Add "<td class=xl25>"
		strXLS.Add trim(Request.Cookies("bizpegasus")("ContactsOne"))  : strXLS.Add "</td> "
		strXLS.Add "<td  class=xl25>טלפון נייד</td>"
        strXLS.Add "<td  class=xl25>Email</td>"
         strXLS.Add "<td  class=xl25>תאריך פתיחת טופס ""מתעניין בטיול""</td>"
         
         strXLS.Add "<td  class=xl25>יעד הטיול</td>"
         strXLS.Add "<td  class=xl25>שם המשתמש  שפתח את טופס ""מתעניין בטיול""</td>"
         strXLS.Add "<td  class=xl25>תאריך פתיחת טופס ""רישום חתום""</td>"
        strXLS.Add "<td  class=xl25>שם המשתמש  שפתח את טופס ""רישום חתום"" </td>"
               strXLS.Add "<td  class=xl25>כמה טפסי ""תיעוד שיחה"" קיימים בין פתיחת טופס ""מתעניין בטיול"" עד לפתיחת טופס ""רישום חתום""</td>"
             strXLS.Add	"</tr>"  : strXLS.Add vbCrlf	

	sql = "Exec dbo.get_appealsLoyalty16504 @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @productID='" & nFix(quest_id)  &"'"
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
		
		
		contact_name=trim(arr_app(3,aa))
		cellular = trim(arr_app(4,aa))
		email = trim(arr_app(5,aa))		
		user=trim(arr_app(6,aa))
		'TourName = trim(arr_app(8,aa))
	
    	Country = trim(arr_app(7,aa))
    	Contactid= trim(arr_app(8,aa))
    	Country_id= trim(arr_app(9,aa))
		'KodTiul= trim(arr_app(11,aa))
	
			If isNumeric(appId) Then
		strXLS.Add  "<tr>"
		strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add contact_name
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add cellular : strXLS.Add "</td>"
	    strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add email : strXLS.Add "</td>"
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add appDate : strXLS.Add "</td>"
		
		strXLS.Add "<td class=xl26>" :strXLS.Add Country  : strXLS.Add "</td>"
		strXLS.Add "<td class=xl26>" : strXLS.Add user :strXLS.Add "</td>"


		sqlStr1="Exec dbo.get_Data16735ByRelationId " & appId
			'	response.Write sqlStr1
	'		response.end
				'תאריך פתיחת טופס מתעניין בטיול,
				'מי פתח את טופס מתעניין בטיול,
				'כמה טפסי תיעוד שיחה קיימים בין פתיחת טופס תלונה עד לפתיחת טופס מתעניין,תאריך
				' פתיחת טופס רישום חתום

		set rs_fields=con.GetRecordSet(sqlStr1)
		If not rs_fields.eof Then	
		User16735=rs_fields.Fields.Item(1)	
		date16735=rs_fields.Fields.Item(0)	
		else
		User16735=""
		date16735=""
		End If
	Set rs_fields = Nothing

    	strXLS.Add "<td class=xl26 x:str>" : strXLS.Add date16735 :strXLS.Add "</td>"
		strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add User16735 : strXLS.Add "</td>"
		
			if isNumeric(date16735) then
			sqlStr1="Exec dbo.get_Data16470 " & ContactId & ",'" & appDate & "','" & date16735 &"'"
	'response.Write sqlStr1
'	response.end
	set rs_fields=con.GetRecordSet(sqlStr1)
		If not rs_fields.eof Then	
			count=rs_fields.Fields.Item(0)	
		else
		count=0
		End If

		Set rs_fields = Nothing
		else
		count=""
		end  if

			strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add  count  : strXLS.Add "</td>"
		
	
		End If
			strXLS.Add "</tr>"
	Next 
	End If
	

'end case 16504
'----------------------------------------------------------------------------------------
case 16735  'טופס רישום חתום
i=0
		strXLS.Add  "<tr style='background:#ff9900;mso-height-source:auto;'>" 
		strXLS.Add "<td class=xl25>"
		strXLS.Add trim(Request.Cookies("bizpegasus")("ContactsOne"))  : strXLS.Add "</td> "
		strXLS.Add "<td  class=xl25>טלפון נייד</td>"
        strXLS.Add "<td  class=xl25>Email</td>"
         strXLS.Add "<td  class=xl25>תאריך פתיחת טופס ""רישום חתום"" ראשון</td>"
       	 strXLS.Add "<td  class=xl25>יתאריך פתיחת טופס ""רישום חתום"" שני</td>"
         strXLS.Add "<td  class=xl25>שםתאריך פתיחת טופס ""רישום חתום"" שלישי</td>"
         strXLS.Add "<td  class=xl25>תאריך פתיחת טופס ""רישום חתום"" רביעי</td>"
		 strXLS.Add "<td  class=xl25>שםתאריך פתיחת טופס ""רישום חתום"" חמישי</td>"
         strXLS.Add "<td  class=xl25>שםתאריך פתיחת טופס ""רישום חתום"" שישי</td>"
		 strXLS.Add "<td  class=xl25>שםתאריך פתיחת טופס ""רישום חתום"" שביעי</td>"
		 strXLS.Add "<td  class=xl25>שםתאריך פתיחת טופס ""רישום חתום"" שמיני</td>"
		 strXLS.Add "<td  class=xl25>שםתאריך פתיחת טופס ""רישום חתום"" תשיעי</td>"
         strXLS.Add	"</tr>"  : strXLS.Add vbCrlf	
         	sql = "Exec dbo.get_appealsLoyalty16735 @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @productID='" & nFix(quest_id)  &"'"
'	response.Write sql
'	response.end
   set rsa = con.getRecordSet(sql)
	If Not rsa.Eof Then
		arr_app = rsa.getRows()
	End If
	Set rsa = Nothing	
	
	If isArray(arr_app) Then
	pContactId=""
	For aa=0 To Ubound(arr_app,2)
		appId = trim(arr_app(0,aa))
		appDate=trim(arr_app(1,aa))
		If IsDate(appDate) Then
			appDate = Day(appDate) & "/" & Month(appDate) & "/" & Year(appDate)
		Else
			appDate = ""
		End If
    	contact_name=trim(arr_app(4,aa))
		cellular = trim(arr_app(5,aa))
		email = trim(arr_app(6,aa))		
		ContactId= trim(arr_app(3,aa))
	if pContactId<>ContactId then
	if pContactId<>"" then
	for j=i to 8
		strXLS.Add "<td class=xl26 x:str></td>"
		next
			strXLS.Add  "</tr>"
			i=0
	end if
	
		strXLS.Add  "<tr>"
		strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add contact_name
		strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add cellular : strXLS.Add "</td>"
	    strXLS.Add "</td><td class=xl26 x:str>"  : strXLS.Add email : strXLS.Add "</td>"
	     pContactId=ContactId 
	end if
	if pContactId=ContactId then
		i=i+1
		strXLS.Add "<td class=xl26 x:str>"  : strXLS.Add appDate : strXLS.Add "</td>"
end if
	    
	 
Next 
	End If
'end case 16735
'----------------------------------------------------------------------------------------

end select
 
    strXLS.Add  "</TABLE></body></html>"    
	Set con = Nothing	 
	
	filestring = "ReportL_" &Minute(now())&"_"&Second(now()) &".xls"
    Response.Clear()
    Response.AddHeader "content-disposition", "inline; filename=" & filestring
    Response.ContentType = "application/vnd.ms-excel"
    Response.Write(strXLS.value)  
    Set strXLS = Nothing %>

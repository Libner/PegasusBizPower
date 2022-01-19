<%Server.ScriptTimeout=20000
  Response.Buffer=True
%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%
  
   OrgID = trim(Request.Cookies("bizpegasus")("OrgID")) 
    
   set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   filestring="../../../download/reports/"
   'Response.Write server.mappath(filestring)
   'Response.End
   fs.DeleteFile server.mappath(filestring) & "/*.*" ,False   
  
   filestring="../../../download/reports/export_" & trim(OrgID) & ".xls"
   'Response.Write server.mappath(filestring)
   'Response.End
   set XLSfile=fs.CreateTextFile(server.mappath(filestring)) ' Create text file as excel file
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
	XLSfile.WriteLine "	mso-height-source:auto userset;"
	XLSfile.WriteLine "	height:20pt;}"
	XLSfile.WriteLine ".xl24"  ' title
	XLSfile.WriteLine "	{mso-style-parent:style0;"
	XLSfile.WriteLine "	font-family:Arial, sans-serif;"
	XLSfile.WriteLine "	font-size:12.0pt;"
	XLSfile.WriteLine "	font-weight:700;"
	XLSfile.WriteLine "	text-align:center;"
	XLSfile.WriteLine "	mso-height-source: userset;"
	XLSfile.WriteLine "	height:18pt;"
	XLSfile.WriteLine "	white-space:normal;}"
	XLSfile.WriteLine ".xl25" ' title cell
	XLSfile.WriteLine "	{mso-style-parent:style0;"
	XLSfile.WriteLine "	color:#000000;"
	XLSfile.WriteLine "	font-size:10.0pt;"
	XLSfile.WriteLine "	font-weight:700;"
	XLSfile.WriteLine "	font-family:Arial, sans-serif;"
	XLSfile.WriteLine "	text-align:center;"
	XLSfile.WriteLine "	background:#FFD011;"
	XLSfile.WriteLine "	border: 0.5pt solid black;"
	XLSfile.WriteLine "	mso-pattern:auto none;"
	XLSfile.WriteLine "	mso-height-source: userset;"
	XLSfile.WriteLine "	height:16pt;"
	XLSfile.WriteLine " white-space:normal;}"
	XLSfile.WriteLine ".xl25_1" ' title cell
	XLSfile.WriteLine "	{mso-style-parent:style0;"
	XLSfile.WriteLine "	color:#000000;"
	XLSfile.WriteLine "	font-size:10.0pt;"
	XLSfile.WriteLine "	font-weight:700;"
	XLSfile.WriteLine "	font-family:Arial, sans-serif;"
	XLSfile.WriteLine "	text-align:center;"
	XLSfile.WriteLine "	background:#C0C0C0;"
	XLSfile.WriteLine "	border: 0.5pt solid black;"
	XLSfile.WriteLine "	mso-pattern:auto none;"
	XLSfile.WriteLine "	mso-height-source: userset;"
	XLSfile.WriteLine "	height:16pt;"
	XLSfile.WriteLine " white-space:normal;}"
	XLSfile.WriteLine ".xl26"  ' data cells
	XLSfile.WriteLine "	{mso-style-parent:style0;"
	XLSfile.WriteLine "	font-weight:400;"
	XLSfile.WriteLine "	font-size:10.0pt;"
	XLSfile.WriteLine "	font-family:Arial, sans-serif;"
	XLSfile.WriteLine "	border: 0.5pt solid black;"
	XLSfile.WriteLine "	mso-height-source: userset;"
	XLSfile.WriteLine "	height:16pt;"
	XLSfile.WriteLine "	direction:rtl;"
	XLSfile.WriteLine "	text-align:right;}"
	XLSfile.WriteLine "-->"
	XLSfile.WriteLine "</style>"
	XLSfile.WriteLine "<!--[if gte mso 9]><xml>"
	XLSfile.WriteLine "<x:ExcelWorkbook>"
	XLSfile.WriteLine "<x:ExcelWorksheets>"
	XLSfile.WriteLine "<x:ExcelWorksheet>"
	XLSfile.WriteLine "<x:Name>Export</x:Name>"
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
	XLSfile.WriteLine "</x:WorksheetOptions>"
	XLSfile.WriteLine "</x:ExcelWorksheet>"
	XLSfile.WriteLine "</x:ExcelWorksheets>"
	XLSfile.WriteLine "<x:WindowTopX>360</x:WindowTopX>"
	XLSfile.WriteLine "<x:WindowTopY>60</x:WindowTopY>"
	XLSfile.WriteLine "<x:AcceptLabelsInFormulas/>"
    XLSfile.WriteLine "<x:ProtectStructure>False</x:ProtectStructure>"
    XLSfile.WriteLine "<x:ProtectWindows>False</x:ProtectWindows>"
    XLSfile.WriteLine "</x:ExcelWorkbook>"
    XLSfile.WriteLine "<x:ExcelName>"
    XLSfile.WriteLine "<x:Name>Contacts</x:Name>"
    XLSfile.WriteLine "<x:Formula>=Contacts!$A$1:$M$20000</x:Formula>"
    XLSfile.WriteLine "</x:ExcelName>"
	XLSfile.WriteLine "</xml><![endif]-->"
	XLSfile.WriteLine "</head>"
	XLSfile.WriteLine "<body>"
	
	XLSfile.WriteLine "<table style='border-collapse:collapse;table-layout:fixed'>"
  
    strLine="<tr><td class=xl25_1 x:str=""'שם"">איש קשר</td><td class=xl25_1 x:str=""'תפקיד"">תפקיד</td><td class=xl25_1 x:str=""'טלפון"">טלפון אישי</td><td class=xl25_1 x:str=""'פקס"">פקס אישי</td><td class=xl25_1 x:str=""'נייד"">נייד</td><td class=xl25_1 x:str=""'אימייל"">אימייל אישי</td><td class=xl25 x:str=""'חברה"">חברה</td><td class=xl25 x:str=""'טלפון בעבודה"">טלפון</td><td class=xl25 x:str=""'פקס בעבודה"">פקס</td><td class=xl25 x:str=""'אימייל בעבודה"">אימייל</td><td class=xl25 x:str=""'אתר"">אתר</td><td class=xl25 x:str=""'כתובת בעבודה"">כתובת</td><td class=xl25 x:str=""'עיר"">עיר</td><td class=xl25 x:str=""'מיקוד"">מיקוד</td></tr>"
    XLSfile.writeline strLine		

    SQL = "SELECT CONTACTS.CONTACT_NAME, CONTACTS.phone, CONTACTS.fax, CONTACTS.cellular, "&_
    " CONTACTS.email, messangers.item_name, COMPANIES.City_Name, COMPANIES.COMPANY_NAME, COMPANIES.address, "&_
    " COMPANIES.phone1, COMPANIES.fax1, COMPANIES.url, COMPANIES.Email, COMPANIES.zip_code FROM CONTACTS "&_
    " INNER JOIN COMPANIES ON CONTACTS.COMPANY_ID = COMPANIES.COMPANY_ID "&_
    " LEFT OUTER JOIN messangers ON CONTACTS.messanger_id = messangers.item_ID "&_   
    " WHERE companies.ORGANIZATION_ID = " & OrgID & " Order BY CONTACTS.CONTACT_NAME"
    'Response.Write sql
    'Response.End
    set listContact=con.GetRecordSet(SQL)
    If not listContact.EOF Then   
	  contArray = listContact.GetRows()	  	  
	  recCount =  listContact.RecordCount 		
	  listContact.close
      set listContact=Nothing
      count = 0
      do while count < recCount
		name = trim(contArray(0,count))
		If Len(trim(contArray(1,count))) > 3 Then
			phone = trim(contArray(1,count))
		Else
			phone = ""	
		End If
		If Len(trim(contArray(2,count))) > 3 Then
			fax = trim(contArray(2,count))
		Else
			fax = ""	
		End If
		If Len(trim(contArray(3,count))) > 3 Then
			cellurar = trim(contArray(3,count)) 
		Else
			cellurar = ""	
		End If     
		email = contArray(4,count)            
		duty = trim(contArray(5,count))
		City_Name = trim(contArray(6,count))
		COMPANY_NAME = trim(contArray(7,count))
		address = trim(contArray(8,count))
		If Len(trim(contArray(9,count))) > 3 Then
			company_phone = trim(contArray(9,count))
		Else
			company_phone = ""	
		End If
		If Len(trim(contArray(10,count))) > 3 Then
			company_fax = trim(contArray(10,count))
		Else
			company_fax = ""	
		End If	
		url = trim(contArray(11,count))
		company_email = trim(contArray(12,count))
		company_zip_code = trim(contArray(13,count))

		strLine = "<tr><td class=xl26>" & name & "</td><td class=xl26>" & duty & "</td><td class=xl26>" & phone & "</td>" & _
		"<td class=xl26>" & fax & "</td><td class=xl26>" & cellurar & "</td><td class=xl26 style=""direction:ltr"">" & email & "</td>"  &_		
		"<td class=xl26>" & COMPANY_NAME & "</td><td class=xl26>" & company_phone & "</td><td class=xl26>" & company_fax & "</td>" & _
		"<td class=xl26 style=""direction:ltr"">" & company_email & "</td><td class=xl26 style=""direction:ltr"">" & url & "</td>" & _
		"<td class=xl26>" & address & "</td><td class=xl26>" & City_Name & "</td><td class=xl26>" & company_zip_code & "</td></tr>"
	    
		XLSfile.writeline strLine      
	      
		count=count+1
      loop
    End If 
	
	XLSfile.WriteLine "</TABLE></body></html>"
	XLSfile.close
	set con=Nothing
	 
	Response.Redirect(filestring) 'Open the excel file in the browser
 %>

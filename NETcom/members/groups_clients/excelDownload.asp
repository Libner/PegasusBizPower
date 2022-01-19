<%Server.ScriptTimeout=20000
  Response.Buffer=True
%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%
   groupId = trim(Request.QueryString("groupId"))
   OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))   
   
   If trim(groupId) <> "" Then
		sqlstr="Select GROUPNAME FROM GROUPS WHERE GROUPID="&groupId&" And Organization_ID = " & OrgID
		set rsn = con.getRecordSet(sqlstr)
		If not rsn.eof Then
			groupname = trim(rsn(0))
			found_group = true
		Else
			found_group = false	
		End if
		set rsn = Nothing	
  Else
	  found_group = false	
  End If  
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body style="margin:0px;background-color:#E5E5E5" onload="window.focus();">
<table cellpadding=0 cellspacing=0 width=100%>
<tr><td width="100%" class="page_title" dir=rtl>&nbsp;הורדת נמענים לקובץ&nbsp;Excel&nbsp;<font color="#6E6DA6"><%=groupName%></font></td></tr>         
<tr><td width="100%" colspan=2 height=10></td></tr>
<tr><td width="100%" colspan=2>
<% 
   set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   filestring="../../../download/reports/"
   'Response.Write server.mappath(filestring)
   'Response.End
   fs.DeleteFile server.mappath(filestring) & "/*.*" ,False   
  
   filestring="../../../download/reports/export_group_" & trim(OrgID) & "_" & groupId & ".xls"
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
XLSfile.WriteLine "	text-align:left;"
XLSfile.WriteLine "	padding-left: 5px;"
XLSfile.WriteLine "	mso-style-name:Normal;"
XLSfile.WriteLine "	mso-height-source:auto userset;"
XLSfile.WriteLine "	height:20pt;}"
XLSfile.WriteLine ".xl24"  ' title
XLSfile.WriteLine "	{mso-style-parent:style0;"
XLSfile.WriteLine "	font-family:Arial, sans-serif;"
XLSfile.WriteLine "	font-size:13.0pt;"
XLSfile.WriteLine "	font-weight:700;"
XLSfile.WriteLine "	text-align:center;"
XLSfile.WriteLine "	white-space:normal;}"
XLSfile.WriteLine ".xl25" ' title cell
XLSfile.WriteLine "	{mso-style-parent:style0;"
XLSfile.WriteLine "	color:#000000;"
XLSfile.WriteLine "	font-size:11.0pt;"
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
XLSfile.WriteLine "	font-family:Arial, sans-serif;"
XLSfile.WriteLine "	mso-font-charset:0;"
XLSfile.WriteLine "	text-align:right;"
XLSfile.WriteLine "	border:.5pt solid black;"
XLSfile.WriteLine "	white-space:normal;"
XLSfile.WriteLine "	}"
XLSfile.WriteLine ".xl27"  ' data cells
XLSfile.WriteLine "	{mso-style-parent:style0;"
XLSfile.WriteLine "	font-family:Arial, sans-serif;"
XLSfile.WriteLine "	mso-font-charset:0;"
XLSfile.WriteLine "	mso-number-format:Fixed;"
XLSfile.WriteLine "	text-align:left;"
XLSfile.WriteLine "	font-weight:600;"
XLSfile.WriteLine "	white-space:normal;"
XLSfile.WriteLine "	}"
XLSfile.WriteLine ".xl30"
XLSfile.WriteLine "	{mso-style-parent:style0;"
XLSfile.WriteLine "	font-family:Arial, sans-serif;"
XLSfile.WriteLine "	mso-font-charset:0;"
XLSfile.WriteLine "	mso-number-format:Fixed;"
XLSfile.WriteLine "	text-align:left;"
XLSfile.WriteLine "	border-top:none;"
XLSfile.WriteLine "	border-left:.5pt solid black;"
XLSfile.WriteLine "	border-bottom:.5pt solid black;"
XLSfile.WriteLine "	border-right:.5pt solid black;"
XLSfile.WriteLine "	white-space:normal;}"
XLSfile.WriteLine ".xl31"
XLSfile.WriteLine "	{mso-style-parent:style0;"
XLSfile.WriteLine "	font-family:Arial, sans-serif;"
XLSfile.WriteLine "	mso-font-charset:0;"
XLSfile.WriteLine "	mso-number-format:""Short Date"";"
XLSfile.WriteLine "	text-align:center;"
XLSfile.WriteLine "	border-top:none;"
XLSfile.WriteLine "	border-left:.5pt solid black;"
XLSfile.WriteLine "	border-bottom:.5pt solid black;"
XLSfile.WriteLine "	border-right:.5pt solid black;"
XLSfile.WriteLine "	white-space:normal;}"
XLSfile.WriteLine "-->"
XLSfile.WriteLine "</style>"
XLSfile.WriteLine "<!--[if gte mso 9]><xml>"
XLSfile.WriteLine "<x:ExcelWorkbook>"
XLSfile.WriteLine "<x:ExcelWorksheets>"
XLSfile.WriteLine "<x:ExcelWorksheet>"
XLSfile.WriteLine "<x:Name>"&groupName&"</x:Name>"
XLSfile.WriteLine "<x:WorksheetOptions>"
XLSfile.WriteLine "<x:Print>"
XLSfile.WriteLine "      <x:ValidPrinterInfo/>"
XLSfile.WriteLine "      <x:PaperSizeIndex>9</x:PaperSizeIndex>"
XLSfile.WriteLine "      <x:HorizontalResolution>600</x:HorizontalResolution>"
XLSfile.WriteLine "      <x:VerticalResolution>600</x:VerticalResolution>"
XLSfile.WriteLine "</x:Print>"
XLSfile.WriteLine "<x:DisplayRightToLeft/>"
XLSfile.WriteLine " <x:Panes>"
XLSfile.WriteLine "   <x:Pane>"
XLSfile.WriteLine "       <x:Number>3</x:Number>"
XLSfile.WriteLine "       <x:ActiveRow>9</x:ActiveRow>"
XLSfile.WriteLine "       <x:ActiveCol>5</x:ActiveCol>"
XLSfile.WriteLine "   </x:Pane>"
XLSfile.WriteLine " </x:Panes>"
XLSfile.WriteLine "<x:Selected/>"
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
	
   XLSfile.WriteLine "<table style='border-collapse:collapse;table-layout:fixed'>"
   strLine="<tr><td colspan=6 class=xl24> נמענים של קבוצה "& groupName
   XLSfile.writeline strLine		
   strLine="<tr><td></td><td></td><td></td><td></td></tr>"
   XLSfile.writeline strLine
   strLine="<tr><td class=xl25>Email</td><td class=xl25>שם מלא</td><td class=xl25>חברה</td>"&_
   "<td class=xl25>טלפון חברה</td><td class=xl25>כתובת חברה</td><td class=xl25>עיר חברה</td>"&_
   "<td class=xl25>מיקוד חברה</td><td class=xl25>תפקיד</td><td class=xl25>טלפון איש קשר</td>"&_
   "<td class=xl25>כתובת איש קשר</td><td class=xl25>עיר</td><td class=xl25>מיקוד</td></tr>"
   XLSfile.writeline strLine		

    sqlStr = "Select PEOPLE_NAME,PEOPLE_PHONE,PEOPLE_EMAIL,PEOPLE_COMPANY,PEOPLE_OFFICE, " &_
	" Contact_Address, Contact_City_Name, contact_zip_code, prefix_phone1 + '-' + phone1, Address,"&_
	" City_Name, zip_code "&_
	" FROM PEOPLES Left Outer Join Contacts On " &_
	" PEOPLES.Contact_ID = Contacts.Contact_ID Left Outer Join Companies "&_
	" On PEOPLES.COMPANY_ID = Companies.COMPANY_ID " &_
	" Where PEOPLES.ORGANIZATION_ID=" & OrgID &_
	" And PEOPLES.groupId="& groupId &" Order by PEOPLE_NAME,PEOPLE_COMPANY" 
	'Response.Write sqlStr
	'Response.End 
	set rs_peoples = con.GetRecordSet(sqlStr)
	If not rs_peoples.eof Then
		arr_p = rs_peoples.getRows()
	End If
	Set rs_peoples = Nothing
	
	If IsArray(arr_p) Then
	For pp=0 To Ubound(arr_p,2)
		peopleName=trim(arr_p(0,pp))				
		peoplePHONE=trim(arr_p(1,pp))	
		peopleEMAIL=trim(arr_p(2,pp))
		peopleCOMPANY=trim(arr_p(3,pp))
		peopleOFFICE=trim(arr_p(4,pp))		
		peopleAddress=trim(arr_p(5,pp))
		peopleCity=trim(arr_p(6,pp))
		peopleZip=trim(arr_p(7,pp))
		Phone=trim(arr_p(8,pp))
		If trim(Phone) = "-" Then
			Phone = ""
		End If
		Address=trim(arr_p(9,pp))
		City=trim(arr_p(10,pp))
		ZipCode=trim(arr_p(11,pp))
		strLine = "<tr><td class=xl26>" & peopleEMAIL & "</td><td class=xl26>" & peopleName & "</td>"&_
		"<td class=xl26>" & peopleCOMPANY & "</td><td class=xl26>" & Phone & "</td>" &_
		"<td class=xl26>" & Address & "</td><td class=xl26>" & City & "</td>"&_
		"<td class=xl26>" & ZipCode & "</td><td class=xl26>" & peopleOFFICE & "</td>"&_
		"<td class=xl26>" & peoplePHONE & "</td><td class=xl26>" & peopleAddress & "</td>"&_
		"<td class=xl26>" & peopleCity & "</td><td class=xl26>" & peopleZip & "</td></tr>"
	    
		XLSfile.writeline strLine		
    Next
    End If
	
	XLSfile.WriteLine "</TABLE></body></html>"
	XLSfile.close

%>

<table border="0" align="right" bgcolor="#e6e6e6" cellpadding="1" cellspacing="0" width="100%">
     <tr><td width="100%" height="100%">&nbsp;</td></tr>
    
      <tr><td align="center" width="100%"><A href="<%=filestring%>" target="_blank" class="file_link" style="font-size:12pt">לחץ להורדת הקובץ</a></td></tr>		
      <tr><td width="100%" height="100%">&nbsp;</td></tr>

</table>

<tr><td height=10 nowrap></td></tr>	
</table></td></tr></table>
</body>		
<%set con=Nothing%>
</html>		
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<%
 
	OrgID	 = trim(Request.Cookies("bizpegasus")("OrgID"))
  	lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))
		
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
 
	if lang_id = "1" then
		arr_Status = Array("","עתידית","הסתיימה","הוכנס סיכום","נדחתה")	
	else
		arr_Status = Array("","Future","Done","Summary added","Postponed")	
	end if	

 	
	If trim(Request("start_date")) <> "" Then	
		start_date = trim(Request("start_date"))
		start_date = Day(start_date) & "/" & Month(start_date) & "/" & Year(start_date) 
	Else
		start_date = FormatDateTime(DateAdd("d", -14, Now()), 2)
	End If	

	If trim(Request("end_date")) <> "" Then
		end_date = trim(Request("end_date"))
		end_date = Day(end_date) & "/" & Month(end_date) & "/" & Year(end_date) 
	Else
		end_date = FormatDateTime(DateAdd("d", 14, Now()), 2)
	End If		
	
	if trim(Request.Form("search_company")) <> "" Or trim(Request.QueryString("search_company")) <> "" then
		search_company = trim(Request.Form("search_company"))
		if trim(Request.QueryString("search_company")) <> "" then
			search_company = trim(Request.QueryString("search_company"))
		end if					
		where_company = " And company_Name LIKE '%"& sFix(search_company) &"%'"
		task_status = "all"			
	else
		where_company = ""		
	end if


	if trim(Request.Form("search_contact")) <> "" Or trim(Request.QueryString("search_contact")) <> "" then
		search_contact = trim(Request.Form("search_contact"))
		if trim(Request.QueryString("search_contact")) <> "" then
			search_contact = trim(Request.QueryString("search_contact"))
		end if					
		where_contact = " And CONTACT_NAME LIKE '%"& sFix(search_contact) &"%'"
		task_status = "all"					
	else
		where_contact = ""		
	end if

	if trim(Request("search_project")) <> "" then		
		search_project = trim(Request("search_project"))	
		where_project = " And project_Name LIKE '%"& sFix(search_project) &"%'"			
		status = "all"	
	else
		where_project = ""	
		search_project = ""		
	end if	
	
	If trim(Request("participant_id")) <> "" Then
		participant_id = trim(Request("participant_id"))
	Else
		participant_id = ""'UserID	
	End If
	
	search_status = trim(Request("search_status"))		
	
	If trim(participant_id) <> "" Then
		sqlstr = "Select FirstName, LastName From Users WHERE User_ID = " & participant_id
		set rs_user = con.getRecordSet(sqlstr)
		if not rs_user.eof then
			userName = "<br>&nbsp;עובד:&nbsp;<font color=""#666699"">" & trim(rs_user(0)) & " " & trim(rs_user(1))& "</font>"
		end if
		set rs_user = nothing
	End If

	If trim(search_company) <> "" Then
		companyName = "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & ":&nbsp;<font color=""#666699"">" & search_company & "</font>"
	End If
	
	If trim(search_contact) <> "" Then
		contactName = "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("ContactsOne")) & ":&nbsp;<font color=""#666699"">" & search_contact & "</font>"
	End If
	
	If trim(search_project) <> "" Then
		projectName = "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("ProjectOne")) & ":&nbsp;<font color=""#666699"">" & search_project & "</font>"
	End If		

	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 73 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(trim(arr_title(0))) = arr_title(1)
			End If
		Next
	End If
	set rstitle = Nothing
	
set fsotmp = createobject("scripting.filesystemobject")
fsotmp.DeleteFile server.MapPath("../../../download/reports/") & "\*.*",true 
set fsotmp = nothing
		
	set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
	filestring="report_meetings_"&Day(Date())&"_"&Month(Date())&"_"&Year(Date())&".xls"
	set XLSfile=fs.CreateTextFile(server.mappath("../../../download/reports/"&filestring))

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
	XLSfile.WriteLine "td"
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
	XLSfile.WriteLine "	text-align:" & align_var & ";"
	XLSfile.WriteLine "	padding-right: 5px;"
	XLSfile.WriteLine "	mso-style-name:Normal;"
	XLSfile.WriteLine "	mso-height-source:userset;"
	XLSfile.WriteLine "	mso-style-id:0;"
	XLSfile.WriteLine "	min-height:20pt;}"
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
	XLSfile.WriteLine "	height:20pt;"
	XLSfile.WriteLine " white-space:normal;}"
	XLSfile.WriteLine ".xl26"  ' data cells
	XLSfile.WriteLine "	{mso-style-parent:style0;"
	XLSfile.WriteLine "	font-family:Arial, sans-serif;"
	XLSfile.WriteLine "	mso-font-charset:0;"
	XLSfile.WriteLine "	text-align:" & align_var & ";"
	XLSfile.WriteLine "	border:.5pt solid black;"
	XLSfile.WriteLine "	white-space:normal;"
    XLSfile.WriteLine "	min-height:20pt;}"	
	XLSfile.WriteLine ".xl27"  ' data cells
	XLSfile.WriteLine "	{mso-style-parent:style0;"
	XLSfile.WriteLine "	font-family:Arial, sans-serif;"
	XLSfile.WriteLine "	mso-font-charset:0;"
	XLSfile.WriteLine "	text-align:" & align_var & ";"
	XLSfile.WriteLine "	font-weight:600;"
	XLSfile.WriteLine "	white-space:normal;"
	XLSfile.WriteLine "	}"
	XLSfile.WriteLine ".xl30"
	XLSfile.WriteLine "	{mso-style-parent:style0;"
	XLSfile.WriteLine "	font-family:Arial, sans-serif;"
	XLSfile.WriteLine "	mso-font-charset:0;"
	XLSfile.WriteLine "	text-align:" & align_var & ";"
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
	XLSfile.WriteLine "<x:Name>" & arrTitles(1) & "</x:Name>"
	XLSfile.WriteLine "<x:WorksheetOptions>"
	XLSfile.WriteLine "<x:Print>"
	XLSfile.WriteLine "      <x:ValidPrinterInfo/>"
	XLSfile.WriteLine "      <x:PaperSizeIndex>9</x:PaperSizeIndex>"
	XLSfile.WriteLine "      <x:HorizontalResolution>600</x:HorizontalResolution>"
	XLSfile.WriteLine "      <x:VerticalResolution>600</x:VerticalResolution>"
	XLSfile.WriteLine "</x:Print>"
	If trim(lang_id) = "1" Then
	XLSfile.WriteLine "<x:DisplayRightToLeft/>"
	End If
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
	XLSfile.WriteLine "<body>"
	XLSfile.WriteLine "<table style='border-collapse:collapse;table-layout:fixed' dir="&dir_var&">"
	XLSfile.WriteLine "<col width=80>"
	XLSfile.WriteLine "<col width=80>"
	XLSfile.WriteLine "<col width=100>"
	XLSfile.WriteLine "<col width=100>"
	XLSfile.WriteLine "<col width=150>"
	XLSfile.WriteLine "<col width=120>"
	XLSfile.WriteLine "<col width=120>"
	XLSfile.WriteLine "<col width=150>"
	XLSfile.WriteLine "<col width=150>"
	XLSfile.WriteLine "<col width=280>"
	If trim(lang_id) = "1" Then
	strLine="<tr><td colspan=8 class=xl24> דו''ח פגישות לתקופה <br> "& end_date & " - " & start_date
	Else
	strLine="<tr><td colspan=8 class=xl24> Meetings report in period <br> "& start_date & " - " & end_date
	End If
	strLine = strLine & userName & companyName & contactName & projectName & "</td></tr>"
	XLSfile.writeline strLine

	strLine="<tr><td></td><td></td><td></td><td></td></tr>"
	XLSfile.writeline strLine

	strLine = "<tr><td class=xl25 x:str>&nbsp;" & arrTitles(13)&"&nbsp;</td>"&_
	"<td class=xl25 x:str>&nbsp;" & arrTitles(2) &"&nbsp;</td>"&_
	"<td class=xl25 x:str>&nbsp;" & arrTitles(9) &"&nbsp;</td>"&_
	"<td class=xl25 x:str>&nbsp;" & arrTitles(10) &"&nbsp;</td>"&_	
	"<td class=xl25 x:str>&nbsp;" & trim(Request.Cookies("bizpegasus")("CompaniesOne")) &"&nbsp;</td>"&_
	"<td class=xl25 x:str>&nbsp;" & trim(Request.Cookies("bizpegasus")("ContactsOne")) &"&nbsp;</td>"&_
	"<td class=xl25 x:str>&nbsp;" & arrTitles(27) &"&nbsp;</td>"&_
	"<td class=xl25 x:str>&nbsp;" & trim(Request.Cookies("bizpegasus")("ProjectOne")) &"&nbsp;</td>"&_
	"<td class=xl25 x:str>&nbsp;" & arrTitles(24) &"&nbsp;</td>"&_
	"<td class=xl25 x:str>&nbsp;" & arrTitles(11) &"&nbsp;</td>"&_
	"<td class=xl25 x:str>&nbsp;" & arrTitles(28) &"&nbsp;</td>"&_
	"<td class=xl25 x:str>&nbsp;" & arrTitles(26) &"&nbsp;</td></tr>"
	XLSfile.writeline strLine
     
	sqlstr = " SELECT dbo.meetings.meeting_id, dbo.meetings.meeting_date, dbo.meetings.start_time, dbo.meetings.end_time, " &_
    " dbo.COMPANIES.COMPANY_NAME, dbo.CONTACTS.CONTACT_NAME, dbo.CONTACTS.phone, dbo.CONTACTS.cellular,  " &_
    " dbo.meetings.meeting_status, dbo.meetings.meeting_content, dbo.meetings.meeting_closing, "&_
    " dbo.projects.project_name, dbo.USERS.FIRSTNAME + ' ' + dbo.USERS.LASTNAME as UserName " &_
    " FROM dbo.meetings " &_
    " INNER JOIN dbo.USERS ON dbo.meetings.user_id = dbo.USERS.USER_ID " &_
    " LEFT OUTER JOIN dbo.CONTACTS ON dbo.meetings.contact_id = dbo.CONTACTS.CONTACT_ID " &_
    " LEFT OUTER JOIN dbo.COMPANIES ON dbo.meetings.company_id = dbo.COMPANIES.COMPANY_ID " &_
    " LEFT OUTER JOIN dbo.projects ON dbo.meetings.project_id = dbo.projects.project_id " &_
    " Where dbo.meetings.organization_id = "&OrgID
    If trim(participant_id) <> "" Then
		sqlstr = sqlstr & " And dbo.meetings.meeting_id IN (Select DISTINCT meeting_id From dbo.meeting_to_users Where dbo.meeting_to_users.user_id = " & participant_id & ")"
    End If
    If trim(start_date) <> "" Then
		sqlstr = sqlstr & " AND DateDiff(d,dbo.meetings.meeting_date, convert(smalldatetime,'"&start_date&"',103)) <= 0 "
	End If
	If trim(end_date) <> "" Then
		sqlstr = sqlstr & " AND DateDiff(d,dbo.meetings.meeting_date, convert(smalldatetime,'"&end_date&"',103)) >= 0 "
	End If
	If trim(search_company) <> "" Then
		sqlstr = sqlstr & " And dbo.COMPANIES.COMPANY_NAME Like '%" & sFix(trim(search_company)) & "%'"
	End If
	If trim(search_contact) <> "" Then
		sqlstr = sqlstr & " And dbo.CONTACTS.CONTACT_NAME Like '%" & sFix(trim(search_contact)) & "%'"
	End If	
	If trim(search_project) <> "" Then
		sqlstr = sqlstr & " And dbo.projects.project_name Like '%" & sFix(trim(search_project)) & "%'"
	End If	
	If trim(search_status) <> "" Then
		sqlstr = sqlstr & " And dbo.meetings.meeting_status IN (" & sFix(trim(search_status)) & ")"
	End If	
	sqlstr = sqlstr & " Order BY dbo.meetings.meeting_date, dbo.meetings.start_time, dbo.meetings.end_time "
	'Response.Write sqlStr
	'Response.End
	set rs_meetings = con.GetRecordSet(sqlStr)	
	do while not rs_meetings.eof
		meetingID = rs_meetings(0)
		meetingDate = rs_meetings(1) 
		startTime = rs_meetings(2)
		endTime = rs_meetings(3)
		company_name = rs_meetings(4)
		contact_name = rs_meetings(5)
		phone = rs_meetings(6)
		cellular = rs_meetings(7)
		status = rs_meetings(8)
		meetingContent = rs_meetings(9)
		meetingClosing = rs_meetings(10)
		project_name = rs_meetings(11)		
		UserName = rs_meetings(12) 
		
		meeting_participants = ""
		sqlstr = "exec dbo.get_meeting_participants '"&meetingID&"','"&OrgID&"'"
		set rs_meeting_participants = con.getRecordSet(sqlstr)
		If not rs_meeting_participants.eof Then
			meeting_participants = rs_meeting_participants.getString(,,"<br>","<br>")
		Else
			meeting_participants = ""
		End If	
		
		If Len(meeting_participants) > 4 Then
			meeting_participants = Left(meeting_participants, Len(meeting_participants)-4)
		End If
		
        strLine = "<tr><td class=xl26 x:str>" & arr_Status(status) & "</td><td class=xl31 x:str>" & meetingDate & "</td>"&_
        "<td class=xl30 x:str>" & startTime & "</td><td class=xl30 x:str>" & endTime & "</td>"&_
        "<td class=xl30 x:str>" & company_name & "</td><td class=xl30 x:str>" & contact_name & "</td>"&_
        "<td class=xl30 x:str>" & phone & "<br>" & cellular & "</td><td class=xl30 x:str>" & project_name & "</td>"&_
        "<td class=xl30 x:str>" & meeting_participants & "</td><td class=xl30 x:str>" & meetingContent & "</td>"&_
        "<td class=xl30 x:str>" & meetingClosing & "</td><td class=xl30 x:str>" & UserName & "</td></tr>"
	    XLSfile.writeline strLine 
     
 	  rs_meetings.movenext		
	 loop
 set rs_meetings = nothing	
set con=Nothing
XLSfile.WriteLine "</TABLE></body></html>"
XLSfile.close

set con=Nothing
Response.Redirect("../../../download/reports/"&filestring)
%>



<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%

start_date = trim(Request("dateStart"))		
end_date = trim(Request("dateEnd"))		
UserID = trim(Request("curr_user_id"))
CompanyID = trim(Request("company_id"))
ProjectID = trim(Request("project_id"))
MechID = trim(Request("mechanism_id"))

If trim(UserID) <> "" Then
	sqlstr = "Select FirstName, LastName From Users WHERE User_ID = " & UserID
	set rs_user = con.getRecordSet(sqlstr)
	if not rs_user.eof then
		userName = "<br>&nbsp;עובד:&nbsp;<font color=""#666699"">" & trim(rs_user(0)) & " " & trim(rs_user(1))& "</font>"
	end if
	set rs_user = nothing
End If

If trim(CompanyID) <> "" Then
	sqlstr = "Select Company_Name From Companies WHERE Company_ID = " & CompanyID
	set rs_com = con.getRecordSet(sqlstr)
	if not rs_com.eof then
		companyName = "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & ":&nbsp;<font color=""#666699"">" & trim(rs_com(0)) & "</font>"
	end if
	set rs_com = nothing
End If

If trim(ProjectID) <> "" Then
	sqlstr = "Select Project_Name From Projects WHERE Project_ID = " & ProjectID
	set rs_project = con.getRecordSet(sqlstr)
	if not rs_project.eof then
		projectName = "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("Projectone")) & ":&nbsp;<font color=""#666699"">" & trim(rs_project(0)) & "</font>"
	end if
	set rs_project = nothing
End If

If trim(MechID) <> "" Then
	sqlstr = "Select mechanism_name, project_name From mechanism Inner Join Projects On mechanism.project_id = projects.project_id WHERE mechanism_id = " & MechID
	set rs_mech = con.getRecordSet(sqlstr)
	if not rs_mech.eof then
		mechName = "<br>&nbsp; מנגנון:&nbsp;<font color=""#666699"">" & trim(rs_mech(1)) & " - " & trim(rs_mech(0)) & "</font>"
	end if
	set rs_mech = nothing
End If

set fsotmp = createobject("scripting.filesystemobject")
fsotmp.DeleteFile server.MapPath("../../../download/reports/") & "\*.*",true 
set fsotmp = nothing
		
set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
filestring="report_user_"&Day(Date())&"_"&Month(Date())&"_"&Year(Date())&".xls"
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
XLSfile.WriteLine "	mso-number-format:Fixed"
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
XLSfile.WriteLine "<x:Name>דו''ח אישי</x:Name>"
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
XLSfile.WriteLine "<body>"
XLSfile.WriteLine "<table style='border-collapse:collapse;table-layout:fixed'>"
XLSfile.WriteLine "<col width=250>"
XLSfile.WriteLine "<col width=90>"
XLSfile.WriteLine "<col width=90>"
XLSfile.WriteLine "<col width=90>"
XLSfile.WriteLine "<col width=90>"
XLSfile.WriteLine "<col width=90>"
XLSfile.WriteLine "<col width=90>"
XLSfile.WriteLine "<col width=90>"
strLine="<tr><td colspan=6 class=xl24> דו''ח אישי לתקופה <br> "& start_date & " - " & end_date
strLine = strLine & userName & companyName & projectName & mechName & "</td></tr>"
XLSfile.writeline strLine

strLine="<tr><td></td><td></td><td></td><td></td></tr>"
XLSfile.writeline strLine
strLine="<tr><td class=xl25>&nbsp; " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & " / " & trim(Request.Cookies("bizpegasus")("Projectone")) & " / מנגנון&nbsp;</td><td class=xl25>&nbsp;תאריך&nbsp;</td><td class=xl25>&nbsp;שעות&nbsp;</td><td class=xl25>&nbsp;עלות שעה&nbsp;</td><td class=xl25>&nbsp;סה''כ עלות&nbsp;</td></tr>"
XLSfile.writeline strLine

con.executeQuery("SET DATEFORMAT dmy")
sqlstr = "Select project_name,company_name,mechanism_name,minuts,hour_id,date,hour_pay from hours_view Where ORGANIZATION_ID = "& OrgID
If trim(UserId) <> "" Then
	sqlstr = sqlstr & " AND USER_ID = " & UserId
End If
If trim(CompanyID) <> "" Then
	sqlstr = sqlstr & " AND Company_ID = " & CompanyID
End If
If trim(ProjectId) <> "" Then
	sqlstr = sqlstr & " AND Project_ID = " & ProjectId
End If
If trim(MechId) <> "" Then
	sqlstr = sqlstr & " AND mechanism_id = " & MechId
End If	
If isDate(start_date) Then
	sqlstr = sqlstr & " And DateDiff(d,date,'"&start_date&"') <= 0"
End If	
If isDate(end_date) Then
	sqlstr = sqlstr & " And DateDiff(d,date,'"&end_date&"') >= 0"			
End If
sqlstr = sqlstr & " Order By date desc,company_name,project_name"
'Response.Write sqlstr
'Response.End
set rs_pr=con.getRecordSet(sqlstr)
if not rs_pr.eof then
	recordCount = rs_pr.RecordCount
do while not rs_pr.EOF	      
    company_name = trim(rs_pr(1))
	project_name = trim(rs_pr(0))
	mechanism_name = trim(rs_pr(2))
	minuts_pr = trim(rs_pr(3))
	hour_id = trim(rs_pr(4))
	date_pr = trim(rs_pr(5))
	hour_pay = trim(rs_pr(6))
	hours_pr = (minuts_pr / 60)
	
	strLine = "<tr><td class=xl26 dir=rtl>" & company_name & " / " & project_name  & " / " & mechanism_name & "</td><td class=xl31>" & date_pr & "</td><td class=xl30 x:num>" & Round(hours_pr, 1) & "</td><td class=xl30 x:num>" & hour_pay & "</td><td class=xl30 x:num>" & hour_pay * hours_pr & "</td></tr>"
	XLSfile.writeline strLine
  
rs_pr.moveNext
Loop
set list = Nothing  
   

strLine = "<tr><td class=xl27></td><td class=xl27>&nbsp;סה''כ&nbsp;</td><td class=xl27 x:num x:fmla=""=SUM(C4:C" & (3 + cInt(recordCount)) & ")"">&nbsp;</td><td class=xl27>&nbsp;&nbsp;</td><td class=xl27 x:num x:fmla=""=SUM(E4:E" & (3 + cInt(recordCount)) & ")"">&nbsp;</td></tr>"

XLSfile.writeline strLine    

End If
XLSfile.WriteLine "</TABLE></body></html>"
XLSfile.close

set con=Nothing
Response.Redirect("../../../download/reports/"&filestring)
%>


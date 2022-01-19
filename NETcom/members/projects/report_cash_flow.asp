<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%

start_date = trim(Request("start_date"))		
end_date = trim(Request("end_date"))		
bank_id = trim(Request("bank_id"))
balance = trim(Request("balance"))
type_id = trim(Request("type_id"))

If IsNumeric(balance) = false Then
	balance = 0
End If

where_dates = " AND payment_date BETWEEN '" & start_date & "' AND '" & end_date & "'"
where_bank = " AND bank_id = " & bank_id

If trim(bank_id) <> "" Then
	sqlStr = "select bank_name, credit from banks where ORGANIZATION_ID = "& OrgID & " AND bank_id = " & bank_id
	set rs_banks = con.getRecordSet(sqlStr)
	if not rs_banks.eof then	
		bank_Name = "<br>&nbsp;חשבון בנק:&nbsp;<font color=""#666699"">" & trim(rs_banks(0))& "</font>"
		credit = rs_banks(1)
	end if
	set rs_banks = nothing
End If

Select Case type_id
	Case "1" : title = "תזרים מזומנים תנועות מתוכננות" 
	Case "2" : title = "תזרים מזומנים תנועות מבוצעות" 
End Select

set fsotmp = createobject("scripting.filesystemobject")
fsotmp.DeleteFile server.MapPath("../../../download/reports/") & "\*.*",true 
set fsotmp = nothing
		
set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
filestring="report_cash_flow_"&Day(Date())&"_"&Month(Date())&"_"&Year(Date())&".xls"
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
XLSfile.WriteLine "	text-align:right;"
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
XLSfile.WriteLine "	text-align:right;"
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
XLSfile.WriteLine "	text-align:right;"
XLSfile.WriteLine "	font-weight:600;"
XLSfile.WriteLine "	white-space:normal;"
XLSfile.WriteLine "	}"
XLSfile.WriteLine ".xl30"
XLSfile.WriteLine "	{mso-style-parent:style0;"
XLSfile.WriteLine "	font-family:Arial, sans-serif;"
XLSfile.WriteLine "	mso-font-charset:0;"
XLSfile.WriteLine "	mso-number-format:Standard;"
XLSfile.WriteLine "	text-align:right;"
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
XLSfile.WriteLine "	text-align:right;"
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
XLSfile.WriteLine "<x:Name>תזרים מזומנים</x:Name>"
XLSfile.WriteLine "<x:WorksheetOptions>"
XLSfile.WriteLine "<x:Print>"
XLSfile.WriteLine "      <x:ValidPrinterInfo/>"
XLSfile.WriteLine "      <x:PaperSizeIndex>9</x:PaperSizeIndex>"
XLSfile.WriteLine "      <x:HorizontalResolution>600</x:HorizontalResolution>"
XLSfile.WriteLine "      <x:VerticalResolution>600</x:VerticalResolution>"
XLSfile.WriteLine "</x:Print>"
XLSfile.WriteLine "<x:DisplayRightToLeft/>"
XLSfile.WriteLine "<x:Selected/>"
XLSfile.WriteLine " <x:Panes>"
XLSfile.WriteLine "   <x:Pane>"
XLSfile.WriteLine "       <x:Number>3</x:Number>"
XLSfile.WriteLine "       <x:ActiveRow>9</x:ActiveRow>"
XLSfile.WriteLine "       <x:ActiveCol>5</x:ActiveCol>"
XLSfile.WriteLine "   </x:Pane>"
XLSfile.WriteLine " </x:Panes>"
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
XLSfile.WriteLine "<table>"
strLine="<tr><td colspan=6 class=xl24>" & title & "<br>" & end_date & " - " & start_date & bank_Name & "</td></tr>"
XLSfile.writeline strLine

If  cDbl(balance) > 0  Then
	style_ = "color:green"	
Else						
	style_ = "color:red"		
End If	
	
strLine="<tr><td></td><td></td><td></td><td></td></tr>"
XLSfile.writeline strLine
strLine="<tr><td class=xl24 colspan=2>יתרה התחלתית</td><td class=xl24 style=""mso-number-format:Standard;text-align:right;"&style_&""">"&balance&"</td><td></td><td></td></tr>"
XLSfile.writeline strLine
strLine="<tr><td></td><td></td><td></td><td></td></tr>"
XLSfile.writeline strLine

strLine="<tr><td class=xl25>&nbsp;תאריך&nbsp;</td><td class=xl25>&nbsp;תנועה&nbsp;</td><td class=xl25>&nbsp;פרטי ביצוע&nbsp;</td><td class=xl25>&nbsp;&#8362;&nbsp;הוצאה&nbsp;</td><td class=xl25>&nbsp;&#8362;&nbsp;הכנסה&nbsp;</td><td class=xl25>&nbsp;יתרה&nbsp;</td><td></td><td class=xl25>&nbsp;גרעון/עודפי תזרים&nbsp;</td></tr>"
XLSfile.writeline strLine

con.executeQuery("SET DATEFORMAT dmy")
If trim(type_id) = "1" Then 'מתוכננות
	sqlStr = "select payment_date, type_id, payment, movement_type_name, movement_id, movement_details "&_
	" from movements_view where movements_view.ORGANIZATION_ID = "& OrgID & where_dates & where_bank &_
	" AND rank_id IN (0,1) order by payment_date"
	'Response.Write sqlStr
Else
	sqlStr = "select payment_date,type_id,payment,movement_type_name, movement_id, line_details "&_
	" from movement_lines_view where ORGANIZATION_ID = " & OrgID & where_dates & where_bank &_
	" order by payment_date" 
End If
set rs_payments = con.getRecordSet(sqlStr)
all_sum = 0
remainder = 0
while not rs_payments.eof 
	payment_date = trim(rs_payments(0))
	typeId = trim(rs_payments(1))
	payment = trim(rs_payments(2)) 	
	movement_type_name = trim(rs_payments(3)) 
	movement_id = trim(rs_payments(4))		
	movement_details = trim(rs_payments(5))
	If trim(typeId) = "2" Then
		style_line = "style=color:green"
		all_sum = all_sum + CDbl(payment)
		remainder = CDbl(balance) + CDbl(all_sum)
		payment_ = FormatNumber(payment,2,-1,0,-1)
		all_sum_ = FormatNumber(all_sum,2,-1,0,-1)
		remainder_ = FormatNumber(remainder,2,-1,0,-1)	
	Else						
		style_line = "style=color:red"
		payment = 0 - payment
		all_sum = all_sum + CDbl(payment)
		remainder = CDbl(balance) + CDbl(all_sum)
		payment_ = FormatNumber(payment,2,-1,0,-1)
		all_sum_ = FormatNumber(all_sum,2,-1,0,-1)
		remainder_ = FormatNumber(remainder,2,-1,0,-1)					
	End If		

	strLine = "<tr><td class=xl26>" & CDate(DateValue(payment_date)) & "</td><td class=xl26>" & movement_type_name & "</td><td class=xl31>" & movement_details & "</td>"
	If trim(typeId) = "2" Then
	strLine = strLine & "<td class=xl30></td><td class=xl30 x:num>" & payment_ & "</td>"
	Else
	strLine = strLine & "<td class=xl30 x:num>" & payment_ & "</td><td class=xl30></td>"
	End If
	strLine = strLine & "<td class=xl30 x:num>" & remainder_ & "</td><td></td><td class=xl30 x:num>" & all_sum_ & "</td></tr>"
	XLSfile.writeline strLine
	
 rs_payments.moveNext
 Wend
set rs_payments = Nothing

strLine="<tr><td></td><td></td><td></td><td></td></tr>"

If cDbl(remainder) > 0 Then
	style_ = "color:green"	
Else						
	style_ = "color:red"		
End If	

XLSfile.writeline strLine
strLine="<tr><td class=xl24 colspan=2>יתרה סופית</td><td class=xl24 style=""mso-number-format:Standard;text-align:right;"&style_&""">"&remainder_&"</td><td></td><td></td></tr>"
XLSfile.writeline strLine
strLine="<tr><td></td><td></td><td></td><td></td></tr>"
XLSfile.writeline strLine

XLSfile.WriteLine "</table></body></html>"
XLSfile.close

set con=Nothing
Response.Redirect("../../../download/reports/"&filestring)
%>


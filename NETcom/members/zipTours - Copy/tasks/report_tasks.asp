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
		arr_Status = Array("","חדש","בטיפול","סגור")	
	else
		arr_Status = Array("","new","active","close")	
	end if  

 task_status = trim(Request.Form("task_status")) 
 UserID = trim(Request.Form("User_ID"))  
 companyID = trim(Request.Form("company_id"))   
 T = trim(Request.Form("T"))
 
 If trim(T) = "OUT" Then 'לפתוח משימות יוצאות סגורות	
    sender_id = UserID		
	where_sender = " AND user_id = " & sender_id
 ElseIf trim(T) = "IN" Then ' לפתוח משימות נכנסות פתוחות	
    reciver_id = UserID			
    where_reciver = " AND reciver_id = " & reciver_id
 End If
 
 where_status = " And (task_status in (" & sFix(task_status) & "))"
 
 If Request("date_start") <> nil Then	
	date_start = Request("date_start")
	start_date_ = Month(date_start) & "/" & Day(date_start) & "/" & Year(date_start)   
 Else
	start_date_ = ""  
 End If

 If Request("date_end") <> nil Then
	date_end = Request("date_end")
    end_date_ = Month(date_end) & "/" & Day(date_end) & "/" & Year(date_end)
 Else
    end_date_ = ""
 End If 
 
 if trim(companyID) <> "" Then			
	 where_company = " And company_ID  = " & companyID
 else
	 where_company = ""		
 end if
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
set fsotmp = createobject("scripting.filesystemobject")
fsotmp.DeleteFile server.MapPath("../../../download/reports/") & "\*.*",true 
set fsotmp = nothing
		
	set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
	filestring="report_tasks_"&Day(Date())&"_"&Month(Date())&"_"&Year(Date())&".xls"
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
	XLSfile.WriteLine "	height:16pt;"
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
	XLSfile.WriteLine "<x:Name>דו''ח משימות</x:Name>"
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
	XLSfile.WriteLine "<col width=60>"
	XLSfile.WriteLine "<col width=80>"
	XLSfile.WriteLine "<col width=80>"
	XLSfile.WriteLine "<col width=100>"
	XLSfile.WriteLine "<col width=100>"
	XLSfile.WriteLine "<col width=150>"
	XLSfile.WriteLine "<col width=120>"
	XLSfile.WriteLine "<col width=150>"
	XLSfile.WriteLine "<col width=150>"
	XLSfile.WriteLine "<col width=280>"
	If trim(lang_id) = "1" Then
	strLine="<tr><td colspan=7 class=xl24> דו''ח משימות לתקופה <br> "& date_start & " - " & date_end
	Else
	strLine="<tr><td colspan=7 class=xl24> Tasks report in period <br> "& date_start & " - " & date_end
	End If
	strLine = strLine & userName & companyName & projectName & "</td></tr>"
	XLSfile.writeline strLine

	strLine="<tr><td></td><td></td><td></td><td></td></tr>"
	XLSfile.writeline strLine
	If trim(lang_id) = "1" Then
	strLine="<tr><td class=xl25>&nbsp;סטטוס&nbsp;</td><td class=xl25>&nbsp;תאריך&nbsp;</td><td class=xl25>&nbsp;תאריך יעד&nbsp;</td>"&_
	"<td class=xl25>&nbsp;מאת&nbsp;</td><td class=xl25>&nbsp;אל&nbsp;</td>"&_
	"<td class=xl25>&nbsp;"&trim(Request.Cookies("bizpegasus")("CompaniesOne"))&"&nbsp;</td>"&_
	"<td class=xl25>&nbsp;"&trim(Request.Cookies("bizpegasus")("ContactsOne"))&"&nbsp;</td>"&_
	"<td class=xl25>&nbsp;"&trim(Request.Cookies("bizpegasus")("Projectone"))&"&nbsp;</td>"&_
	"<td class=xl25>&nbsp;סוג&nbsp;</td><td class=xl25>&nbsp;תוכן&nbsp;</td>"&_
	"<td class=xl25>&nbsp;תאריך סגירה&nbsp;</td><td class=xl25>&nbsp;תוכן סגירה&nbsp;</td></tr>"
	Else
	strLine="<tr><td class=xl25>&nbsp;Status&nbsp;</td><td class=xl25>&nbsp;Date&nbsp;</td><td class=xl25>&nbsp;Target date&nbsp;</td>"&_
	"<td class=xl25>&nbsp;From&nbsp;</td><td class=xl25>&nbsp;To&nbsp;</td>"&_
	"<td class=xl25>&nbsp;"&trim(Request.Cookies("bizpegasus")("CompaniesOne"))&"&nbsp;</td>"&_
	"<td class=xl25>&nbsp;"&trim(Request.Cookies("bizpegasus")("ContactsOne"))&"&nbsp;</td>"&_
	"<td class=xl25>&nbsp;"&trim(Request.Cookies("bizpegasus")("Projectone"))&"&nbsp;</td>"&_
	"<td class=xl25>&nbsp;Type&nbsp;</td><td class=xl25>&nbsp;Content&nbsp;</td>"&_
	"<td class=xl25>&nbsp;Closing date&nbsp;</td><td class=xl25>&nbsp;Closing Content&nbsp;</td></tr>"	
	End If
	XLSfile.writeline strLine
  
	dim sortby(12)	
	If trim(T) = "OUT" Then 
	sortby(0) = "task_date DESC,task_status,task_id DESC"
	Else
	sortby(0) = "task_status,task_date DESC,task_id DESC"
	End If
	sort = 0
     
   Set tasksList = Server.CreateObject("ADODB.RECORDSET")  
   sqlstr = "EXECUTE get_tasks '','','','" & task_status & "','" & OrgID & "','','" & reciver_id & "','" & sender_id & "','" & sortby(sort) & "','" & start_date_ & "','" & end_date_ & "','" & companyID & "'"
   'Response.Write sqlstr
   'Response.End
   set  tasksList = con.getRecordSet(sqlstr)
   current_date = Day(date()) & "/" & Month(date()) & "/" & Year(date())
   dim IS_DESTINATION 
   If not tasksList.EOF then		
		arrtasks = tasksList.getRows()
		tasksList.Close()
		set tasksList = Nothing	
       
	do while i <= uBound(arrtasks,2)            
       companyName =  trim(arrtasks(0,i))
       'Response.Write companyName
       'Response.End
       contactName =  trim(arrtasks(1,i))
       projectName = trim(arrtasks(2,i))  
       reciver_name = trim(arrtasks(3,i))
	   sender_name = trim(arrtasks(4,i))
       task_types = trim(arrtasks(6,i))
       companyID =  trim(arrtasks(11,i))
       contactID = trim(arrtasks(12,i))
       taskId = trim(arrtasks(13,i))
       status = trim(arrtasks(17,i))
       task_replay = trim(arrtasks(18,i))
       task_close_date = trim(arrtasks(19,i))
       task_content = trim(arrtasks(8,i))   
       task_sender = trim(arrtasks(9,i))
       parentID = trim(arrtasks(15,i))         
       If IsDate(trim(arrtasks(21,i))) Then
			task_open_date = Day(trim(arrtasks(21,i))) & "/" & Month(trim(arrtasks(21,i))) & "/" & Right(Year(trim(arrtasks(21,i))),2)
       Else
		    task_open_date = ""
       End If
       If IsDate(trim(arrtasks(7,i))) Then
		  task_date=Day(trim(arrtasks(7,i))) & "/" & Month(trim(arrtasks(7,i))) & "/" & Right(Year(trim(arrtasks(7,i))),2)
		  if DateDiff("d",task_date,current_date) >= 0 then
			IS_DESTINATION = true
		  else
			IS_DESTINATION = false
		  end if
	   Else
		   task_date=""
		   IS_DESTINATION = false
	   End If     				
	  
		sqlstr = "Exec dbo.get_task_types '"&taskID&"','"&OrgID&"'"
		set rs_task_types = con.getRecordSet(sqlstr)
		If not rs_task_types.eof Then
			task_types_names = rs_task_types.getString(,,",",",")
		Else
			task_types_names = ""
		End If		
		
		If Len(task_types_names) > 0 Then
			task_types_names = Left(task_types_names,(Len(task_types_names)-1))
		End If

       strLine = "<tr><td class=xl26>" & arr_Status(status)  & "</td><td class=xl31>" & task_open_date &_
       "</td><td class=xl31>" & task_date & "</td><td class=xl30>" & sender_name & "</td><td class=xl30>" &_
       reciver_name & "</td><td class=xl30>" & companyName & "</td><td class=xl30>" & contactName &_
       "</td><td class=xl30>" & projectName & "</td><td class=xl30>" & task_types_names & "</td><td class=xl30>" &_
       task_content & "</td><td class=xl30 x:str>" & task_close_date & "</td><td class=xl30>" & task_replay & "</td></tr>"
	   XLSfile.writeline strLine 
     
	  i=i+1	
	  loop
 End If
set con=Nothing
XLSfile.WriteLine "</TABLE></body></html>"
XLSfile.close

set con=Nothing
Response.Redirect("../../../download/reports/"&filestring)
%>



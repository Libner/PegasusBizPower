<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 39 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
	arrTitlesD = rstitle.getRows()		
	redim arrTitles(Ubound(arrTitlesD,2)+1)
	For i=0 To Ubound(arrTitlesD,2)		
		arrTitles(arrTitlesD(0,i)) = arrTitlesD(1,i)		
	Next
	End If
	set rstitle = Nothing

	set fsotmp = createobject("scripting.filesystemobject")
	fsotmp.DeleteFile server.MapPath("../../../download/reports/") & "\*.*",true 
	set fsotmp = nothing
		
	set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
	filestring="report_workers_"&Day(Date())&"_"&Month(Date())&"_"&Year(Date())&".xls"
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
	XLSfile.WriteLine "<x:Name>דו''ח עובדים</x:Name>"
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
	XLSfile.WriteLine "<col width=100>"
	XLSfile.WriteLine "<col width=100>"
	XLSfile.WriteLine "<col width=100>"
	XLSfile.WriteLine "<col width=100>"
	XLSfile.WriteLine "<col width=100>"
	XLSfile.WriteLine "<col width=100>"
	XLSfile.WriteLine "<col width=100>"
	XLSfile.WriteLine "<col width=100>"
	XLSfile.WriteLine "<col width=200>"
	XLSfile.WriteLine "<col width=80>"
	XLSfile.WriteLine "<col width=200>"
	XLSfile.WriteLine "<col width=200>"
	XLSfile.WriteLine "<col width=500>"
	strLine="<tr><td colspan=10 class=xl24> " & org_name & " - " & arrTitles(25) & "</td></tr>"
	XLSfile.writeline strLine

	strLine="<tr><td></td><td></td><td></td><td></td></tr>"
	XLSfile.writeline strLine
	
	strLine="<tr><td class=xl25>&nbsp;"& arrTitles(3) &"&nbsp;</td><td class=xl25>&nbsp;"& arrTitles(4) &"&nbsp;</td>"&_
	"<td class=xl25>&nbsp;"& arrTitles(5) &"&nbsp;</td><td class=xl25>&nbsp;"& arrTitles(6) &"&nbsp;</td>"&_
	"<td class=xl25>&nbsp;"& arrTitles(7) &"&nbsp;</td><td class=xl25>&nbsp;"& arrTitles(8) &"&nbsp;</td>"&_
	"<td class=xl25>&nbsp;"& arrTitles(9) &"&nbsp;</td><td class=xl25>&nbsp;"& arrTitles(10) &"&nbsp;</td>"&_
	"<td class=xl25>&nbsp;"& arrTitles(11) &"&nbsp;</td><td class=xl25>&nbsp;יעד הזמנות מינימלי&nbsp;</td>"&_
	"<td class=xl25>&nbsp;"& arrTitles(14) &"&nbsp;</td><td class=xl25>&nbsp;"& arrTitles(17) &"&nbsp;</td><td class=xl25>&nbsp;"& arrTitles(13) &"&nbsp;</td></tr>"
	
	XLSfile.writeline strLine
  
  sqlStr = "select USER_ID,FIRSTNAME,LASTNAME, LOGINNAME, PASSWORD,OFFICE,TELEPHONE,MOBILE,FAX,"&_
  "EMAIL,job_name,Month_Min_Order FROM USERS_VIEW() where ORGANIZATION_ID=" & OrgID  
  ''Response.Write sqlStr
  set rs_USERS = con.GetRecordSet(sqlStr)
  while not rs_USERS.eof
		USER_ID   = rs_USERS("USER_ID")
		FIRSTNAME = rs_USERS("FIRSTNAME")
		LASTNAME  = rs_USERS("LASTNAME")
		LOGINNAME = rs_USERS("LOGINNAME")	
		PASSWORD_U  = rs_USERS("PASSWORD")					
		office    = rs_USERS("OFFICE")
        phone     = rs_USERS("TELEPHONE")
        mobile    = rs_USERS("MOBILE")
        fax       = rs_USERS("FAX")
        email     = rs_USERS("EMAIL")
        job_name  = rs_USERS("job_name")
        Month_Min_Order = rs_USERS("Month_Min_Order")
		Groups=""		
		sqlstr="Select Group_Name From Users_to_Groups_view Where User_id = " & USER_ID & " Order By Group_id"
		set rssub = con.getRecordSet(sqlstr)		   
		If not rssub.eof Then
			Groups = rssub.getString(,,",",",")		
		Else	
			Groups = ""
		End If		    
		set rssub=Nothing
		If Len(Groups) > 0 Then
			Groups = Left(Groups,(Len(Groups)-1))
		End If	
		
		ResInGroups=""		
		sqlstr="Select Group_Name From Responsibles_to_Groups_view Where Responsible_id = " & USER_ID & " Order By Group_id"
		set rssub = con.getRecordSet(sqlstr)		   
		If not rssub.eof Then
			ResInGroups = rssub.getString(,,",",",")
		Else
			ResInGroups=""
		End If		    
		set rssub=Nothing
		If Len(ResInGroups) > 0 Then
			ResInGroups = Left(ResInGroups,(Len(ResInGroups)-1))
		End If    
		
		sql_obj="Select bar_title From dbo.bar_users_table('" & OrgID & "','" & UserID & "') WHERE "&_
		" PARENT_ID IS NOT NULL AND IS_VISIBLE = '1' Order by PARENT_Order, bar_Order"
		set rs_obj=con.getRecordSet(sql_obj)
		If not rs_obj.eof Then
			PermitionsList = rs_obj.getString(,,",",",")	
		Else
			PermitionsList=""
		End If		    
		set rs_obj=Nothing
		If Len(PermitionsList) > 0 Then
			PermitionsList = Left(PermitionsList,(Len(PermitionsList)-1))
		End If    

       	strLine = "<tr><td class=xl26>" & FIRSTNAME  & "</td><td class=xl26>" & LASTNAME  & "</td>" &_
       	"<td class=xl26>" & LOGINNAME  & "</td><td class=xl26>" & PASSWORD_U  & "</td>" &_
       	"<td class=xl26>" & job_name  & "</td><td class=xl26>" & phone  & "</td>" &_
       	"<td class=xl26>" & mobile  & "</td><td class=xl26>" & fax  & "</td>" &_
       	"<td class=xl26>" & email  & "</td><td class=xl26>" & Month_Min_Order  & "</td>" &_
       	"<td class=xl26>" & Groups  & "</td><td class=xl26>" & ResInGroups  & "</td>" &_
       	"<td class=xl26>" & PermitionsList  & "</td></tr>"
	    XLSfile.writeline strLine 
     
		rs_USERS.MoveNext
	Wend	
	set rs_USERS=nothing
set con=Nothing
XLSfile.WriteLine "</TABLE></body></html>"
XLSfile.close

set con=Nothing
Response.Redirect("../../../download/reports/"&filestring)
%>


<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
</head>
<body BGCOLOR="c0c0c0"><br><br><br><br>
<H4  align=center> Please wait ... </H4>
<!--#include file="code.asp" -->
<% Server.ScriptTimeout = 100000 

   Function nFix(myString)
		If isNumeric(nFix) And Len(myString) > 0 And trim(myString) <> "" Then
			nFix = "'" & trim(myString) & "'"
		Else
			nFix = "NULL"
		End If	
	End Function  
  
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
  
	quest_id = trim(Request("quest_id"))
	start_date = trim(Request("dateStart"))		
	end_date = trim(Request("dateEnd"))
	If IsDate(start_date) Then
		start_date_ = Month(start_date) & "/" & Day(start_date) & "/" & Year(start_date)
	Else
		start_date_ = ""	 
	End If
	If IsDate(end_date) Then
		end_date_ = Month(end_date) & "/" & Day(end_date) & "/" & Year(end_date)
	Else
		end_date_ = ""	 
	End If
	
	UserID = trim(Request("user_id")) 
	CompanyID = trim(Request("company_id"))
	OrgID = trim(Request.Cookies("bizpegasus")("ORGID"))
    
	If isNumeric(quest_id) Then
		sqlstr="Select product_name From PRODUCTS WHERE PRODUCT_ID = " & quest_id
		set rsn = con.getRecordSet(sqlstr)
		If not rsn.eof Then
			prodName=trim(rsn(0))		
		End If
		set rsn = Nothing
	End If
	
  	Dim strXLS
	Set strXLS = New StrConCatArray 
  
    strXLS.Add "<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
	strXLS.Add "<head>"
	strXLS.Add "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1255"">"
	strXLS.Add "<meta name=ProgId content=Excel.Sheet>"
	strXLS.Add "<meta name=Generator content=""Microsoft Excel 9"">"
	strXLS.Add "<style>"
	strXLS.Add "<!--table"
	strXLS.Add "	{mso-displayed-decimal-separator:""\."";"
	strXLS.Add "	mso-displayed-thousand-separator:""\,"";}"
	strXLS.Add "@page"
	strXLS.Add "{ margin:1.0in .75in 1.0in .75in;"
	strXLS.Add "	mso-header-margin:.5in;"
	strXLS.Add "	mso-footer-margin:.5in;"
	strXLS.Add " mso-page-orientation:landscape;}"
	strXLS.Add "tr"
	strXLS.Add "	{mso-height-source:auto;}"
	strXLS.Add "br"
	strXLS.Add "	{mso-data-placement:same-cell;}"
	strXLS.Add ".style0"
	strXLS.Add "	{mso-number-format:General;"
	strXLS.Add "	vertical-align:bottom;"
	strXLS.Add "	white-space:nowrap;"
	strXLS.Add "	mso-rotate:0;"
	strXLS.Add "	mso-background-source:auto;"
	strXLS.Add "	mso-pattern:auto;"
	strXLS.Add "	color:windowtext;"
	strXLS.Add "	font-size:10.0pt;"
	strXLS.Add "	font-weight:400;"
	strXLS.Add "	font-style:normal;"
	strXLS.Add "	text-decoration:none;"
	strXLS.Add "	font-family:Arial;"
	strXLS.Add "	mso-generic-font-family:auto;"
	strXLS.Add "	mso-font-charset:0;"
	strXLS.Add "	border:none;"
	strXLS.Add "	mso-protection:locked visible;"
	strXLS.Add "	text-align:center;"
	strXLS.Add "	mso-style-name:Normal;"
	strXLS.Add "	mso-height-source:auto userset;"
	strXLS.Add "	height:20pt;}"
	strXLS.Add ".xl24"  ' title
	strXLS.Add "	{mso-style-parent:style0;"
	strXLS.Add "	font-family:Arial, sans-serif;"
	strXLS.Add "	font-size:12.0pt;"
	strXLS.Add "	font-weight:700;"
	strXLS.Add "	text-align:center;"
	strXLS.Add "	mso-height-source: userset;"
	strXLS.Add "	height:18pt;"
	strXLS.Add "	white-space:normal;}"
	strXLS.Add ".xl25" ' title cell
	strXLS.Add "	{mso-style-parent:style0;"
	strXLS.Add "	color:#000000;"
	strXLS.Add "	font-size:10.0pt;"
	strXLS.Add "	font-weight:700;"
	strXLS.Add "	font-family:Arial, sans-serif;"
	strXLS.Add "	text-align:"&align_var&";"
	strXLS.Add "	background:#ff9900;"
	strXLS.Add "	border: 0.5pt solid black;"
	strXLS.Add "	mso-pattern:auto none;"
	strXLS.Add "	mso-height-source: userset;"
	strXLS.Add "	height:16pt;"
	strXLS.Add " white-space:normal;}"
	strXLS.Add ".xl26"  ' data cells
	strXLS.Add "	{mso-style-parent:style0;"
	strXLS.Add "	font-weight:400;"
	strXLS.Add "	font-size:10.0pt;"
	strXLS.Add "	font-family:Arial, sans-serif;"
	strXLS.Add "	border: 0.5pt solid black;"
	strXLS.Add "	mso-height-source: userset;"
	strXLS.Add "	vertical-align:top;"
	strXLS.Add "	text-align:"&align_var&";}"
	strXLS.Add "-->"
	strXLS.Add "</style>"
	strXLS.Add "<!--[if gte mso 9]><xml>"
	strXLS.Add "<x:ExcelWorkbook>"
	strXLS.Add "<x:ExcelWorksheets>"
	strXLS.Add "<x:ExcelWorksheet>"
	strXLS.Add "<x:Name>" & prodName & "</x:Name>"
	strXLS.Add "<x:WorksheetOptions>"
	strXLS.Add "<x:Print>"
	strXLS.Add "      <x:ValidPrinterInfo/>"
	strXLS.Add "      <x:PaperSizeIndex>9</x:PaperSizeIndex>"
	strXLS.Add "      <x:HorizontalResolution>600</x:HorizontalResolution>"
	strXLS.Add "      <x:VerticalResolution>600</x:VerticalResolution>"
	strXLS.Add "</x:Print>"
	strXLS.Add "<x:Selected/>"
	strXLS.Add "<x:ProtectContents>False</x:ProtectContents>"
	strXLS.Add "<x:ProtectObjects>False</x:ProtectObjects>"
	strXLS.Add "<x:ProtectScenarios>False</x:ProtectScenarios>"
	strXLS.Add "<x:WorksheetOptions>"
	strXLS.Add "<x:Selected/>"
	If trim(lang_id) = "1" Then
	strXLS.Add "<x:DisplayRightToLeft/>"
	End If
	strXLS.Add "<x:Panes>"
	strXLS.Add "<x:Pane>"
	strXLS.Add "<x:Number>1</x:Number>"
	strXLS.Add "<x:ActiveRow>1</x:ActiveRow>"
	strXLS.Add "</x:Pane>"
	strXLS.Add "</x:Panes>"
	strXLS.Add "<x:ProtectContents>False</x:ProtectContents>"
	strXLS.Add "<x:ProtectObjects>False</x:ProtectObjects>"
	strXLS.Add "<x:ProtectScenarios>False</x:ProtectScenarios>"
	strXLS.Add "</x:WorksheetOptions>"
	strXLS.Add "</x:ExcelWorksheet>"
	strXLS.Add "</x:ExcelWorksheets>"
	strXLS.Add "<x:WindowTopX>360</x:WindowTopX>"
	strXLS.Add "<x:WindowTopY>60</x:WindowTopY>"
	strXLS.Add "</x:ExcelWorkbook>"
	strXLS.Add "</xml><![endif]-->"
	strXLS.Add "</head>"
	strXLS.Add "<body>"
	strXLS.Add "<table style='border-collapse:collapse;table-layout:fixed' dir='"&dir_var&"'>"  
	strXLS.Add "<tr><td colspan=7 class=xl24>" & trim(prodName) & "</td></tr>"
	strXLS.Add "<tr><td colspan=7 class=xl24>" & start_date & " - " & end_date & "</td></tr>"
	
	If trim(lang_id) = "1" Then
		strXLS.Add "<tr style='mso-height-source:userset;height:28.5pt'><td class=xl25>שם נרשם</td><td class=xl25>חברה</td><td class=xl25>עיסוק</td><td class=xl25>מייל הנרשם</td><td class=xl25>התקבל</td>"
	Else
		strXLS.Add "<tr style='mso-height-source:userset;height:28.5pt'><td class=xl25>Recipient</td><td class=xl25>Company</td><td class=xl25>Job</td><td class=xl25>Email</td><td class=xl25>Date</td>"
	End If
	  
	sql = "SELECT FIELD_TITLE FROM FORM_FIELD WHERE product_id = " & quest_id & " AND FIELD_TYPE <> '10' Order By FIELD_ORDER"
	set rst = con.getRecordSet(sql)  
	while not rst.eof      
		strXLS.Add "<td class=xl25>" & trim(trim(rst(0))) & "</td>"
	rst.moveNext
	Wend
	set rst=Nothing    
	strXLS.Add "<td class=xl25>ID</td></tr>"
 
 	Set con_pr = Server.CreateObject("AdoDB.Connection")
	con_pr.Open connString
	
	Set rs_command = Server.CreateObject("AdoDB.Command")
	rs_command.CommandText = "Exec dbo.get_feedbacks_report @OrgID='" & cInt(OrgID) & "', " & _
	"@UserId=" & nFix(UserId) & ", @quest_id='" & cInt(quest_id) & "', @lang_id='" & trim(lang_id) & "', " & _
	"@start_date='" & trim(start_date) & "', @end_date='" & end_date & "', @company_id=" & nFix(CompanyID)
	rs_command.ActiveConnection = con_pr 
	
	'Response.Write rs_command.CommandText
 
	Set rs_app = Server.CreateObject("ADODB.RecordSet")
	rs_app.CursorType = 0
	rs_app.LockType = 1	
	rs_app.Open(rs_command)
	'strXLS.Add "<tr>"	
	'for x = 0 to rs_app.Fields.count - 1
    '    strXLS.Add "<td class=xl25>" & trim(rs_app.Fields(x).Name) & "</td>"
    'next
	'strXLS.Add "</tr>"
	If Not rs_app.EOF Then	
		strXLS.Add "<TR><TD class=xl26 x:str>"
		strXLS.Add rs_app.GetString(,,"</TD><TD class=xl26 x:str>", "</TD></TR><TR><TD class=xl26 x:str>")
		strXLS.Add "</TD></TR>"
	End If
	Set rs_app = Nothing
	Set rs_command = Nothing
	Set con_pr = Nothing
	
	strXLS.Add "</TABLE></body></html>"    			
	set con = Nothing	
	
	filestring = "export_answers_" & quest_id & ".xls"

    Response.Clear()
    Response.AddHeader("content-disposition", "inline; filename=" & filestring)
    Response.ContentType = "application/vnd.ms-excel"
    Response.Write(strXLS.value)  
    Set strXLS = Nothing %>
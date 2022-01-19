<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
</head>
<body BGCOLOR="c0c0c0"><br><br><br><br>
<H4  align=center> Please wait ... </H4>
<%
  
    Function enterFix(word)
	    word=replace("" & word,chr(13),Space(1))
	    word=replace("" & word,chr(10),Space(1))	 
	    enterFix=word
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
	
  set fs = createobject("scripting.filesystemobject")
  
  set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
  filestring="../../../download/reports/"
  'Response.Write server.mappath(filestring)
  'Response.End
  fs.DeleteFile server.mappath(filestring) & "/*.*" ,False
  
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
  ProjectID = trim(Request("project_id"))
  OrgID = trim(Request.Cookies("bizpegasus")("ORGID"))
  filestring="../../../download/reports/export_answers_" & trim(quest_id) & ".xls"
  'Response.Write server.mappath(filestring)
  'Response.End
  set XLSfile=fs.CreateTextFile(server.mappath(filestring)) ' Create text file as excel file
    
  If isNumeric(quest_id) Then
	sqlstr="Select product_name From PRODUCTS WHERE PRODUCT_ID = " & quest_id
	set rsn = con.getRecordSet(sqlstr)
	If not rsn.eof Then
		prodName=trim(rsn(0))		
	End If
	set rsn = Nothing
  End If
  
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
	XLSfile.WriteLine "	text-align:"&align_var&";"
	XLSfile.WriteLine "	background:#ff9900;"
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
	XLSfile.WriteLine "	vertical-align:top;"
	XLSfile.WriteLine "	text-align:"&align_var&";}"
	XLSfile.WriteLine "-->"
	XLSfile.WriteLine "</style>"
	XLSfile.WriteLine "<!--[if gte mso 9]><xml>"
	XLSfile.WriteLine "<x:ExcelWorkbook>"
	XLSfile.WriteLine "<x:ExcelWorksheets>"
	XLSfile.WriteLine "<x:ExcelWorksheet>"
	XLSfile.WriteLine "<x:Name>" & prodName & "</x:Name>"
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
	XLSfile.WriteLine "<x:WorksheetOptions>"
	XLSfile.WriteLine "<x:Selected/>"
	If trim(lang_id) = "1" Then
	XLSfile.WriteLine "<x:DisplayRightToLeft/>"
	End If
	XLSfile.WriteLine "<x:Panes>"
	XLSfile.WriteLine "<x:Pane>"
	XLSfile.WriteLine "<x:Number>1</x:Number>"
	XLSfile.WriteLine "<x:ActiveRow>1</x:ActiveRow>"
	XLSfile.WriteLine "</x:Pane>"
	XLSfile.WriteLine "</x:Panes>"
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

  
  strLine="<tr style='background:#808080;color:#ffffff'><td colspan=7 class=xl24 style='background:#808080;color:#ffffff'> " &  trim(prodName) & "</td></tr>"
  strLine=strLine&"<tr style='background:#808080;color:#ffffff'><td colspan=7 class=xl24 style='background:#808080;color:#ffffff'> " & start_date & " - " & end_date & "</td></tr>"
  XLSfile.writeline strLine 
  
'  strLine="<tr><td></td><td></td><td></td><td></td></tr>"
'  XLSfile.writeline strLine
  
  If trim(lang_id) = "1" Then
  strLine="<tr style='background:#ff9900;mso-height-source:userset;height:28.5pt'><td class=xl25>שם נרשם</td><td class=xl25>חברה</td><td class=xl25>עיסוק</td><td class=xl25>מייל הנרשם</td><td class=xl25>התקבל</td>"
  Else
  strLine="<tr style='background:#ff9900;mso-height-source:userset;height:28.5pt'><td class=xl25>Recipient</td><td class=xl25>Company</td><td class=xl25>Job</td><td class=xl25>Email</td><td class=xl25>Date</td>"
  End If
  
  sql="SELECT FIELD_TITLE FROM FORM_FIELD WHERE product_id = " & quest_id & " AND FIELD_TYPE <> '10' Order By FIELD_ORDER"
  set rst = con.getRecordSet(sql)  
  while not rst.eof      
  If Len(trim(rst(0))) > 0 Then
  strLine=strLine&"<td class=xl25>"&trim(trim(rst(0))) & "</td>"
  End If
  rst.moveNext
  Wend
  set rst=Nothing    
  strLine = strLine & "</tr>"
  XLSfile.writeline strLine	'write titles	
  strLine=""

  sql = "EXECUTE get_feedbacks '','','','','" & OrgID & "','APPEAL_DATE DESC','"& start_date_ &"','"& end_date_ &"','"& CompanyID &"','','"& ProjectID &"','','" & quest_id & "','"& UserID &"'"	
  'Response.Write sql
  'Response.End
  set rsa = con.getRecordSet(sql)
  while not rsa.eof
    appId = trim(rsa("APPEAL_ID"))
    appDate=trim(rsa("APPEAL_DATE"))
    email=trim(rsa("PEOPLE_EMAIL"))
    name=trim(rsa("PEOPLE_NAME"))
    company=trim(rsa("PEOPLE_COMPANY"))
    job=trim(rsa("PEOPLE_OFFICE"))
   
    If isNumeric(appId) Then
    strLine = "<tr><td class=xl26>" & name & "</td><td  class=xl26>" & company & "</td><td  class=xl26>" & job & "</td>" & _
    "<td  class=xl26>" & email & "</td><td  class=xl26>" & appDate & "</td>"
    sql="SELECT FIELD_ID FROM FORM_FIELD WHERE product_id = " & quest_id & " AND FIELD_TYPE <> '10' Order By FIELD_ORDER"
    set rs_fields = con.getRecordSet(sql)  
    while not rs_fields.eof  
		fieldID = trim(rs_fields(0))
        sql = "EXECUTE get_field_value '" & OrgID & "','"& fieldID &"','"& appId & "','" & quest_id & "',''"
		'Response.Write sql
		'Response.End
		set rs_value=con.GetRecordSet(sql)
		If not rs_value.EOF Then          		
			FIELD_TYPE=rs_value.Fields(6)
			FIELD_TITLE=rs_value.Fields(5)		
			FIELD_VALUE=rs_value.Fields(8)
			If trim(FIELD_TYPE) <> "5" Then
			strLine = strLine & "<td class=xl26 x:str>" & enterFix(trim(FIELD_VALUE)) & "</td>"       
			ElseIf trim(FIELD_VALUE) = "on" Then
			strLine = strLine & "<td class=xl26 x:str>" & enterFix(trim(FIELD_TITLE)) & "</td>"
			Else
			strLine = strLine & "<td class=xl26>&nbsp;</td>"
			End If
	    Else
			strLine = strLine & "<td class=xl26>&nbsp;</td>"       
	    End If
	    set rs_value = Nothing 	    
	rs_fields.moveNext
	Wend
	set rs_fields = Nothing    
    strLine = strLine & "</tr>"
	XLSfile.writeline strLine  
	strLine=""
	End If
	rsa.moveNext
Wend
set rsa = Nothing  
 XLSfile.WriteLine "</TABLE></body></html>"    
 XLSfile.close
 set con=Nothing
 Response.Redirect(filestring) 'Open the excel file in the browser
 set con= Nothing	 
 %>

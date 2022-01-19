<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<% Server.ScriptTimeout=20000
   Response.Buffer=True
   Response.CharSet = 1255
   	start_date = trim(Request("dateStart"))		
	end_date = trim(Request("dateEnd"))		

%>
<html>
<head>
<title>Bizpower</title>
<meta charset="windows-1255">
</head>
<body>
<div id="div_save" name="div_save" style="position:absolute; left:0px; top:0px; width:100%; height:100%; ">  												
  <table height="100%" width="100%" cellspacing="2" cellpadding="2" border=0 ID="Table1">  
     <tr><td align="center" valign=middle>
         <table border="0" height="100" width="400" cellspacing="1" cellpadding="1" ID="Table2">
            <tr>  
              <td align="center" bgcolor="#d0d0d0" valign=middle>
              <h4><b>מתבצעת טעינת הדוח</b></h4>              
              <h4>... אנא המתן</h4>
              </td>
            </tr>
         </table>
         </td>
     </tr>
  </table>
</div>
<% lang_id = cInt(trim(Request.Cookies("bizpegasus")("LANGID")))
   OrgID = cInt(trim(Request.Cookies("bizpegasus")("OrgID")))
   org_name = cStr(trim(Request.Cookies("bizpegasus")("ORGNAME")))
   
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
  
   set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
   filestring="../../../download/reports/"
   'Response.Write server.mappath(filestring)
   'Response.End
   fs.DeleteFile server.mappath(filestring) & "/*.*" ,False  
   
   filestring="../../../download/reports/export_SMS_" & trim(OrgID) & ".xls"
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
	XLSfile.WriteLine "	margin-right:5pt;"
	XLSfile.WriteLine "	margin-left:5pt;"
	XLSfile.WriteLine "	font-weight:700;"
	XLSfile.WriteLine "	font-family:Arial, sans-serif;"
	XLSfile.WriteLine "	text-align:center;"
	XLSfile.WriteLine "	background:#FF9900;"
	XLSfile.WriteLine "	border: 0.5pt solid black;"
	XLSfile.WriteLine "	mso-pattern:auto none;"
	XLSfile.WriteLine "	mso-height-source: userset;"
	XLSfile.WriteLine "	height:16pt;"
	XLSfile.WriteLine " white-space:normal;}"
	XLSfile.WriteLine ".xl26"  ' data cells
	XLSfile.WriteLine "	{mso-style-parent:style0;"
	XLSfile.WriteLine "	font-weight:400;"
	XLSfile.WriteLine "	font-size:10.0pt;"
	XLSfile.WriteLine "	margin-right:5pt;"
	XLSfile.WriteLine "	margin-left:5pt;"
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
	XLSfile.WriteLine "<x:Name>דוח שליחת SMS</x:Name>"
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
	
    strLine="<tr><td colspan=8 class=xl24>דוח שליחת SMS</td></tr>"
 '   strLine=strLine&"<tr><td colspan=8 class=xl24></td></tr>"    
 if start_date<>"" then   
    strLine=strLine&"<tr><td colspan=8 class=xl24>"&start_date &"-"& end_date &"</td></tr>"
    strLine=strLine&"<tr><td colspan=8 class=xl24></td></tr>"    
 end if
   
   		strLine =strLine& "<tr><td class=xl26 x:str></td>" & _
		"<td class=xl26 x:str></td><td class=xl26 x:str></td><td class=xl26 x:str></td>" &_
		"<td class=xl26 x:str></td><td class=xl26 x:str></td>"&_
		"<td class=xl26 x:str></td></tr>"

    
    arrTitles = Array("תאריך", "שם חברה", "איש קשר", "טלפון נייד" , "שדה כותרת", "מעת", "תוכן","סטטוס")
    
    If IsArray(arrTitles) Then
    strLine = strLine & "<tr>"
    For aa=0 To Ubound(arrTitles)
		strLine = strLine & "<td class=xl25> " & arrTitles(aa) & "</td>"
    Next
    strLine = strLine & "</tr>"
    End If   
   
    XLSfile.writeline strLine 

    	sql = "Exec dbo.get_SMS_report  @start_date='" & start_date & "', @end_date='" & end_date &"'"

 '  Response.Write sql
   ' Response.End
    set listContact=con.GetRecordSet(SQL)
    If not listContact.EOF Then   
	  contArray = listContact.GetRows()	  	  
	  recCount =  listContact.RecordCount 		
	  listContact.close
      set listContact=Nothing
      count = 0
      do while count < recCount
		'name = trim(contArray(0,count))
		date_send = trim(contArray(1,count))
		companyName= trim(contArray(2,count))
		contactName= trim(contArray(3,count))
		cellurar = trim(contArray(4,count)) 
		statusTypeName=trim(contArray(5,count)) 
		userName=trim(contArray(6,count)) 
		statusName=trim(contArray(7,count)) 
		smsContent=trim(contArray(8,count)) 
	'	If Len(trim(contArray(1,count))) > 3 Then
	'		phone = trim(contArray(1,count))
	'	Else
	'		phone = ""	
	'	End If
	'	If Len(trim(contArray(2,count))) > 3 Then
	'		cellurar = trim(contArray(2,count)) 
	'	Else
	'		cellurar = ""	
	'	End If     	
			
		strLine = "<tr><td class=xl26 x:str>" & date_send & "</td>" & _
		"<td class=xl26 x:str>" & companyName & "</td><td class=xl26 x:str>" & contactName & "</td><td class=xl26 x:str>" & cellurar & "</td>" &_
		"<td class=xl26 x:str>" & statusTypeName & "</td><td class=xl26 x:str>" & userName &  "</td>"&_
		"<td class=xl26 x:str>" & smsContent & "</td><td class=xl26 x:str>" & statusName & "</td></tr>"
			    
		XLSfile.writeline strLine
	      
		count=count+1
      loop
    End If
    	
	XLSfile.WriteLine "</TABLE></body></html>"    
	XLSfile.close
	set con=Nothing
    set fs = Nothing
	Response.Redirect(filestring) 'Open the excel file in the browser%>
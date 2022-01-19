
<% Response.CharSet = "windows-1255"
       Response.Buffer = True 
       Server.ScriptTimeout = 1000  %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
<head></head>
<body>
<div id="div_save" name="div_save" style="position:absolute; left:0px; top:0px; width:100%; height:100%;">  												
  <table height="100%" width="100%" cellspacing="2" cellpadding="2" border="0" ID="Table1">
     <tr><td align="center" valign=middle>
     <table dir=ltr border="0" height="100" width="400" cellspacing="1" cellpadding="1" ID="Table2">
     <tr><td align="center" valign=middle><font size="5" color=#0066ff face="Arial">אנא המתן בעת טעינת הדוח</font></td></tr>
      </table>
    </td></tr></table>
</div>
</body>
</html>
<!--#include file="../reports/code.asp" -->
<%  Session.LCID = 1037
	

	start_date = DateValue(trim(Request("dateStart")))
    end_date = DateValue(trim(Request("dateEnd")))
    	dir_var = "rtl" :  align_var = "center" : dir_obj_var = "ltr"
   Dim strXLS
	Set strXLS = New StrConCatArray
  
    strXLS.Add  "<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
	strXLS.Add  "<head>"
	strXLS.Add  "<meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1255"">"
	strXLS.Add  "<meta name=ProgId content=Excel.Sheet>"
	strXLS.Add  "<meta name=Generator content=""Microsoft Excel 9"">"
	strXLS.Add  "<style>"
	strXLS.Add  "<!--table"
	strXLS.Add  "	{mso-displayed-decimal-separator:""\."";"
	strXLS.Add  "	mso-displayed-thousand-separator:""\,"";}"
	strXLS.Add  "@page"
	strXLS.Add  "{ margin:1.0in .75in 1.0in .75in;"
	strXLS.Add  "	mso-header-margin:.5in;"
	strXLS.Add  "	mso-footer-margin:.5in;"
	strXLS.Add  " mso-page-orientation:landscape;}"
	strXLS.Add  "tr"
	strXLS.Add  "	{mso-height-source:auto;}"
	strXLS.Add  "br"
	strXLS.Add  "	{mso-data-placement:same-cell;}"
	strXLS.Add  ".style0"
	strXLS.Add  "	{mso-number-format:General;"
	strXLS.Add  "	vertical-align:bottom;"
	strXLS.Add  "	white-space:nowrap;"
	strXLS.Add  "	mso-rotate:0;"
	strXLS.Add  "	mso-background-source:auto;"
	strXLS.Add  "	mso-pattern:auto;"
	strXLS.Add  "	color:windowtext;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	font-weight:400;"
	strXLS.Add  "	font-style:normal;"
	strXLS.Add  "	text-decoration:none;"
	strXLS.Add  "	font-family:Arial;"
	strXLS.Add  "	mso-generic-font-family:auto;"
	strXLS.Add  "	mso-font-charset:0;"
	strXLS.Add  "	border:none;"
	strXLS.Add  "	mso-protection:locked visible;"
	strXLS.Add  "	text-align:center;"
	strXLS.Add  "	mso-style-name:Normal;"
	strXLS.Add  "	mso-height-source:auto;"
	strXLS.Add  "	height:20pt;}"
	strXLS.Add  ".xl24"  ' title
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	font-size:12.0pt;"
	strXLS.Add  "	font-weight:700;"
	strXLS.Add  "	text-align:center;"
	strXLS.Add  "	direction: " & dir_obj_var & ";"
	strXLS.Add  "	white-space:normal;}"
	strXLS.Add  ".xl25" ' title cell
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	color:#000000;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	font-weight:700;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	text-align:center;"
	strXLS.Add  "	background:#FF9900;"
	strXLS.Add  "   width:150px;"	
	strXLS.Add  "	border: 0.5pt solid black;"
	strXLS.Add  "	mso-pattern:auto none;"
	strXLS.Add  "	mso-height-source: auto;"	
	strXLS.Add  " white-space:normal;}"
	
	strXLS.Add  ".xl25_blue" ' title cell
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	color:#000000;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	font-weight:700;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	text-align:center;"
	strXLS.Add  "	background:#D4E3F2;"
	strXLS.Add  "   width:150px;"	
	strXLS.Add  "	border: 0.5pt solid black;"
	strXLS.Add  "	mso-pattern:auto none;"
	strXLS.Add  "	mso-height-source: auto;"	
	strXLS.Add  " white-space:normal;}"
	
	strXLS.Add  ".xl25_white" ' title cell
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	color:#000000;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	font-weight:700;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	text-align:center;"
	strXLS.Add  "	background:#ffffff;"
	strXLS.Add  "   width:150px;"	
	strXLS.Add  "	border: 0.5pt solid black;"
	strXLS.Add  "	mso-pattern:auto none;"
	strXLS.Add  "	mso-height-source: auto;"	
	strXLS.Add  " white-space:normal;}"
	
	
		strXLS.Add  ".xl25_red" ' title cell
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	color:#C00000;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	font-weight:700;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	text-align:center;"
	strXLS.Add  "	background:#ffffff;"
	strXLS.Add  "   width:150px;"	
	strXLS.Add  "	border: 0.5pt solid black;"
	strXLS.Add  "	mso-pattern:auto none;"
	strXLS.Add  "	mso-height-source: auto;"	
	strXLS.Add  " white-space:normal;}"
	
	
	strXLS.Add  ".xl25_green" ' title cell
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	color:#00B050;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	font-weight:700;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	text-align:center;"
	strXLS.Add  "	background:#ffffff;"
	strXLS.Add  "   width:150px;"	
	strXLS.Add  "	border: 0.5pt solid black;"
	strXLS.Add  "	mso-pattern:auto none;"
	strXLS.Add  "	mso-height-source: auto;"	
	strXLS.Add  " white-space:normal;}"
	
	strXLS.Add  ".xl25_blueRight" ' title cell
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	color:#000000;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	font-weight:700;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	text-align:right;"
	strXLS.Add  "	background:#D4E3F2;"
	strXLS.Add  "   width:150px;"	
	strXLS.Add  "	border: 0.5pt solid black;"
	strXLS.Add  "	mso-pattern:auto none;"
	strXLS.Add  "	mso-height-source: auto;"	
	strXLS.Add  " white-space:normal;}"
	
		
	strXLS.Add  ".xl26"  ' data cells
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	font-weight:400;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	border: 0.5pt solid black;"
	strXLS.Add  "	mso-height-source: auto;"	
	strXLS.Add  "	vertical-align:top;"
	strXLS.Add  "	text-align:"&align_var&";}"
	strXLS.Add  ".xl26_yellow"  ' data cells
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	font-weight:400;"
	strXLS.Add  "	font-size:10.0pt;"
	strXLS.Add  "	margin-right:5pt;"
	strXLS.Add  "	margin-left:5pt;"
	strXLS.Add  "	background:yellow;"
	strXLS.Add  "	font-family:Arial, sans-serif;"
	strXLS.Add  "	border: 0.5pt solid black;"
	strXLS.Add  "	mso-height-source: auto;"	
	strXLS.Add  "	vertical-align:top;"
	strXLS.Add  "	text-align:"&align_var&";}"
	strXLS.Add  "-->"
	strXLS.Add  "</style>"
	strXLS.Add  "<!--[if gte mso 9]><xml>"
	strXLS.Add  "<x:ExcelWorkbook>"
	strXLS.Add  "<x:ExcelWorksheets>"
	strXLS.Add  "<x:ExcelWorksheet>"
	strXLS.Add  "<x:Name>דוח מקורות הגעה</x:Name>"
	strXLS.Add  "<x:WorksheetOptions>"
	strXLS.Add  "<x:Print>"
	strXLS.Add  "      <x:ValidPrinterInfo/>"
	strXLS.Add  "      <x:PaperSizeIndex>9</x:PaperSizeIndex>"
	strXLS.Add  "      <x:HorizontalResolution>600</x:HorizontalResolution>"
	strXLS.Add  "      <x:VerticalResolution>600</x:VerticalResolution>"
	strXLS.Add  "</x:Print>"
	strXLS.Add  "<x:Selected/>"
	strXLS.Add  "<x:ProtectContents>False</x:ProtectContents>"
	strXLS.Add  "<x:ProtectObjects>False</x:ProtectObjects>"
	strXLS.Add  "<x:ProtectScenarios>False</x:ProtectScenarios>"
	strXLS.Add  "<x:WorksheetOptions>"
	strXLS.Add  "<x:Selected/>"
	strXLS.Add  "<x:DisplayRightToLeft/>"
	strXLS.Add  "<x:Panes>"
	strXLS.Add  "<x:Pane>"
	strXLS.Add  "<x:Number>1</x:Number>"
	strXLS.Add  "<x:ActiveRow>1</x:ActiveRow>"
	strXLS.Add  "</x:Pane>"
	strXLS.Add  "</x:Panes>"
	strXLS.Add  "<x:ProtectContents>False</x:ProtectContents>"
	strXLS.Add  "<x:ProtectObjects>False</x:ProtectObjects>"
	strXLS.Add  "<x:ProtectScenarios>False</x:ProtectScenarios>"
	strXLS.Add  "</x:WorksheetOptions>"
	strXLS.Add  "</x:ExcelWorksheet>"
	strXLS.Add  "</x:ExcelWorksheets>"
	strXLS.Add  "<x:WindowTopX>360</x:WindowTopX>"
	strXLS.Add  "<x:WindowTopY>60</x:WindowTopY>"
	strXLS.Add  "</x:ExcelWorkbook>"
	strXLS.Add  "</xml><![endif]-->"
	strXLS.Add  "</head>"
	strXLS.Add  "<body>"
	strXLS.Add  "<table style='border-collapse:collapse;table-layout:fixed' dir=" : strXLS.Add dir_var  : strXLS.Add ">"  
    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=9 class=xl24 style='background:#808080;color:#ffffff'>דוח מקורות הגעה</td></tr>"

    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=9 class=xl24 style='background:#808080;color:#ffffff'> " & start_date & " - " & end_date & "</td></tr>"
  	strXLS.Add  "<tr>"
		strXLS.Add  "<td class=xl25_blue   x:str>איש קשר</td>"
		strXLS.Add  "<td class=xl25_blue  x:str style='background:#C5D9F1'>טלפון סלולרי</td>"
		strXLS.Add  "<td class=xl25_blue x:str style='background:#C5D9F1'>מקור ההגעה</td>"
		strXLS.Add  "<td class=xl25_blue x:str style='background:#C5D9F1'>קמפיין ההגעה</td>"
		strXLS.Add  "<td class=xl25_blue  x:str style='background:#C5D9F1'>מדיום ההגעה</td>"
	strXLS.Add  "<td class=xl25_blue  x:str style='background:#C5D9F1'>תאריך פתיחת טופס צור קשר מקוצר או צור קשר רגיל</td>"
		strXLS.Add  "<td class=xl25_blue  x:str style='background:#C5D9F1'>תאריך פתיחת טופס מתעניין בטיול</td>"
		strXLS.Add  "<td class=xl25_blue  x:str style='background:#C5D9F1'>תאריך שינוי טופס מתעניין בטיול בסטטוס ""סגור נרשם""</td>"
		strXLS.Add  "<td class=xl25_blue x:str style='background:#C5D9F1'>סטטוס טופס ""המתעניין בטיול"" הנוכחי</td>"

	strXLS.Add "</tr>"

  	sql = "Exec dbo.get_LeadReport  @date_start='" & start_date & "', @date_end='" & end_date &"'"
 
	set rsa = con.getRecordSet(sql)
	If Not rsa.Eof Then
		arr_app = rsa.getRows()
	End If
	Set rsa = Nothing	
	If isArray(arr_app) Then
	For aa=0 To Ubound(arr_app,2)
		pCONTACT_NAME=trim(arr_app(0,aa))
		pcellular=trim(arr_app(2,aa))
		pAppDate=trim(arr_app(3,aa))
		pAPPID=trim(arr_app(4,aa))
		pQUESTIONS_ID=trim(arr_app(5,aa))
		pCONTACT_ID=trim(arr_app(6,aa))
		
		IF pQUESTIONS_ID=16724 THEN
		sqlSelect="SELECT Field_Id,FIELD_VALUE from FORM_VALUE WHERE  (APPEAL_ID = "& pAPPID &") and (dbo.FORM_VALUE.Field_Id=40820 OR dbo.FORM_VALUE.Field_Id=40821 OR dbo.FORM_VALUE.Field_Id=40822)"
		END IF
		IF pQUESTIONS_ID=17012 then
		sqlSelect="SELECT Field_Id,FIELD_VALUE from FORM_VALUE WHERE  (APPEAL_ID = "& pAPPID &") and (dbo.FORM_VALUE.Field_Id=40812 OR dbo.FORM_VALUE.Field_Id=40813 OR dbo.FORM_VALUE.Field_Id=40823 or  dbo.FORM_VALUE.Field_Id=40776 )"
	
		END IF	
	'	response.Write 	sqlSelect &"<BR>"
		'Response.end
		set LeadList = con.getRecordSet(sqlSelect)
		Do while (Not LeadList.eof)
	  	
		FID=LeadList(0)
		FValue=rtrim(ltrim(LeadList(1)))
		select  case FID
		case "40820" 
		pMakor=FValue
		case "40812" 
		pMakor=FValue
		case "40776"
		pMakor=FValue
		case "40821" 
		pCampaign=FValue
		case "40813"
		pCampaign=FValue
		
		case "40822" 
		pMedium=FValue
		case "40823"
		pMedium=FValue
		end select
	LeadList.movenext
	
		loop
	
	set LeadList=Nothing
	Status16504=""
	Date16504Close=""
	pDate16504=""
	 	sqlForm = "Exec dbo.get_LeadReport_16504  @CONTACT_ID='" & pCONTACT_ID & "', @AppDate='" & pAppDate &"'"
 'response.Write sqlForm &"<BR>"
' response.end
	set rsaF = con.getRecordSet(sqlForm)
	If Not rsaF.Eof Then
		arr_appF = rsaF.getRows()
	
	Set rsaF = Nothing	
	If isArray(arr_appF) Then
	For aaF=0 To Ubound(arr_appF,2)
		pDate16504=trim(arr_appF(0,aaF))
		if isDate(pDate16504) then
			pDate16504=day(pDate16504)&"/"& month(pDate16504) &"/" & year(pDate16504)
	    end if
		Status16504=trim(arr_appF(1,aaF))
		Date16504Close=trim(arr_appF(2,aaF))
		if isDate(Date16504Close) then
			Date16504Close=day(Date16504Close)&"/"& month(Date16504Close) &"/" & year(Date16504Close)
	   else
	   Date16504Close=""
	    end if
	    if Status16504=5 then
		else
		Date16504Close=""
		end if
		select case Status16504
		case 1
		Status16504="לא נקרא"
		case 2
		Status16504="נקרא"
		case 3
		Status16504="סגור"
		case 4
		Status16504="במעקב"
		case 5
		Status16504="סגור נרשם"
		end select 
	Next	
	end if
	else

	End If
		'pMakor=trim(arr_app(6,aa))
		'pCampaign=trim(arr_app(7,aa))
		'pMedium=trim(arr_app(8,aa))
		'pDate16504=trim(arr_app(9,aa))
		'if isDate(pDate16504) then
		'	pDate16504=day(pDate16504)&"/"& month(pDate16504) &"/" & year(pDate16504)
	    'end if
		'Status16504=trim(arr_app(10,aa))
		'Date16504Close=trim(arr_app(11,aa))
		'if isDate(Date16504Close) then
		'	Date16504Close=day(Date16504Close)&"/"& month(Date16504Close) &"/" & year(Date16504Close)
	    'end if
		'if Status16504=5 then
		'else
		'Date16504Close=""
		'end if
		'select case Status16504
		'case 1
		'Status16504="לא נקרא"
		'case 2
		'Status16504="נקרא"
		'case 3
		'Status16504="סגור"
		'case 4
		'Status16504="במעקב"
		'case 5
		'Status16504="סגור נרשם"
		'end select 
		
	 	strXLS.Add  "<tr>"
		strXLS.Add  "<td class=xl26   x:str>":strXLS.Add  pCONTACT_NAME : strXLS.Add  "</td>"
		strXLS.Add  "<td class=xl26  x:str style='background:#C5D9F1'>" :strXLS.Add pcellular : strXLS.Add  "</td>"
		strXLS.Add  "<td class=xl26  x:str style='background:#C5D9F1'>" :strXLS.Add pMakor : strXLS.Add  "</td>"
		strXLS.Add  "<td class=xl26  x:str style='background:#C5D9F1'>" :strXLS.Add pCampaign : strXLS.Add  "</td>"
		strXLS.Add  "<td class=xl26  x:str style='background:#C5D9F1'>" :strXLS.Add pMedium : strXLS.Add  "</td>"
		strXLS.Add  "<td class=xl26  x:str style='background:#C5D9F1'>" :strXLS.Add pAppDate : strXLS.Add  "</td>"
		strXLS.Add  "<td class=xl26  x:str style='background:#C5D9F1'>" :strXLS.Add pDate16504 : strXLS.Add  "</td>"
		strXLS.Add  "<td class=xl26  x:str style='background:#C5D9F1'>" :strXLS.Add Date16504Close : strXLS.Add  "</td>"
	
		strXLS.Add  "<td class=xl26  x:str style='background:#C5D9F1'>" :strXLS.Add Status16504 : strXLS.Add  "</td>"
	
	strXLS.Add "</tr>"
	'response.Write strXLS.value &"<BR>"
	

	next
end if
'response.end


   
       strXLS.Add  "</TABLE></body></html>"    
	Set con = Nothing	 
	
	filestring = "LeadReport.xls"
    Response.Clear()
    Response.AddHeader "content-disposition", "inline; filename=" & filestring
    Response.ContentType = "application/vnd.ms-excel"
    Response.Write(strXLS.value)  
    Set strXLS = Nothing %>

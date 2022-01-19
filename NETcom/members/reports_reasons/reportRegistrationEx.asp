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
<!--#include file="code.asp" -->
<%  Session.LCID = 1037
	quest_id = 16735
	countryid=trim(Request("country_id"))

	start_date = DateValue(trim(Request("dateStart")))
    end_date = DateValue(trim(Request("dateEnd")))
    
    
reasons=""
	'only for tofes mitan'en
	
			'הגורם ליצירת הטופס
			'---------------P2932 - type of form- REASONS ---------------
				'select creation reason-------------------------------
				'added by Mila 22/10/2019-----------------------------	
			reasons=Request.Form("reasons")
			reasons=replace(reasons," ","")
	'-----------------------------------------------------------
	
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
	strXLS.Add  "<x:Name>דוח מעקב רישום</x:Name>"
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
    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=10 class=xl24 style='background:#808080;color:#ffffff'>דוח מעקב רישום</td></tr>"

    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=10 class=xl24 style='background:#808080;color:#ffffff'> " & start_date & " - " & end_date & " טפסי רישום שנפתחו בין תאריכים </td></tr>"
   
sqlQuery=" and appeal_CountryId   in ("& countryID &")"
'	response.Write "countryid-="&countryid   
    	'---------------P2932 - type of form- REASONS ---------------
   	If reasons<>"" Then
		sqlstr = "SELECT  Reason_Title FROM  Appeals_CreationReasons where QUESTIONS_ID=16504 and Reason_Id in (" & reasons & ") order by Reason_Order"
			set rs_Reason = con.getRecordSet(sqlstr)
			If Not rs_Reason.eof Then
				arr_Reason = rs_Reason.getRows()
			End If
			Set rs_Reason = Nothing
			If isArray(arr_Reason) Then
			
    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=10 class=xl24 style='background:#808080;color:#ffffff'> הגורם ליצירת הטופס:"
 				'response.Write(Ubound(arr_Reason,2))
				'response.end
				For mm=0 To Ubound(arr_Reason,2)
					if mm>0 then
						 strXLS.Add     "<font color=""#FF9900"">,&nbsp;</font>"
					end if
					 strXLS.Add   "<font color=""#FF9900"">" & trim(arr_Reason(0,mm)) & "</font>"
				next
				   strXLS.Add  "</td></tr>"

			end if
	end if
 
 sqlQuery=" and appeal_CountryId   in ("& countryID &")"
	
		'---------------P2932 - type of form- REASONS ---------------

		'---------------P2932 - type of form- REASONS ---------------
	If reasons<>"" Then 	
		sql = "Exec dbo.get_reportRegistration_Reasons  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
			" @sqlQuery='" & sqlQuery & "', @productID=" & nFix(quest_id) & ", @reasons=' " & reasons &  "'"
	else
	sql = "Exec dbo.get_reportRegistration_Reasons  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
			" @sqlQuery='" & sqlQuery & "', @productID=" & nFix(quest_id) 
		
	end if 
   
   '	sql = "Exec dbo.get_reportRegistration  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	'" @sqlQuery='" & sqlQuery & "', @productID=" & nFix(quest_id) 
	'Response.Write sql
	'Response.End
	set rsa = con.getRecordSet(sql)
	If Not rsa.Eof Then
		arr_app = rsa.getRows()
	End If
	Set rsa = Nothing	
	If isArray(arr_app) Then
	For aa=0 To Ubound(arr_app,2)
		pcountryid=trim(arr_app(0,aa))
		countryName=trim(arr_app(1,aa))
		count02=trim(arr_app(2,aa))
		count3_7=trim(arr_app(3,aa))
		count8_14=trim(arr_app(4,aa))
		count15_21=trim(arr_app(5,aa))
		count22_30=trim(arr_app(6,aa))
		count31_60=trim(arr_app(7,aa))
		count61_90=trim(arr_app(8,aa))
		count91_180=trim(arr_app(9,aa))
		'''''
		    pquest_id=16504
  		'---------------P2932 - type of form- REASONS ---------------
	If reasons<>"" Then 	
   sql = "Exec dbo.get_AppealsCount16504_Reasons  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @countryID=" & nFix(pcountryid) & ", @productID=" & nFix(pquest_id)  & ", @reasons=" & nFix(reasons) & ""
else
   sql = "Exec dbo.get_AppealsCount16504_Reasons  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @countryID=" & nFix(pcountryid) & ", @productID=" & nFix(pquest_id)
end if
'response.Write sql
'response.end
	set rs_count = con.getRecordSet(sql)
	if not rs_count.eof then
		appealsCount16504 = cLng(rs_count(0))
		else
		appealsCount16504=""
	end if
	set rs_count = nothing
	
 	 sql = "Exec dbo.get_AppealsCount16504_NotExists_Reasons  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @countryID=" & nFix(pcountryid) & ", @productID=" & nFix(pquest_id) & ", @StatusId=3 , @reasons=" & nFix(reasons) & ""'status סגור



	set rs_count = con.getRecordSet(sql)
	if not rs_count.eof then
		appealsCountNotExist_Close16504 = cLng(rs_count(0))
		else
		appealsCountNotExist_Close16504=""
	end if
	set rs_count = nothing

 sql = "Exec dbo.get_AppealsCount16504_NotExists_Reasons  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @countryID=" & nFix(pcountryid) & ", @productID=" & nFix(pquest_id) & ", @StatusId=4, @reasons=" & nFix(reasons) & "" 'status במעקב
	set rs_count = con.getRecordSet(sql)
	if not rs_count.eof then
		appealsCountNotExist_Track16504 = cLng(rs_count(0))
		else
		appealsCountNotExist_Track16504=""
	end if
	set rs_count = nothing

		''''
		
		sumCountry=	Cint(count02) +Cint(count3_7) +Cint(count8_14)+Cint(count15_21)+Cint(count22_30)+Cint(count31_60)+Cint(count61_90)+Cint(count91_180)
		pcount02=Round(100*Cint(count02)/Cint(sumCountry))
		pcount3_7=Round(100*Cint(count3_7)/Cint(sumCountry))
		pcount8_14=Round(100*Cint(count8_14)/Cint(sumCountry))
		pcount15_21=Round(100*Cint(count15_21)/Cint(sumCountry))
		pcount22_30=Round(100*Cint(count22_30)/Cint(sumCountry))
		pcount31_60=Round(100*Cint(count31_60)/Cint(sumCountry))
		pcount61_90=Round(100*Cint(count61_90)/Cint(sumCountry))
		pcount91_180=Round(100*Cint(count91_180)/Cint(sumCountry))
		
	    strXLS.Add	"<tr><td class=xl24 colspan=10>ימים מפתיחת טופס מתעניין עד ההרשמה </td></tr>"
  
   strXLS.Add	"<tr> <td class=xl25></td>"
   strXLS.Add	"<td class=xl25>0-2 ימים</td>"
   strXLS.Add	"<td class=xl25>3-7 ימים</td>"
   strXLS.Add	"<td class=xl25>8-14 ימים</td>"
   strXLS.Add	"<td class=xl25>15-21 ימים</td>"
   strXLS.Add	"<td class=xl25>22-30 ימים</td>"
   strXLS.Add	"<td class=xl25>31-60 ימים</td>" 
   strXLS.Add	"<td class=xl25>61-90 ימים</td>"
   strXLS.Add	"<td class=xl25>91-180 ימים</td>"
   strXLS.Add	"<td class=xl25>סה``כ נרשמו</td></tr>"


		strXLS.Add  "<tr><td class=xl24 x:str>"  : strXLS.Add countryName :strXLS.Add "</td>"
		strXLS.Add  "<td class=xl26 x:str>" : strXLS.Add count02  : strXLS.Add "</td>"
		strXLS.Add  "<td class=xl26  x:str>" : strXLS.Add count3_7 : strXLS.Add"</td>"
		strXLS.Add  "<td class=xl26  x:str>" : strXLS.Add count8_14 : strXLS.Add"</td>"
		strXLS.Add  "<td class=xl26  x:str>" : strXLS.Add count15_21 : strXLS.Add"</td>"
		strXLS.Add  "<td class=xl26  x:str>" : strXLS.Add count22_30 : strXLS.Add"</td>"
		strXLS.Add  "<td class=xl26  x:str>" : strXLS.Add count31_60 : strXLS.Add"</td>"
		strXLS.Add  "<td class=xl26  x:str>" : strXLS.Add count61_90 : strXLS.Add"</td>"
		strXLS.Add  "<td class=xl26  x:str>" : strXLS.Add count91_180 : strXLS.Add"</td>"
		strXLS.Add  "<td class=xl26  x:str>" : strXLS.Add sumCountry : strXLS.Add"</td></tr>"
		

	
		strXLS.Add  "<tr><td class=xl26_yellow x:str>אחוז מסך הכל נרשמים</td>"
		strXLS.Add  "<td class=xl26_yellow x:str>" : strXLS.Add pcount02  : strXLS.Add "%</td>"
		strXLS.Add  "<td class=xl26_yellow  x:str>" : strXLS.Add pcount3_7 : strXLS.Add"%</td>"
		strXLS.Add  "<td class=xl26_yellow  x:str>" : strXLS.Add pcount8_14 : strXLS.Add"%</td>"
		strXLS.Add  "<td class=xl26_yellow  x:str>" : strXLS.Add pcount15_21 : strXLS.Add"%</td>"
		strXLS.Add  "<td class=xl26_yellow  x:str>" : strXLS.Add pcount22_30 : strXLS.Add"%</td>"
		strXLS.Add  "<td class=xl26_yellow  x:str>" : strXLS.Add pcount31_60 : strXLS.Add"%</td>"
		strXLS.Add  "<td class=xl26_yellow  x:str>" : strXLS.Add pcount61_90 : strXLS.Add"%</td>"
 		strXLS.Add  "<td class=xl26_yellow  x:str>" : strXLS.Add pcount91_180 : strXLS.Add"%</td>"
 
    	strXLS.Add  "<td class=xl26_yellow  x:str></td></tr>"
		

		strXLS.Add  "<tr><td class=xl25_blueRight colspan=5 x:str>נפתחו ל" : strXLS.Add countryName : strXLS.Add"</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>טפסי מתעניין</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>נרשמו</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>% סגירה מכירה ל" : strXLS.Add countryName : strXLS.Add" הוא</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>כמות טפסי מתעניין ללא טופס רישום בסטטוס סגור</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>כמות טפסי מתעניין ללא טופס רישום בסטטוס מעקב</td></tr>"
	
	if appealsCount16504>0 then
		pSum= Round(cint(sumCountry)*100/cint(appealsCount16504)) 
		else
		pSum=0
		end if
		appealsCount16504All=appealsCount16504All+appealsCount16504
		sumCountryAll=sumCountryAll+sumCountry
	
		strXLS.Add  "<TR><td class=xl25_blue colspan=5 x:str></td>"
        strXLS.Add  "<td class=xl25_blue  x:str>"  : strXLS.Add appealsCount16504  : strXLS.Add "</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>"  : strXLS.Add sumCountry  : strXLS.Add "</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>" : strXLS.Add pSum : strXLS.Add "%</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>" : strXLS.Add appealsCountNotExist_Close16504: strXLS.Add "</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>" : strXLS.Add appealsCountNotExist_Track16504: strXLS.Add "</td></tr>"
		
			
		
		
		
		'strXLS.Add  "<tr><td style='background:#92D050 ;height:6px' colspan=10></td></tr>"
		strXLS.Add  "<tr><td class=xl26 colspan=10></td></tr>"
		strXLS.Add  "<tr><td class=xl26 colspan=10></td></tr>"
		
		
	
	
	next
	if appealsCount16504All>0 then
	pRound=round(cint(sumCountryAll)*100/cint(appealsCount16504All))
	
	else
	pRound=0
	end if
	strXLS.Add  "<tr><td class=xl25_blueRight colspan=7></td><td class=xl25_blueRight>טפסי מתעניין</td><td class=xl25_blueRight>נרשמו</td><td class=xl25_blueRight>אחוז סגירת מחירה ליעדים</td></tr>"
	strXLS.Add  "<tr><td class=xl25_blueRight colspan=7>בין התאריכים" &start_date & " - " & end_date &" <BR> נפתחו ל "& Ubound(arr_app,2)+1 &" היעדים</td><td class=xl25_blue>"& appealsCount16504All &"</td><td class=xl25_blue>"&sumCountryAll&"</td><td class=xl25_blue>"& pRound &"%</td></tr>"

	end if
	Set con = Nothing	 

   
   
   
       strXLS.Add  "</TABLE></body></html>"    
	Set con = Nothing	 
	
	filestring = "RegistrationForm.xls"
    Response.Clear()
    Response.AddHeader "content-disposition", "inline; filename=" & filestring
    Response.ContentType = "application/vnd.ms-excel"
    Response.Write(strXLS.value)  
    Set strXLS = Nothing %>

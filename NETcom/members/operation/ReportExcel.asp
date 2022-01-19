
<% Response.CharSet = "windows-1255"
       Response.Buffer = True 
       Server.ScriptTimeout = 1000  %>
       <%type_reports=request.QueryString("type_reports")
       select case type_reports
       case "1"
       reportTitle="טיולים שעומדים לצאת בשבועיים הקרובים וטרם נקבע להם בריף"
       case "2"
       reportTitle="טיולים שנמצאים כרגע בחו``ל ולא בוצעו שיחות למדריך"
       case "3"
       reportTitle="טיולים שחזרו וטרם תואם מפגש חזרה עם מדריך"
       case "4"
       reportTitle="טיולים שעומדים לצאת בשבועיים הקרובים וחסרה הזמנה מסימולטני"
       case "5"
       reportTitle="בריף שמתקיים ביומיים הקרובים ולא הופקו שוברים"
       end select%>
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
	
		strXLS.Add  ".xl25_red" ' title cell
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	color:#C00000;"
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
	
	
	strXLS.Add  ".xl25_green" ' title cell
	strXLS.Add  "	{mso-style-parent:style0;"
	strXLS.Add  "	color:#00B050;"
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
	strXLS.Add  "	{mso-number-format:\@;"
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
	strXLS.Add  "<x:Name>דוח אופרציה</x:Name>"
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
    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=9 align=right class=xl24 style='background:#808080;color:#ffffff'>" : strXLS.Add reportTitle : strXLS.Add"</td><td colspan=20></td></tr>"
   strXLS.Add  "<tr style='background:#C5D9F1;color:#000000'><td colspan=9 align=right class=xl24 style='background:#C5D9F1;color:#000000'>תאריך הפקה דוח " : strXLS.Add Now() : strXLS.Add"</td><td colspan=20></td></tr>"

  strXLS.Add  "<tr>"
	strXLS.Add  "<td class=xl25_blue>אופרטור</td>"
	strXLS.Add  "<td class=xl25_blue>טיול </td>"
	strXLS.Add  "<td class=xl25_blue>תאריך</td>"
	strXLS.Add  "<td class=xl25_blue>קוד טיול</td>"
	strXLS.Add  "<td class=xl25_blue>תאריך יציאה </td>"	
	strXLS.Add  "<td class=xl25_blue>תאריך חזרה </td>"
	strXLS.Add  "<td class=xl25_blue>?כרגע בחול</td>"
	strXLS.Add  "<td class=xl25_blue>שם מדריך </td>"
	strXLS.Add  "<td class=xl25_blue>טלפון של מדריך</td>"
	strXLS.Add  "<td class=xl25_blue>ספק </td>"
	strXLS.Add  "<td class=xl25_blue>תאריך קבלת Itinerary </td>"
	strXLS.Add  "<td class=xl25_blue>תאריך בריף </td>"
	strXLS.Add  "<td class=xl25_blue>שעת בריף</td>"	
	strXLS.Add  "<td class=xl25_blue>תאריך מפגש קבוצה </td>"
	strXLS.Add  "<td class=xl25_blue>שעת מפגש קבוצה</td>"
	strXLS.Add  "<td class=xl25_blue>הוקלדו מלונות בגלבוע?</td>"
	strXLS.Add  "<td class=xl25_blue>שובר לסימולטני?</td>"
	strXLS.Add  "<td class=xl25_blue>שוברי הוצאות קבוצה ומדריך?</td>"
	strXLS.Add  "<td class=xl25_blue>תמחיר</td>"
	strXLS.Add  "<td class=xl25_blue>שוברי ספק קרקע?</td>"
	strXLS.Add  "<td class=xl25_blue>הערכות כספית -$</td>"
	strXLS.Add  "<td class=xl25_blue>הערכות כספית -€</td>"
	strXLS.Add  "<td class=xl25_blue>תאריך פגישה לאחר טיול </td>"
	strXLS.Add  "<td class=xl25_blue>שעת פגישה לאחר טיול</td>"
	strXLS.Add  "<td class=xl25_blue>נשלח מכתב למפגש קבוצה</td>"
	strXLS.Add  "<td class=xl25_blue>הועבר איטינררי למדריך</td>"
	strXLS.Add  "<td class=xl25_blue>הועברו כרטיסי טיסה</td>"
	strXLS.Add  "<td class=xl25_blue>טיפול הסתיים?</td>"
	strXLS.Add  "<td class=xl25_blue>כמות שיחות</td>"
	
strXLS.Add  "</tr>"
    	sql = "Exec dbo.GetWorkScreen_Excel  @tab=" & nFix(type_reports)
    	
	'Response.End
	set rsa = conPegasus.Execute(sql)
 If Not rsa.Eof Then
 Do while Not rsa.eof 
 strXLS.Add  "<tr>"
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("User_Name")  :strXLS.Add "</td>"
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Series_Name")  :strXLS.Add "</td>"
 mm = Month(rsa("Departure_Date"))
     dd = Day(rsa("Departure_Date"))
     
     IF len(mm) = 1 THEN
       mm = "0" & mm
     END IF
     IF len(dd) = 1 THEN
       dd = "0" & dd
     END IF
strXLS.Add"<td class=xl26>" :strXLS.Add mm :strXLS.Add dd :strXLS.Add "</td>"   '--{0:MMdd}
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Departure_Code")  :strXLS.Add "</td>"
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Departure_Date")   :strXLS.Add "</td>" 
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Departure_Date_End")  :strXLS.Add "</td>" 
strXLS.Add"<td class=xl26>" :strXLS.Add  TourStatusForOperation(rsa("Departure_Id"),dbNullFix(rsa("Departure_Date")),dbNullFix(rsa("Departure_Date_End"))) :strXLS.Add "</td>"   '----TourStatusForOperation(rsa("Departure_Id"),dbNullFix(rsa("Departure_Date")),dbNullFix(rsa("Departure_Date_End")))
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("GuideName")  :strXLS.Add "</td>"
strXLS.Add"<td class=xl26>" :strXLS.Add breaks(rsa("Departure_GuideTelphone")) :strXLS.Add "</td>" 
strXLS.Add"<td class=xl26>" :strXLS.Add GetSelectSupplierName(rsa("supplier_Id"))  :strXLS.Add "</td>"   '---IIf(func.sFix(DataBinder.Eval(Container.DataItem, "supplier_Id"))="","<img src=../../images/v.png width=20 style=display:none  id=""Imgsupplier_Id_"& Container.DataItem("Departure_Id")&""">", "<img src=../../images/v.png width=20 id=""Imgsupplier_Id_"& Container.DataItem("Departure_Id")&""" title="""& func.GetSelectSupplierName(func.sFix(Container.DataItem("supplier_Id")))&""">"
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Departure_Itinerary") :strXLS.Add "</td>"  ''---
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Departure_DateBrief") :strXLS.Add "</td>"  
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Departure_TimeBrief"):strXLS.Add "</td>"
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Departure_DateGroupMeeting"):strXLS.Add "</td>" 
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Departure_TimeGroupMeeting"):strXLS.Add "</td>"  
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("GilboaHotel") :strXLS.Add "</td>" 
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Voucher_Simultaneous"):strXLS.Add "</td>" 
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Voucher_Group"):strXLS.Add "</td>" 
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Departure_Costing")  :strXLS.Add "</td>" 
strXLS.Add"<td class=xl26>" :strXLS.Add GetVouchers_ProviderStatus(rsa("Departure_Id"))  :strXLS.Add "</td>"   
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Financial_Dolar"):strXLS.Add "</td>" 
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("Financial_Euro") :strXLS.Add "</td>"  
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("DateMeetingAfterTrip")  :strXLS.Add "</td>" 
strXLS.Add"<td class=xl26>" :strXLS.Add rsa("TimeMeetingAfterTrip"):strXLS.Add "</td>" 
if rsa("isSendLetterGroupMeeting") then
strXLS.Add"<td class=xl26>כן</td>"	
else
strXLS.Add"<td class=xl26></td>"	
end if
if rsa("isSendToGuide") then
strXLS.Add"<td class=xl26>כן</td>"	
else
strXLS.Add"<td class=xl26></td>"	
end if

if rsa("isSendTicket") then
strXLS.Add"<td class=xl26>כן</td>"
else
strXLS.Add"<td class=xl26></td>"
end if
strXLS.Add"<td class=xl26>" :strXLS.Add StatusOperation(rsa("Status")) :strXLS.Add "</td>" 
strXLS.Add"<td class=xl26>" :strXLS.Add GetCountGuideMessages(rsa("Departure_Id")) :strXLS.Add "</td>"   
strXLS.Add  "</tr>"
 rsa.movenext

 Loop
 end if


  
		

   
   
       strXLS.Add  "</TABLE></body></html>"    
	
	filestring = "OperationExcel.xls"
    Response.Clear()
    Response.AddHeader "content-disposition", "inline; filename=" & filestring
    Response.ContentType = "application/vnd.ms-excel"
    Response.Write(strXLS.value)  
    Set strXLS = Nothing %>

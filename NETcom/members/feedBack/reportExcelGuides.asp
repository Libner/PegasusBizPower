<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<% Response.CharSet = "windows-1255"
       Response.Buffer = True 
       Server.ScriptTimeout = 1000  %>
       <%guideId=Request("guide_id")
       currentYear=request("currentYear")
       
       FromDate ="1/01/"& currentYear
		 ToDate ="31/12/"& currentYear
	'	 response.Write FromDate &":"& ToDate
	'	 response.end
	sqlstrGuide = "SELECT Guide_Id, (Guide_FName + ' - ' + Guide_LName) as Guide_Name,Guide_Phone,Guide_Email  FROM Guides  where Guide_Id=" & guideId

	set rs_Guide= conPegasus.Execute(sqlstrGuide)
   if not rs_Guide.EOF then
		Guide_Id =rs_Guide("Guide_Id")
		Guide_Name =rs_Guide("Guide_Name")
		Guide_Phone =rs_Guide("Guide_Phone")
		Guide_Email=rs_Guide("Guide_Email")
end if
	NumDayGuide="XXX"
	AvarageGrade="XXX%"
	%>

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
	strXLS.Add  "<x:Name>דוח שנתי למדריך</x:Name>"
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
   strXLS.Add  "<tr style='background:#C5D9F1;color:#000000'><td colspan=6 align=right class=xl24 style='background:#C5D9F1;color:#000000'>נכון לתאריך " : strXLS.Add Now() : strXLS.Add"</td></tr>"

    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=6 align=right class=xl24 style='background:#808080;color:#ffffff'>"   : strXLS.Add "משובים דוח שנתי למדריך  " : strXLS.Add Guide_Name : strXLS.Add "</td></tr>"
       strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=6 align=right class=xl24 style='background:#808080;color:#ffffff'>" : strXLS.Add "לתקופה מ-" : strXLS.Add FromDate : strXLS.Add "    עד-": strXLS.Add ToDate

      IF Guide_Phone<>"" THEN
         strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=6 align=right class=xl24 style='background:#808080;color:#ffffff'>" : strXLS.Add "טלפון:"  : strXLS.Add Guide_Phone :  strXLS.Add " </TD></TR>"
    END IF
    IF Guide_Email<>"" THEN
      strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=6 align=right class=xl24 style='background:#808080;color:#ffffff'>"  : strXLS.Add Guide_Email: strXLS.Add " </TD></TR>"
    END IF
     
' strXLS.Add  "<tr style='background:#C5D9F1;color:#000000'><td colspan=6 align=right class=xl24 style='background:#C5D9F1;color:#000000'> ציון ממוצע שנתי של מדריך: " : strXLS.Add AvarageGrade : strXLS.Add"</td></tr>"
	sql = "Exec dbo.[get_GuideFeedbackReport]  '"& guideId & "','"& FromDate & "','"& ToDate  &"'"
'Response.Write sql
'Response.End
	set rsa = conPegasus.Execute(sql)
'If Not rsa.Eof Then
' Departure_Code=rsa("Departure_Code")
Tour_GradeAll=0
i=0
GuideDate=0
 Do while Not rsa.eof 
 if IsDate(rsa("Departure_Date")) and IsDate(rsa("Departure_Date_End")) then
 GuideDate=GuideDate + datediff("d",rsa("Departure_Date"), rsa("Departure_Date_End"))
end if
if rsa("Guide_Grade")>0 and  rsa("Guide_Grade")<80 then 
  strBody= strBody &   "<tr style=background-color:#FFC7CE>"
elseif rsa("Guide_Grade")>80 then 
strBody= strBody & "<tr style=background-color:#C6EFCE>"
else 
 strBody= strBody &  "<tr>"
end if
strBody= strBody & "<td class=xl26>" & rsa("Departure_Code") &  "</td>"
strBody= strBody & "<td class=xl26>" & rsa("Countries") & "</td>"
if rsa("Guide_Grade")>0 then
strBody= strBody & "<td class=xl26>" & rsa("Guide_Grade")   & "%</td>"
else
strBody= strBody & "<td class=xl26>&nbsp;</td>"
end if
strBody= strBody &  "<td class=xl26>" &rsa("CountFeedBack")& "</td>"
strBody= strBody &  "<td class=xl26>" & rsa("CountMembers") & "</td>"
if CDbl(rsa("Tour_Grade"))>0 then
strBody= strBody & "<td class=xl26>"&  rsa("Tour_Grade")& "%</td>"
i=i+1
Tour_GradeAll=Tour_GradeAll+CDbl(rsa("Tour_Grade"))
else
strBody= strBody & "<td class=xl26>&nbsp;</td>"
end if


rsa.movenext
'
 Loop
 'end if


  if Tour_GradeAll>0 then
   strXLS.Add  "<tr style='background:#C5D9F1;color:#000000'><td colspan=6 align=right class=xl24 style='background:#C5D9F1;color:#000000'> ציון ממוצע שנתי של מדריך: " : strXLS.Add Round(Tour_GradeAll/i,1) : strXLS.Add"%</td></tr>"

  end if
  strTop=  "<tr style='background:#C5D9F1;color:#000000'><td colspan=6 align=right class=xl24 style='background:#C5D9F1;color:#000000'> כמות הימים שהדרכת בפגסוס השנה היא: " & GuideDate & "</td></tr>"

	strTop= strTop &  "<tr>"
	strTop= strTop &  "<td class=xl25_blue>קוד טיול</td>"
	strTop= strTop &   "<td class=xl25_blue>מדינה בה הטיול מתרחש</td>"	
	strTop= strTop &   "<td class=xl25_blue>ציון מדריך בטיול</td>"
	strTop= strTop &   "<td class=xl25_blue>כמות האנשים שמילאו משוב</td>"
	strTop= strTop &   "<td class=xl25_blue>כמות האנשים הרשומים לטיול</td>"
	strTop= strTop &   "<td class=xl25_blue>ציון סופי של הטיול</td>"
   strTop= strTop &  "</tr>"

strXLS.Add  " " : strXLS.Add strTop : strXLS.Add strBody
   
   
       strXLS.Add  "</TABLE></body></html>"    
	
	filestring = "FeedbackGuideExcel.xls"
    Response.Clear()
    Response.AddHeader "content-disposition", "inline; filename=" & filestring
    Response.ContentType = "application/vnd.ms-excel"
    Response.Write(strXLS.value)  
    Set strXLS = Nothing %>

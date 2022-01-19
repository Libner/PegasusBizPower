
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
	strXLS.Add  "<x:Name>השוואה של טפסי מתעניין ואחוזי סגירה, לפי 3 שנים</x:Name>"
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
    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=8 class=xl24 style='background:#808080;color:#ffffff'>השוואה דוח</td></tr>"

    strXLS.Add  "<tr style='background:#808080;color:#ffffff'><td colspan=8 class=xl24 style='background:#808080;color:#ffffff'> " & start_date & " - " & end_date & "  תקופת  הייחוס </td></tr>"
  	'---------------P2932 - type of form- REASONS ---------------
   	If reasons<>"" Then
		sqlstr = "SELECT  Reason_Title FROM  Appeals_CreationReasons where QUESTIONS_ID=16504 and Reason_Id in (" & reasons & ") order by Reason_Order"
			set rs_Reason = con.getRecordSet(sqlstr)
			If Not rs_Reason.eof Then
				arr_Reason = rs_Reason.getRows()
			End If
			Set rs_Reason = Nothing
			If isArray(arr_Reason) Then
			
    strXLS.Add  "<tr  style='background:#808080;color:#ffffff'><td colspan=8  class=xl24 style='background:#808080;color:#ffffff'> הגורם ליצירת הטופס: "
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
  	strXLS.Add  "<tr>"
		strXLS.Add  "<td class=xl25_blue colspan=5 x:str style='background:#C5D9F1'>ביקושים</td>"
		strXLS.Add  "<td class=xl25_blue colspan=3  x:str>אחוזי סגירה</td></tr>"
	 	strXLS.Add  "<tr><td class=xl25_blueRight  x:str> </td>"
	 	strXLS.Add  "<td class=xl25_blue x:str>ביקוש " : strXLS.Add DateAdd("yyyy",-2,start_date) :  strXLS.Add " - " : strXLS.Add DateAdd("yyyy",-2,end_date) : strXLS.Add "</td>"
 		strXLS.Add  "<td class=xl25_blue x:str>ביקוש " : strXLS.Add DateAdd("yyyy",-1,start_date) :  strXLS.Add " - " : strXLS.Add DateAdd("yyyy",-1,end_date) : strXLS.Add "</td>"
	 	strXLS.Add  "<td class=xl25_blue x:str>ביקוש " : strXLS.Add start_date :  strXLS.Add " - " : strXLS.Add end_date : strXLS.Add"</td>"
	 	strXLS.Add  "<td class=xl25_blue x:str>ביקוש %" : strXLS.Add "</td>"
	 	strXLS.Add  "<td class=xl25_blue x:str>" : strXLS.Add DateAdd("yyyy",-2,start_date) :  strXLS.Add " - " : strXLS.Add DateAdd("yyyy",-2,end_date) : strXLS.Add "</td>"
	 	strXLS.Add  "<td class=xl25_blue x:str>" : strXLS.Add DateAdd("yyyy",-1,start_date) :  strXLS.Add " - " : strXLS.Add DateAdd("yyyy",-1,end_date) : strXLS.Add "</td>"
	 strXLS.Add  "<td class=xl25_blue x:str>" : strXLS.Add start_date :  strXLS.Add " - " : strXLS.Add end_date : strXLS.Add"</td>"
				
		strXLS.Add  "</tr>"
  
     
  
   If Right(countryID, 1) = "," Then 
    countryID = Left(countryID, Len(countryID) - 1) 
 
  End If
  
sqlQuery=" and appeal_CountryId   in ("& countryID &")"
'response.Write "countryid-="&countryid   
  
		 sqlstr = "SELECT distinct appeal_CountryId,dbo.get_Country_CRMName(appeal_CountryId) Country_CRMName FROM dbo.APPEALS APP where appeal_CountryId   in ("& countryID &") order by  dbo.get_Country_CRMName(appeal_CountryId),appeal_CountryId"				
  	'response.Write sqlstr
  '	response.end
  	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	
 do while not rstitle.EOF 
 countryName=rstitle("Country_CRMName")
 pcountryid=rstitle("appeal_CountryId")
 
 
		    pquest_id=16504
     sql = "Exec dbo.get_AppealsCount16504  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @countryID=" & nFix(pcountryid) & ", @productID=" & nFix(pquest_id)   & ", @reasons='" & reasons &  "'"
'response.Write sql
'response.end
	set rs_count = con.getRecordSet(sql)
	if not rs_count.eof then
		appealsCount16504 = cLng(rs_count(0))
		else
		appealsCount16504=""
	end if
	set rs_count = nothing
	

  sql = "Exec dbo.get_AppealsCount16504  @start_date='" & DateAdd("yyyy",-1,start_date) & "', @end_date='" & DateAdd("yyyy",-1,end_date) & "'," & _
	" @countryID=" & nFix(pcountryid) & ", @productID=" & nFix(pquest_id)   & ", @reasons='" & reasons &  "'"
	set rs_count = con.getRecordSet(sql)
	if not rs_count.eof then
		appealsCount16504_1 = cLng(rs_count(0))
		else
		appealsCount16504_1=""
	end if
	set rs_count = nothing
	
	 sql = "Exec dbo.get_AppealsCount16504  @start_date='" & DateAdd("yyyy",-2,start_date) & "', @end_date='" & DateAdd("yyyy",-2,end_date) & "'," & _
	" @countryID=" & nFix(pcountryid) & ", @productID=" & nFix(pquest_id)   & ", @reasons='" & reasons &  "'"
	set rs_count = con.getRecordSet(sql)
	if not rs_count.eof then
		appealsCount16504_2 = cLng(rs_count(0))
		else
		appealsCount16504_2=""
	end if
	set rs_count = nothing
		
		if appealsCount16504_1<>0 then
	calculateappealsCount=Round((cLng(appealsCount16504)/cLng(appealsCount16504_1))*100)
	else
	calculateappealsCount=Round((cLng(appealsCount16504)/1)*100)
	end if
 
  
		strXLS.Add  "<TR><td class=xl25_white  x:str>" : strXLS.Add countryName: strXLS.Add "</td>"
        strXLS.Add  "<td class=xl25_white  x:str>" : strXLS.Add  appealsCount16504_2: strXLS.Add "</td>"
	
		 if appealsCount16504_1=appealsCount16504_2 then
	    strXLS.Add  "<td class=xl25_white  x:str>" : strXLS.Add  appealsCount16504_1: strXLS.Add "</td>"
	    elseif appealsCount16504_1>appealsCount16504_2 then
	       strXLS.Add  "<td class=xl25_green  x:str>" : strXLS.Add  appealsCount16504_1: strXLS.Add "</td>"
	    else
	          strXLS.Add  "<td class=xl25_red  x:str>" : strXLS.Add  appealsCount16504_1: strXLS.Add "</td>"
	    end if
	    
	 if appealsCount16504=appealsCount16504_1 then
		    strXLS.Add  "<td class=xl25_white   x:str>" : strXLS.Add  appealsCount16504: strXLS.Add "</td>"
	elseif appealsCount16504>appealsCount16504_1 then
			strXLS.Add  "<td class=xl25_green  x:str>" : strXLS.Add  appealsCount16504: strXLS.Add "</td>"
	else
		    strXLS.Add  "<td class=xl25_red  x:str>" : strXLS.Add  appealsCount16504: strXLS.Add "</td>"
	end if
	 if calculateappealsCount>100 then
	      strXLS.Add  "<td class=xl25_green  x:str>" : strXLS.Add  calculateappealsCount : strXLS.Add "%</td>"
	      elseif calculateappealsCount=100 then
     strXLS.Add  "<td class=xl25_white  x:str>" : strXLS.Add  calculateappealsCount : strXLS.Add "%</td>"
	else
  strXLS.Add  "<td class=xl25_red  x:str>" : strXLS.Add  calculateappealsCount : strXLS.Add "%</td>"
	
	      end if
	      
	
start_date1=DateAdd("yyyy",-1,start_date)
start_date2=DateAdd("yyyy",-2,start_date)

end_date1=DateAdd("yyyy",-1,end_date)
end_date2=DateAdd("yyyy",-2,end_date)

'select from 16735
if reasons="" then
	sqlall="SELECT COUNT(APPEAL_ID) as CountAll FROM APPEALS where appeal_CountryId="& nFix(pcountryid) &"  and QUESTIONS_ID="& quest_id  & _
	" and RelationId>0 and (DateDiff(d,appeal_date, convert(smalldatetime, '"& start_date &"' ,103)) <= 0 OR (Len('"& start_date &"' ) = 0 )) " & _
	" AND (DateDiff(d,appeal_date, convert(smalldatetime,'" & end_date &"' ,103)) >= 0 OR (Len('"& end_date &"') = 0))"
else
	sqlall="SELECT COUNT(APPEALS.APPEAL_ID) as CountAll FROM APPEALS inner join APPEALS app16504 on APPEALS.RelationId=app16504.APPEAL_ID " & _
	" where APPEALS.appeal_CountryId="& nFix(pcountryid) &"  and APPEALS.QUESTIONS_ID="& quest_id  & _
	" and APPEALS.RelationId>0 and (DateDiff(d,APPEALS.appeal_date, convert(smalldatetime, '"& start_date &"' ,103)) <= 0 OR (Len('"& start_date &"' ) = 0 )) " & _
	" AND (DateDiff(d,APPEALS.appeal_date, convert(smalldatetime,'" & end_date &"' ,103)) >= 0 OR (Len('"& end_date &"') = 0)) " & _
	" AND app16504.Reason_Id in (" & reasons & ") "
end if
'response.Write sqlall
'response.end
set rsaAll= con.getRecordSet(sqlall)
	if not rsaAll.eof then
	CountAll=rsaAll("CountAll")
	else
	CountAll=0	
	end if
	
if reasons="" then
	sqlall1="SELECT COUNT(APPEAL_ID) as CountAll1 FROM APPEALS where appeal_CountryId="& nFix(pcountryid) &"  and QUESTIONS_ID="& quest_id  & _
	" and RelationId>0 and(DateDiff(d,appeal_date, convert(smalldatetime, '"& start_date1 &"' ,103)) <= 0 OR (Len('"& start_date1 &"' ) = 0 )) " & _
	" AND (DateDiff(d,appeal_date, convert(smalldatetime,'" & end_date1 &"' ,103)) >= 0 OR (Len('"& end_date1 &"') = 0))"
else
	sqlall1="SELECT COUNT(APPEALS.APPEAL_ID) as CountAll1 FROM APPEALS inner join APPEALS app16504 on APPEALS.RelationId=app16504.APPEAL_ID " & _
	" where APPEALS.appeal_CountryId="& nFix(pcountryid) &"  and APPEALS.QUESTIONS_ID="& quest_id  & _
	" and APPEALS.RelationId>0 and(DateDiff(d,APPEALS.appeal_date, convert(smalldatetime, '"& start_date1 &"' ,103)) <= 0 OR (Len('"& start_date1 &"' ) = 0 )) " & _
	" AND (DateDiff(d,APPEALS.appeal_date, convert(smalldatetime,'" & end_date1 &"' ,103)) >= 0 OR (Len('"& end_date1 &"') = 0)) " & _
	" AND app16504.Reason_Id in (" & reasons & ") "
end if
'response.Write sqlall1
'response.end
set rsaAll1= con.getRecordSet(sqlall1)
	if not rsaAll1.eof then
	CountAll1=rsaAll1("CountAll1")
	else
	CountAll1=0	
	end if
	
	
	
if reasons="" then
	sqlall2="SELECT COUNT(APPEAL_ID) as CountAll2 FROM APPEALS where appeal_CountryId="& nFix(pcountryid) &"  and QUESTIONS_ID=" & quest_id & _
	" and RelationId>0 and(DateDiff(d,appeal_date, convert(smalldatetime, '"& start_date2 &"' ,103)) <= 0 OR (Len('"& start_date2 &"' ) = 0 )) " & _
	" AND (DateDiff(d,appeal_date, convert(smalldatetime,'" & end_date2 &"' ,103)) >= 0 OR (Len('"& end_date2 &"') = 0))"
else
	sqlall2="SELECT COUNT(APPEALS.APPEAL_ID) as CountAll2 FROM APPEALS inner join APPEALS app16504 on APPEALS.RelationId=app16504.APPEAL_ID " & _
	" where APPEALS.appeal_CountryId="& nFix(pcountryid) &"  and APPEALS.QUESTIONS_ID=" & quest_id & _
	" and APPEALS.RelationId>0 and(DateDiff(d,APPEALS.appeal_date, convert(smalldatetime, '"& start_date2 &"' ,103)) <= 0 OR (Len('"& start_date2 &"' ) = 0 )) " & _
	" AND (DateDiff(d,APPEALS.appeal_date, convert(smalldatetime,'" & end_date2 &"' ,103)) >= 0 OR (Len('"& end_date2 &"') = 0))" & _
	" AND app16504.Reason_Id in (" & reasons & ") "
end if

set rsaAll2= con.getRecordSet(sqlall2)
	if not rsaAll2.eof then
	CountAll2=rsaAll2("CountAll2")
	else
	CountAll2=0	
	end if
	'response.Write  "CountAll="&cint(CountAll) &"--<Br>"&"appealsCount16504=" &cint(appealsCount16504) &"--"
	'response.Write "<BR>" & cint(CountAll)*100/cint(appealsCount16504) &"<BR>"
	'response.end
	if appealsCount16504<>0 then
	pSum= Round(cint(CountAll)*100/cint(appealsCount16504)) 
	else
	pSum=0
	end if
	if appealsCount16504_1<>0 then
			pSum1= Round(cLng(CountAll1)*100/cLng(appealsCount16504_1)) 
	else
		pSum1= Round(cLng(CountAll1)*100/1) 
	end if
	if appealsCount16504_2<>0 then
       pSum2= Round(cLng(CountAll2)*100/cLng(appealsCount16504_2)) 
	else
	pSum2= Round(cLng(CountAll2)*100/1) 
	end if
	
	'response.Write calculateappealsCount &":"& appealsCount16504 &":"& appealsCount16504_1
	'response.end
	sumAppealsCount16504_2=sumAppealsCount16504_2+appealsCount16504_2
	sumAppealsCount16504_1=sumAppealsCount16504_1+appealsCount16504_1
	sumAppealsCount16504=sumAppealsCount16504+appealsCount16504
	 strXLS.Add  "<td class=xl25_white  x:str>": strXLS.Add pSum2: strXLS.Add "% ( ":  strXLS.Add CountAll2 :  strXLS.Add " )</td>"
	strXLS.Add  "<td class=xl25_white  x:str>": strXLS.Add pSum1: strXLS.Add "% ( " :  strXLS.Add CountAll1 :  strXLS.Add " )</td>"
		if pSum>pSum1 then
		strXLS.Add "<td class=xl25_green  x:str>": strXLS.Add pSum: strXLS.Add "% ( ":  strXLS.Add CountAll :  strXLS.Add " )</td></tr>"
		elseif pSum=pSum1 then
			strXLS.Add "<td class=xl25_white  x:str>": strXLS.Add pSum: strXLS.Add "% ( ":  strXLS.Add CountAll :  strXLS.Add " )</td></tr>"
			else
			strXLS.Add "<td class=xl25_red  x:str>": strXLS.Add pSum: strXLS.Add "% ( ":  strXLS.Add CountAll :  strXLS.Add " )</td></tr>"

			end if		'		strXLS.Add  "<td class=xl25_blue  x:str>pSum1-": strXLS.Add pSum1: strXLS.Add"</td><td class=xl25_blue  x:str>pSum-": strXLS.Add pSum:strXLS.Add "-":strXLS.Add CountAll:strXLS.Add "-":strXLS.Add appealsCount16504: strXLS.Add"</td></tr>"
		
		
		
		'strXLS.Add  "<tr><td style='background:#92D050 ;height:6px' colspan=10></td></tr>"

		
	
	


	'Set con1 = Nothing	 
		rstitle.moveNext
		loop
		end if
		rstitle.close
		set rstitle=Nothing	
		strXLS.Add  "<TR><td class=xl25_white COLSPAN=8 x:str></td></TR>"
   		strXLS.Add  "<TR><td class=xl25_white  x:str>סה``כ</td>"
        strXLS.Add  "<td class=xl25_white  x:str>" : strXLS.Add  sumAppealsCount16504_2: strXLS.Add "</td>"

     strXLS.Add  "<td class=xl25_white  x:str>" : strXLS.Add  sumAppealsCount16504_1: strXLS.Add "</td>"
     strXLS.Add  "<td class=xl25_white  x:str>" : strXLS.Add  sumAppealsCount16504: strXLS.Add "</td>"
     strXLS.Add  "<td class=xl25_white colspan=4  x:str></td></tr>"



   
       strXLS.Add  "</TABLE></body></html>"    
	Set con = Nothing	 
	
	filestring = "RegistrationForm.xls"
    Response.Clear()
    Response.AddHeader "content-disposition", "inline; filename=" & filestring
    Response.ContentType = "application/vnd.ms-excel"
    Response.Write(strXLS.value)  
    Set strXLS = Nothing %>

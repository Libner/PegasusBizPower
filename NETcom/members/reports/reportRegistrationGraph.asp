<% Response.CharSet = "windows-1255"
       Response.Buffer = True 
       Server.ScriptTimeout = 1000  %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->

<!--#include file="code.asp" -->
<%  Session.LCID = 1037
	quest_id = 16735
	countryid=trim(Request("country_id"))

	start_date = DateValue(trim(Request("dateStart")))
    end_date = DateValue(trim(Request("dateEnd")))
 	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"  :  	align_var = "left"  :  dir_obj_var = "ltr" 
	Else
		dir_var = "ltr"  :  	align_var = "right"  :  	dir_obj_var = "rtl"  
	End If		  
		  
sqlQuery=" and appeal_CountryId   in ("& countryID &")"
'response.Write "countryid-="&countryid   
   
   	sql = "Exec dbo.get_reportRegistration  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @sqlQuery='" & sqlQuery & "', @productID=" & nFix(quest_id) 
	Response.Write sql
	Response.End
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
     sql = "Exec dbo.get_AppealsCount16504  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @countryID=" & nFix(pcountryid) & ", @productID=" & nFix(pquest_id) 
	set rs_count = con.getRecordSet(sql)
	if not rs_count.eof then
		appealsCount16504 = cLng(rs_count(0))
		else
		appealsCount16504=""
	end if
	set rs_count = nothing
	

	 sql = "Exec dbo.get_AppealsCount16504_NotExists  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @countryID=" & nFix(pcountryid) & ", @productID=" & nFix(pquest_id) & ", @StatusId=3" 'status סגור
	set rs_count = con.getRecordSet(sql)
	if not rs_count.eof then
		appealsCountNotExist_Close16504 = cLng(rs_count(0))
		else
		appealsCountNotExist_Close16504=""
	end if
	set rs_count = nothing

 sql = "Exec dbo.get_AppealsCount16504_NotExists  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @countryID=" & nFix(pcountryid) & ", @productID=" & nFix(pquest_id) & ", @StatusId=4" 'status במעקב
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
		strXLS.Add  "<td class=xl25_blue  x:str>טפסי נתעניין</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>נרשמו</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>% סגירה מכירה ל" : strXLS.Add countryName : strXLS.Add" הוא</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>כמות טפסים מתעניין ללא טופס רישום בסטטוס סגור</td>"
		strXLS.Add  "<td class=xl25_blue  x:str>כמות טפסים מתעניין ללא טופס רישום בסטטוס מעקב</td></tr>"
	
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

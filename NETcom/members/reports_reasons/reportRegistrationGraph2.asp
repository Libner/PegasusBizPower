<!--#include file="../../connect.asp"-->

<!--#include file="../../reverse.asp"-->

<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<TITLE></TITLE>
</head>

<%Session.LCID = 1037
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
%>
<body topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 bgcolor="#E5E5E5">

<div id="div_save" name="div_save" style="position:absolute; left:0px; top:0px; width:100%; height:100%;">  												
  <table height="100%" width="100%" cellspacing="2" cellpadding="2" border="0" ID="Table1">
        <tr><td align="center" valign=middle>דוח התפלגות מועדי קבלת טפסי רישום</td></tr>
               <tr><td align="center" valign=middle>ביחס למועד פתיחת טופס מתעניין לפי יעדים </td></tr>
                      <tr><td align="center" valign=middle>בין תאריכים  <%=start_date%> - <%=end_date%>  </td></tr>
                         	
   <%'---------------P2932 - type of form- REASONS ---------------
   	If reasons<>"" Then
		sqlstr = "SELECT  Reason_Title FROM  Appeals_CreationReasons where QUESTIONS_ID=16504 and Reason_Id in (" & reasons & ") order by Reason_Order"
			set rs_Reason = con.getRecordSet(sqlstr)
			If Not rs_Reason.eof Then
				arr_Reason = rs_Reason.getRows()
			End If
			Set rs_Reason = Nothing
			If isArray(arr_Reason) Then
			%>
			<tr ><td  align="center" valign=middle> הגורם ליצירת הטופס:
 				<%For mm=0 To Ubound(arr_Reason,2)
					if mm>0 then%>
					,&nbsp;
					<%end if%>
					<%= trim(arr_Reason(0,mm))%>
				<%next%>
				</td></tr>
<%
			end if
	end if
	%>
     <tr><td align="center" valign=middle>
     <table dir=ltr border="1" height="800" width="800" cellspacing="0" cellpadding="0" ID="Table2">
  <%   sqlQuery=" and appeal_CountryId   in ("& countryID &")"
'	response.Write "countryid-="&countryid   
   
   	sql = "Exec dbo.get_reportRegistration_Reasons  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @sqlQuery='" & sqlQuery & "', @productID=" & nFix(quest_id)  & ", @reasons='" & reasons &  "'"
	'Response.Write sql
'Response.End
	set rsa = con.getRecordSet(sql)
	If Not rsa.Eof Then
		arr_app = rsa.getRows()
	End If
	Set rsa = Nothing	%>

 <tr>
 	<%If isArray(arr_app) Then
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
	sumCountry=	Cint(count02) +Cint(count3_7) +Cint(count8_14)+Cint(count15_21)+Cint(count22_30)+Cint(count31_60)+Cint(count61_90)+Cint(count91_180)
		pcount02=Round(100*Cint(count02)/Cint(sumCountry))
		pcount3_7=Round(100*Cint(count3_7)/Cint(sumCountry))
		pcount8_14=Round(100*Cint(count8_14)/Cint(sumCountry))
		pcount15_21=Round(100*Cint(count15_21)/Cint(sumCountry))
		pcount22_30=Round(100*Cint(count22_30)/Cint(sumCountry))
		pcount31_60=Round(100*Cint(count31_60)/Cint(sumCountry))
		pcount61_90=Round(100*Cint(count61_90)/Cint(sumCountry))
		pcount91_180=Round(100*Cint(count91_180)/Cint(sumCountry))
			    pquest_id=16504
     sql = "Exec dbo.get_AppealsCount16504_Reasons  @start_date='" & start_date & "', @end_date='" & end_date & "'," & _
	" @countryID=" & nFix(pcountryid) & ", @productID=" & nFix(pquest_id)  & ", @reasons='" & reasons &  "'"
	'Response.Write sql
'Response.End
	set rs_count = con.getRecordSet(sql)
	if not rs_count.eof then
		appealsCount16504 = cLng(rs_count(0))
		else
		appealsCount16504=""
	end if
	set rs_count = nothing
	
			appealsCount16504All=appealsCount16504All+appealsCount16504
		sumCountryAll=sumCountryAll+sumCountry
		%>
	
 
 <td valign=bottom>
 <table border=0 cellpadding=0 cellspacing=0 border=0> 
 <tr><td align=center>
<table border=0 cellpadding=0 cellspacing=0 style="height:800px">
 <tr><td style="height:<%=100*Cint(count91_180)/Cint(sumCountry)%>%;background-color:#000080" align=center><table border=0 cellpadding=2 cellspacing=2 ><tr><td bgcolor=#ffffff><span style=font-size:14pt;font-weight:bold"><%=pcount91_180%>%</span></td></tr></table></td></tr>
 <tr><td style="height:<%=100*Cint(count61_90)/Cint(sumCountry)%>%;background-color:#008080" align=center><table border=0 cellpadding=2 cellspacing=2 ID="Table3"><tr><td bgcolor=#ffffff><span style=font-size:14pt;font-weight:bold"><%=pcount61_90%>%</span></td></tr></table></td></tr>

 <tr><td style="height:<%=100*Cint(count31_60)/Cint(sumCountry)%>%;background-color:#F27027" align=center><table border=0 cellpadding=2 cellspacing=2 ID="Table4"><tr><td bgcolor=#ffffff><span style=font-size:14pt;font-weight:bold"><%=pcount31_60%>%</span></td></tr></table></td></tr>

 <tr><td style="height:<%=100*Cint(count22_30)/Cint(sumCountry)%>%;background-color:#008080" align=center><table border=0 cellpadding=2 cellspacing=2 ID="Table5"><tr><td bgcolor=#ffffff><span style=font-size:14pt;font-weight:bold"><%=pcount22_30%>%</span></td></tr></table></td></tr>

 <tr><td style="height:<%=100*Cint(count15_21)/Cint(sumCountry)%>%;background-color:#808000" align=center><table border=0 cellpadding=2 cellspacing=2 ID="Table6"><tr><td bgcolor=#ffffff><span style=font-size:14pt;font-weight:bold"><%=pcount15_21%>%</span></td></tr></table></td></tr>

 <tr><td style="height:<%=100*Cint(count8_14)/Cint(sumCountry)%>%;background-color:#800000" align=center><table border=0 cellpadding=2 cellspacing=2 ID="Table7"><tr><td bgcolor=#ffffff><span style=font-size:14pt;font-weight:bold"><%=pcount8_14%>%</span></td></tr></table></td></tr>

  <tr><td style="height:<%=100*Cint(count3_7)/Cint(sumCountry)%>%;background-color:#808080" align=center><table border=0 cellpadding=2 cellspacing=2 ID="Table8"><tr><td bgcolor=#ffffff><span style=font-size:14pt;font-weight:bold"><%=pcount3_7%>%</span></td></tr></table></td></tr>

 <tr><td style="height:<%=100*Cint(count02)/Cint(sumCountry)%>%;background-color:#FFFF00" align=center><table border=0 cellpadding=2 cellspacing=2 ID="Table9"><tr><td bgcolor=#ffffff><span style=font-size:14pt;font-weight:bold"><%=pcount02%>%</span></td></tr></table></td></tr>

 </table>
 
 
 </td></tr>
 
<tr><td valign=bottom align=center width="<%=100/(Ubound(arr_app,2)+1)%>%"><span style=font-size:14pt;font-weight:bold"><%=countryName%></span></td></tr> 
<tr><td valign=bottom align=center width="<%=100/(Ubound(arr_app,2)+1)%>%"><span style=font-size:12pt;font-weight:normal"><%=appealsCount16504All%>/<%=sumCountryAll%></span></td></tr> 
<%if false then%><tr><td valign=bottom align=center width="<%=100/(Ubound(arr_app,2)+1)%>%"><span style=font-size:14pt;font-weight:bold">טפסי נתעניין/נרשמו</span></td></tr> <%end if%>


 </table>
 
 
 </td>
 <%
 appealsCount16504All=0
 sumCountryAll=0

 
 Next
 end if
 %>



 </tr>
 
 
 
      </table>
    </td></tr>
<tr><TD>

    <Table border=0 cellpadding=0 cellspacing=0 align=center width="800px" ID="Table10">
    <tr>
<td style="height:20px;width:100px;background-color:#FFFF00" align=center>0-2 ימים</td>
<td style="height:20px;width:100px;background-color:#808080" align=center>3-7 ימים</td>

<td style="height:20px;width:100px;background-color:#800000" align=center>8-14 ימים</td>

<td style="height:20px;width:100px;background-color:#808000" align=center>15-21 ימים</td>

<td style="height:20px;width:100px;background-color:#008080" align=center>22-30 ימים</td>

<td style="height:20px;width:100px;background-color:#F27027" align=center>31-60 ימים</td>

<td style="height:20px;width:100px;background-color:#008080" align=center>61-90 ימים</td>

<td style="height:20px;width:100px;background-color:#000080" align=center><span style="color:#ffffff">91-180 ימים</span></td>
</tr>
</table>
</TD></tr>
   <tr><td height=10></td></tr> 
    </table>




</div>
</body>
</html>
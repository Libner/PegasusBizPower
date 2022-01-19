<% server.ScriptTimeout=10000 %>
<% Response.Buffer = True %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%  
  
  sort = Request.QueryString("sort")     
  companyID = trim(Request("companyID"))
   if trim(lang_id) = "1" then
	   arr_status_comp = Array("","עתידי","פעיל","סגור","פונה")
   else
	   arr_status_comp = Array("","new","active","close","appeal")
   end if	     
  
  sqlStr = "SELECT company_name,address,city_Name"&_
  ", prefix_phone,phone,prefix_fax,fax"&_
  ", date_update, status,contact_id FROM companies_private_view WHERE company_id="& companyID 
  'Response.Write "test"
  'Response.End
  set pr=con.GetRecordSet(sqlStr)
  if not pr.EOF then	
	company_name  =pr("company_name")  
	address	      =pr("address")	
	cityName	  =pr("city_Name")
	prefix_phone  =pr("prefix_phone")	
	prefix_fax	  =pr("prefix_fax")	
	phone	      =pr("phone")	
	fax           =pr("fax")			
	date_update	  =pr("date_update")
	status_company =pr("status")
	contactID	   =pr("contact_id")			    
    If isNumeric(trim(companyID)) Then
		types=""		
		sqlstr="Select type_Name From company_to_types_view Where company_id = " & companyID & " Order By type_id"
		set rssub = con.getRecordSet(sqlstr)			
		If not rssub.eof Then
			types = rssub.getString(,,",",",") 				
		End If
		set rssub=Nothing
		If Len(types) > 0 Then
			types = Left(types,(Len(types)-1))
		End If					
		
		sqlstr = "Select COMPANY_DESC from companies WHERE company_id = " & companyID
		set rs_desc = con.getRecordSet(sqlstr)
		if not rs_desc.eof then
			comp_desc = rs_desc(0)
		end if
		set  rs_desc = nothing
      
    End If    
      
  End If   
 
 If trim(contactID)<>"" then 
   sql="SELECT company_ID,contact_ID,email,prefix_phone,phone,prefix_cellular,prefix_fax,fax,"&_
   "cellular,CONTACT_NAME,messanger_name,date_update FROM contacts"&_
   " WHERE contact_ID=" & contactID
   set listContact=con.GetRecordSet(sql)
   if not listContact.EOF then 
      contactID=listContact("contact_ID") 
      companyID=listContact("company_ID") 
      CONTACT_NAME=listContact("CONTACT_NAME")  
      contacter = trim(CONTACT_NAME)
      email=listContact("email")
      prefix_phone=listContact("prefix_phone")
      phone=listContact("phone")
      prefix_cellular=listContact("prefix_cellular")
      cellular=listContact("cellular")
      prefix_fax=listContact("prefix_fax")
      fax=listContact("fax")       
      messangerName=listContact("messanger_name")     
    End if   
    set  listContact = Nothing
  End If  
 
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 14 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing
  
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<style>
TD
{
	font-family: Arial;
	font-size: 14px;
	font-weight: normal
}
TD.title
{
	font-weight: bolder
}
.Form_R
{
	font-family: Arial;
	font-size: 13px;
	font-weight: normal

}
.page_title
{
	color: Black;
	font-size:16px;
	font-weight:bold
}
</style>
</head>
<body style="margin:0;" onload="javascript:window.print();">
<table border="0" bordercolor="navy" width=640 cellspacing="0" cellpadding="0" align="center" valign="top">
<tr bgcolor=#E2E2E2>
<td class="page_title" width=100% align=center style="line-height:25px" dir="<%=dir_obj_var%>">&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<font color="#6F6DA6"><%=company_name%></font>&nbsp;</td></tr>         	
<tr><td height=1 bgcolor=#808080 nowrap colspan=2></td></tr>
<tr><td height=10 nowrap colspan=2></td></tr>
<tr><td width=640 colspan=2>
<table dir="<%=dir_var%>" cellpadding=2 cellspacing=0 width=630 align=center border=1 BORDERCOLOR=black style="border-collapse:collapse">
  <tr><td colspan=2 class="title" align="<%=align_var%>" bgcolor=#E6E6E6><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;</td></tr>
  <tr>
     <td width=520 align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=company_name%>&nbsp;</td>
     <td width="110" nowrap align="<%=align_var%>" class=title style="padding-right:15px">&nbsp;<span id="word14" name=word14><!--שם מלא--><%=arrTitles(14)%></span>&nbsp;</td> 
  </tr> 
      <tr>
         <td align="<%=align_var%>" dir="<%=dir_obj_var%>">                      
         &nbsp;<%=messangerName%>&nbsp;
         </td>         
         <td align="<%=align_var%>" class=title style="padding-right:15px">&nbsp;<span id="word15" name=word15><!--תפקיד--><%=arrTitles(15)%></span>&nbsp;</td>
      </tr>            
      <tr>
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=address%>&nbsp;</td>
        <td align="<%=align_var%>" class=title style="padding-right:15px">&nbsp;<span id="word5" name=word5><!--כתובת--><%=arrTitles(5)%></span>&nbsp;</td>
     </tr>             
     <tr>
         <td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>">&nbsp;<%=cityName%>&nbsp;</td>
         <td align="<%=align_var%>" class=title style="padding-right:15px">&nbsp;<span id="word6" name=word6><!--עיר--><%=arrTitles(6)%></span>&nbsp;</td>
      </tr>                        
      <tr>
          <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=phone%>&nbsp;</td>
          <td align="<%=align_var%>" class=title style="padding-right:15px">&nbsp;<span id="word16" name=word16><!--טלפון--><%=arrTitles(16)%></span>&nbsp;</td>
      </tr>                                
      <tr>  
          <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=fax%>&nbsp;</td>
          <td align="<%=align_var%>" class=title style="padding-right:15px">&nbsp;<span id="word17" name=word17><!--טלפון נייד--><%=arrTitles(17)%></span>&nbsp;</td>
      </tr>       
      <tr>  
          <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=cellular%>&nbsp;</td>
          <td align="<%=align_var%>" class=title style="padding-right:15px">&nbsp;<span id="word18" name=word18><!--פקס--><%=arrTitles(18)%></span>&nbsp;</td>
      </tr>                          
      <tr>
         <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=email%>&nbsp;</td>
         <td align="<%=align_var%>" class=title style="padding-right:15px">&nbsp;Email&nbsp;</td>
      </tr>              
      <tr>
           <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=types%>&nbsp;</td>
           <td align="<%=align_var%>" class=title style="padding-right:15px" valign=top>&nbsp;<span id="word11" name=word11><!--קבוצה--><%=arrTitles(11)%></span>&nbsp;</td>
      </tr>          
      <tr>
           <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=comp_desc%>&nbsp;</td>
           <td align="<%=align_var%>" class=title style="padding-right:15px" valign=top>&nbsp;<span id="word12" name=word12><!--פרטים נוספים--><%=arrTitles(12)%></span>&nbsp;</td>
      </tr>     
      <tr>
           <td align="<%=align_var%>">
           &nbsp;<%=arr_status_comp(status_company)%>&nbsp;
           </td>
           <td align="<%=align_var%>" class=title style="padding-right:15px" valign=top>&nbsp;<span id="word13" name=word13><!--סטטוס--><%=arrTitles(13)%></span>&nbsp;</td>
      </tr>
</table>
</td>
</tr>
<tr><td height=10 nowrap colspan=2></td></tr>
</table>
</body>
<%set con=Nothing%>
</html>


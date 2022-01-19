<% server.ScriptTimeout=10000 %>
<% Response.Buffer = True %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<% 
  companyID = trim(Request("companyID"))
  contactID = trim(Request("contactID"))
   if trim(lang_id) = "1" then
	   arr_status_comp = Array("","עתידי","פעיל","סגור","פונה")
   else
	   arr_status_comp = Array("","new","active","close","appeal")
   end if	     
  
   sql="SELECT contacts.company_ID,contact_ID,contacts.email,contacts.phone,contacts.fax,"&_
   " contacts.cellular,contacts.phone_additional,CONTACT_NAME,messanger_name,contacts.date_update,contact_city_name,contact_zip_code,"&_
   " contact_address, contact_desc, responsible_id, Users.FIRSTNAME + ' ' + Users.LASTNAME as responsible_name "&_
   " FROM contacts Inner Join Companies On contacts.company_ID = "&_
   " Companies.company_ID Left Outer Join Users On Contacts.Responsible_ID = Users.User_ID " &_
   " WHERE contact_ID=" & contactID & " And contacts.organization_id = " & OrgID
   set listContact=con.GetRecordSet(sql)
   if not listContact.EOF then 
      contactID=listContact("contact_ID") 
      companyID=listContact("company_ID") 
      CONTACT_NAME=listContact("CONTACT_NAME")  
      contacter = trim(CONTACT_NAME)
      email=listContact("email")      
      phone=listContact("phone")     
      cellular=listContact("cellular")      
      phone_additional=listContact("phone_additional")     
      fax=listContact("fax")       
      messangerName=listContact("messanger_name")
      contact_address = listContact("contact_address")
      contact_city_Name = listContact("contact_city_Name")
      contact_zip_code = listContact("contact_zip_code")
      contact_desc = listContact("contact_desc")
      responsible_id = listContact("responsible_id")
      responsible_name = listContact("responsible_name") 
    End if   
    set  listContact = Nothing
    
	If isNumeric(trim(contactId)) Then
		contact_types=""
		sqlstr = "get_contact_types '" & contactId & "'"
		set rstype = con.getRecordSet(sqlstr)		   						
		If not rstype.eof Then
			contact_types = rstype.getString(,,",",",")						
		End If		    
		set rstype=Nothing		
		If Len(contact_types) > 0 Then
			contact_types = Left(contact_types,(Len(contact_types)-1))
		End If	
	End If 
	
  If trim(companyID)<>"" then   
  sqlStr = "SELECT company_name,company_name_E,address,address2,city_Name"&_
  ", street_number,apartment,post_box,zip_code,prefix_phone1" &_
  ", prefix_phone2,prefix_fax1,prefix_fax2,phone1,phone2,fax1,fax2,url,email"&_
  ", date_update,status,company_desc,private_flag FROM companies WHERE company_id="& companyID &_ 
  " And organization_id = " & OrgID
  'Response.Write "test"
  'Response.End
  set pr=con.GetRecordSet(sqlStr)
  if not pr.EOF then	
	company_name   =pr("company_name")
  	company_name_E =pr("company_name_E")
	address	      =pr("address")
	address2	  =pr("address2")	
	street_number =pr("street_number")
	apartment	  =pr("apartment")
	post_box 	  =pr("post_box")	
	cityName	  =pr("city_Name")
	zip_code	  =pr("zip_code")	
	prefix_phone1 =pr("prefix_phone1")
	prefix_phone2 =pr("prefix_phone2")
	prefix_fax1	  =pr("prefix_fax1")
	prefix_fax2	  =pr("prefix_fax2")
	phone1	      =pr("phone1")
	phone2	      =pr("phone2")
	fax1          =pr("fax1")
	fax2	      =pr("fax2")
	company_email =pr("email")
	url	          =pr("url")	
	date_update	  =pr("date_update")			    
    status_company =pr("status")
    company_desc   =pr("company_desc")
    private_flag  = pr("private_flag")  
  End if   
 End if 	          
 
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 9 Order By word_id"				
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
<td class="page_title" width=100% align=center style="line-height:25px" dir="<%=dir_obj_var%>">&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<font color="#6F6DA6"><%=CONTACT_NAME%></font>&nbsp;</td></tr>         	
<tr><td height=1 bgcolor=#808080 nowrap colspan=2></td></tr>
<tr><td height=10 nowrap colspan=2></td></tr>
<tr><td width=640 colspan=2>
<table dir="<%=dir_var%>" cellpadding=2 cellspacing=0 width=630 align=center border=1 BORDERCOLOR=black style="border-collapse:collapse">
  <tr><td colspan=2 class="title" align="<%=align_var%>" bgcolor=#E6E6E6><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;</td></tr>
       <tr>
         <td width="510" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=CONTACT_NAME%>&nbsp;</td>
         <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<span id="word14" name=word14><!--שם מלא--><%=arrTitles(14)%></span>&nbsp;</td>
      </tr>  
      <tr>
         <td width="510" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=messangerName%>&nbsp;</td>                    
         <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<span id="word15" name=word15><!--תפקיד--><%=arrTitles(15)%></span>&nbsp;</td>
      </tr>       
      <tr>
        <td width="510" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=contact_address%>&nbsp;</td>
        <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<!--כתובת--><%=arrTitles(5)%>&nbsp;</td>
     </tr>             
     <tr>
         <td width="510" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=contact_city_Name%>&nbsp;</td>
         <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<!--עיר--><%=arrTitles(6)%>&nbsp;</td>
      </tr>
     <tr>
         <td width="510" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=contact_zip_code%>&nbsp;</td>
         <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<!--מיקוד--><%=arrTitles(58)%>&nbsp;</td>
      </tr>                      
      <tr>
         <td width="510" align="<%=align_var%>" dir="ltr">&nbsp;<%=phone%>&nbsp;</td>
         <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px"><!--טלפון--><%=arrTitles(16)%>&nbsp;</td>
      </tr>
      <tr>
         <td width="510" align="<%=align_var%>" dir="ltr">&nbsp;<%=cellular%>&nbsp;</td>
         <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px"><!--טלפון נייד--><%=arrTitles(17)%>&nbsp;</td>
      </tr>
         <tr>
         <td width="510" align="<%=align_var%>" dir="ltr">&nbsp;<%=phone_additional%>&nbsp;</td>
         <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">טלפון נוסף&nbsp;</td>
      </tr>
   
      <tr>
         <td width="510" align="<%=align_var%>" dir="ltr">&nbsp;<%=fax%>&nbsp;</td>        
         <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px"><!--פקס--><%=arrTitles(18)%>&nbsp;</td>
      </tr>
      <tr>
         <td width="510" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=trim(email)%>&nbsp;</td>
         <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">Email&nbsp;</td>
      </tr>
       <tr>
           <td width="510" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=contact_types%>&nbsp;</td>
           <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px" valign=top>&nbsp;<!--קבוצה--><%=arrTitles(11)%>&nbsp;</td>
      </tr>  
      <tr>
      <td width="510" align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top>&nbsp;<%=breaks(trim(contact_desc))%>&nbsp;</td>
      <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px" valign=top>&nbsp;<!--פרטים נוספים--><%=arrTitles(12)%>&nbsp;</td>
      </tr>         
       <tr>
           <td align="<%=align_var%>">&nbsp;<%=responsible_name%>&nbsp;</td>
           <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px" valign=top>&nbsp;<!--אחראי--><%=arrTitles(59)%>&nbsp;</td>
      </tr> 
    </table></td></tr>  
	<tr><td height=10 nowrap colspan=2></td></tr>
	<%If trim(private_flag) = "0" Then%>
	<tr><td width=640 colspan=2 dir="<%=dir_var%>">
	<table dir="<%=dir_var%>" cellpadding=2 cellspacing=0 width=630 align=center border=1 BORDERCOLOR=black style="border-collapse:collapse">       
	<tr><td colspan=6 class="title" align="<%=align_var%>" bgcolor=#E6E6E6><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;</td></tr>
    <tr>
         <td width="510" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=company_name%>&nbsp;</td>
         <td width="120" nowrap align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<span id=word4 name=word4><!--שם--><%=arrTitles(4)%></span>&nbsp;</td> 
      </tr>            
      <tr>
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=address%>&nbsp;</td>
        <td align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<span id="word5" name=word5><!--כתובת--><%=arrTitles(5)%></span>&nbsp;</td>
     </tr>             
     <tr>
         <td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>">&nbsp;<%=cityName%>&nbsp;</td>
         <td align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<span id="word6" name=word6><!--עיר--><%=arrTitles(6)%></span>&nbsp;</td>
      </tr>
     <tr>
         <td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>">&nbsp;<%=zip_code%>&nbsp;</td>
         <td align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<!--מיקוד--><%=arrTitles(58)%>&nbsp;</td>
      </tr>                              
      <tr>
          <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=phone1%>&nbsp;</td>
          <td align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<span id="word7" name=word7><!--1 טלפון--><%=arrTitles(7)%></span>&nbsp;</td>
      </tr>                                
      <tr>  
          <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=phone2%>&nbsp;</td>
          <td align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<span id="word8" name=word8><!--2 טלפון--><%=arrTitles(8)%></span>&nbsp;</td>
      </tr>  
      <tr>  
          <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=fax1%>&nbsp;</td>
          <td align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<span id="word9" name=word9><!--פקס--><%=arrTitles(9)%></span>&nbsp;</td>
      </tr>                          
      <tr>
         <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=company_email%>&nbsp;</td>
         <td align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;Email&nbsp;</td>
      </tr>       
       <tr>
         <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=url%>&nbsp;</td>
         <td align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px">&nbsp;<span id="word10" name=word10><!--אתר--><%=arrTitles(10)%></span>&nbsp;</td>
      </tr>                
      <tr>
           <td align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=comp_desc%>&nbsp;</td>
           <td align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px" valign=top>&nbsp;<span id="word12" name=word12><!--פרטים נוספים--><%=arrTitles(12)%></span>&nbsp;</td>
      </tr>     
      <tr>
           <td align="<%=align_var%>">
           &nbsp;<%=arr_status_comp(status_company)%>&nbsp;
           </td>
           <td align="<%=align_var%>" class=title style="padding-<%=align_var%>:15px" valign=top>&nbsp;<span id="word13" name=word13><!--סטטוס--><%=arrTitles(13)%></span>&nbsp;</td>
      </tr>
   </table></td></tr>              
   <%End If%>
<tr><td height=10 nowrap colspan=2></td></tr>
</table>
</body>
<%set con=Nothing%>
</html>


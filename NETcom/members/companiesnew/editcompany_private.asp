<% server.ScriptTimeout=10000 %>
<% Response.Buffer = True %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%   
  sort = Request.QueryString("sort")   
  contactID = trim(Request("contactID"))
  companyID = trim(Request("companyID"))
   if trim(lang_id) = "1" then
	   arr_status_comp = Array("","עתידי","פעיל","סגור","פונה")
   else
	   arr_status_comp = Array("","new","active","close","appeal")
   end if	  
  
  If trim(companyID)<>"" then   
  sqlStr = "SELECT company_name,address,city_Name"&_
  ", prefix_phone,phone,prefix_fax,fax"&_
  ", date_update,status FROM companies_private_view WHERE company_id="& companyID 
  'Response.Write "test"
  'Response.End
  set pr=con.GetRecordSet(sqlStr)
  if not pr.EOF then	
	company_name   =pr("company_name")  
	address	       =pr("address")	
	cityName	   =pr("city_Name")
	prefix_phone   =pr("prefix_phone")	
	prefix_fax	   =pr("prefix_fax")	
	phone	       =pr("phone")	
	fax            =pr("fax")			
	date_update	   =pr("date_update")
	status_company =pr("status")			    
    If isNumeric(trim(companyID)) Then
		company_types=""		
		sqlstr="Select type_Id From company_to_types_view Where company_id = " & companyID & " Order By type_id"
		set rssub = con.getRecordSet(sqlstr)		   			
		If not rssub.eof Then
			company_types = rssub.getString(,,",",",")	
		End If
		set rssub=Nothing				
		
		sqlStr = "SELECT company_desc FROM companies WHERE company_id="& companyID 
		'Response.Write "test"
		'Response.End
		set rs_d=con.GetRecordSet(sqlStr)
		If not rs_d.EOF then				
  			company_desc = trim(rs_d(0))
		End if 
		set rs_d = Nothing
   End If    
  end if  
 Else
	status_company = 2  
 End if 
 
 If trim(contactID)<>"" then 
   sql="SELECT contact_ID,email,prefix_phone,phone,prefix_cellular,prefix_fax,fax,"&_
   "cellular,CONTACT_NAME,messanger_name,date_update FROM contacts"&_
   " WHERE contact_ID=" & contactID
   set listContact=con.GetRecordSet(sql)
   if not listContact.EOF then 
      contactID=listContact("contact_ID")      
      CONTACT_NAME=listContact("CONTACT_NAME")      
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
  
   
  delContactID=trim(Request("delContactID"))  
  If delContactID<>nil And delContactID<>"" Then 	
	con.ExecuteQuery "delete from Contacts WHERE Contact_Id=" & delContactID	
	con.executeQuery "delete From tasks WHERE task_id IN (Select task_id FROM activities Where contact_id=" &  delContactID & ")"' delete contact also from activities	
	con.executeQuery "delete From activities Where contact_id=" &  delContactID ' delete contact also from activities	
    Response.Redirect "contact.asp?companyId=" & companyId               
  End If   
  
  delActivityID = trim(Request("delActivityID"))
  If delActivityID<>nil And delActivityID<>"" Then 			
	con.executeQuery "delete From tasks WHERE task_id IN (Select task_id FROM activities Where activity_id=" &  delActivityID & ")"' delete contact also from activities	
	con.executeQuery "delete From activities Where activity_id=" &  delActivityID ' delete contact also from activities	
    Response.Redirect "contact.asp?companyID=" & companyId & "&contactID=" & contactID              
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
	  
	  sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	  set rsbuttons = con.getRecordSet(sqlstr)
	  If not rsbuttons.eof Then
		arrButtons = ";" & rsbuttons.getString(,,",",";")		
		arrButtons = Split(arrButtons, ";")
	  End If
	  set rsbuttons=nothing	  
%>
<html>
<head>
<!--#include file="../../../include/title_meta_inc.asp"-->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">

function CheckFields(act)
{   
		if (window.document.all("CONTACT_NAME").value=='')
		{
		   <%
				If trim(lang_id) = "1" Then
					str_alert = "נא למלא שם"
				Else
					str_alert = "Please, insert the full name"
				End If	
			%>
 			window.alert("<%=str_alert%>");
 			window.document.all("CONTACT_NAME").focus();
			return false;
		}
		/*if (document.all("type_id").options[document.all("type_id").selectedIndex].value=='null')
		{
		   <%
				If trim(lang_id) = "1" Then
					str_alert = "נא לבחור קבוצה"
				Else
					str_alert = "Please, choose the group"
				End If	
			%>
				window.alert("<%=str_alert%>");
				document.all("type_id").focus();
				return false;
		}*/			
		if (window.document.all("email").value=='')
		{
			<%
				If trim(lang_id) = "1" Then
					str_confirm = Space(72) & ",שים לב\n\n ,לא  הזנת מייל ולכן " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & " לא יכלל ברשימות הדיוור\n\n" & Space(47) & "? האם ברצונך להמשיך"
				Else
					str_confirm = "Pay attention !"&Space(80)&" \n\n You didn't insert the email, so the " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & " will not be included in the mailing list. \n\n Are you sure want to proceed ?"& Space(55)
				End If	
			%>
			answer = window.confirm("<%=str_confirm%>");
			if(answer == false)
			{
				window.document.all("email").focus();
				return false;
			}	
		}
		else if (!checkEmail(document.all("email").value))
		{
			<%
				If trim(lang_id) = "1" Then
					str_alert = "כתובת דואר אלקטרוני לא חוקית"
				Else
					str_alert = "The email address is not valid!"
				End If	
			%>
				window.alert("<%=str_alert%>");
				document.all("email").focus();
				return false;
		}  		
		
		window.opener.name = "contact";                 
		document.form_contact.target = "contact";
		document.form_contact.submit();        
		window.close();
		return false; 
}

function checkEmail(addr)
{
	if (addr == '') {
	return false;
	}
	var invalidChars = '\/\'\\ ";:?!()[]\{\}^|';
	for (i=0; i<invalidChars.length; i++) {
	if (addr.indexOf(invalidChars.charAt(i),0) > -1) {
		return false;
	}
	}
	for (i=0; i<addr.length; i++) {
	if (addr.charCodeAt(i)>127) {     
		return false;
	}
	}

	var atPos = addr.indexOf('@',0);
	if (atPos == -1) {
	return false;
	}
	if (atPos == 0) {
	return false;
	}
	if (addr.indexOf('@', atPos + 1) > - 1) {
	return false;
	}
	if (addr.indexOf('.', atPos) == -1) {
	return false;
	}
	if (addr.indexOf('@.',0) != -1) {
	return false;
	}
	if (addr.indexOf('.@',0) != -1){
	return false;
	}
	if (addr.indexOf('..',0) != -1) {
	return false;
	}
	var suffix = addr.substring(addr.lastIndexOf('.')+1);
	if (suffix.length != 2 && suffix != 'com' && suffix != 'net' && suffix != 'org' && suffix != 'edu' && suffix != 'int' && suffix != 'mil' && suffix != 'gov' & suffix != 'arpa' && suffix != 'biz' && suffix != 'aero' && suffix != 'name' && suffix != 'coop' && suffix != 'info' && suffix != 'pro' && suffix != 'museum') {
	return false;
	}
return true;
}
//-->
</script>  
</head>
<body style="margin:0;background:'#e6e6e6'">
<table border="0" bordercolor="navy" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top">
<tr><td class="page_title">&nbsp;<%if trim(companyId)<>"" then%><font style="color:#000000;"><b><span id=word1 name=word1><!--עדכן--><%=arrTitles(1)%></span> <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></b></font><%elseIf trim(companyId) = "" Then%><font style="color:#000000;"><b><span id="word2" name=word2><!--הוסף--><%=arrTitles(2)%></span> <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></b><%end if%>&nbsp;</td></tr>         	
<tr><td width="100%">
<table cellpadding=0 bgcolor="#e6e6e6" cellspacing=0 width=100% border=0>
<FORM name="form_contact" ACTION="addcontact.asp?contactID=<%=contactID%>" METHOD="post" ID="form_contact">   
<tr><td height=10 nowrap></td></tr>
<tr>       		 		                  
  <td width=100% style="padding-<%=align_var%>:10px;">  
  <input type="hidden" name="companyID" id="companyID" value="<%=companyID%>">  
  <input type="hidden" name="CONTACT_ID" id="CONTACT_ID" value="<%=contactID%>"> 
  <input type="hidden" name="private" id="private" value="1">     
   <table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">   
      <tr>
         <td width="75%" align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">           
         <input type=text dir="<%=dir_obj_var%>" name="contact_name" value="<%=vFix(CONTACT_NAME)%>" maxlength=50 style="width:200" class="Form">           
         </td>
         <td width="25%" align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<span id="word14" name=word14><!--שם מלא--><%=arrTitles(14)%></span>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
      </tr>     
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">                        
         <a class="but_menu" name=messangers_list id="messangers_list" href="#" onclick='window.open("messangers_list.asp","MessangersList","left=300,top=150,width=200,height=60")' style="width:120"><!--בחר מהרשימה--><%=arrTitles(19)%></a>&nbsp;
         <input type=text dir="<%=dir_obj_var%>" name="messangerName" value="<%=vFix(messangerName)%>" style="width:200" maxlength=50 class="Form" ID="messangerName">         
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<span id="word15" name=word15><!--תפקיד--><%=arrTitles(15)%></span>&nbsp;</td>
      </tr> 
      <tr>
        <td align="<%=align_var%>"><input type=text dir="<%=dir_obj_var%>" name="address" value="<%=vFix(address)%>"  style="width:200" class="Form" ></td>
        <td align="<%=align_var%>">&nbsp;<span id="word5" name=word5><!--כתובת--><%=arrTitles(5)%></span>&nbsp;</td>
     </tr>             
      <tr>
         <td align="<%=align_var%>">
         <a class="but_menu" href="#" onclick='window.open("cities_list.asp","CitiesList","left=100,top=200,width=200,height=60")' style="width:120"><span id="word19" name=word19><!--בחר מהרשימה--><%=arrTitles(19)%></span></a>&nbsp;
         <input type=text dir="<%=dir_obj_var%>" name="cityName" value="<%=vFix(cityName)%>"  style="width:200" class="Form" ID="cityName">         
         </td>
         <td align="<%=align_var%>">&nbsp;<span id="word6" name=word6><!--עיר--><%=arrTitles(6)%></span>&nbsp;</td>
      </tr>            
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" dir=ltr>             
             <input type=text dir="ltr" name="phone" value="<%=vFix(phone)%>" size="20" maxlength=20 class="Form">                
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<span id="word16" name=word16><!--טלפון--><%=arrTitles(16)%></span>&nbsp;</td>
      </tr>
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" dir=ltr>             
          <input type=text dir="ltr" name="cellular" value="<%=vFix(cellular)%>" size="20" maxlength=20 class="Form">              
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<span id="word17" name=word17><!--טלפון נייד--><%=arrTitles(17)%></span>&nbsp;</td>
      </tr>
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" dir=ltr>                   
         <input type=text dir="ltr" name="fax" value="<%=vFix(fax)%>" size="20" maxlength=20 class="Form"></td>           
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<span id="word18" name=word18><!--פקס--><%=arrTitles(18)%></span>&nbsp;</td>
      </tr>
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">                        
               <input type=text name="email" dir=ltr value="<%=vFix(email)%>" size="30" class="Form">              
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;Email&nbsp;</td>
      </tr>     
     <tr>
      <td align="<%=align_var%>">
      <textarea dir="<%=dir_obj_var%>" class=Form name="company_desc" style="width:330" rows=6 ID="company_desc"><%=trim(company_desc)%></textarea>
      </td>
      <td align="<%=align_var%>" valign=top>&nbsp;<span id="word12" name=word12><!--פרטים נוספים--><%=arrTitles(12)%></span>&nbsp;</td>
     </tr> 
      <tr>
      <td align="<%=align_var%>">
		    <select id="status" name="status" dir="<%=dir_obj_var%>" class="norm">
			  <%For i=1 To Ubound(arr_status_comp)%>
			   <option value="<%=i%>" <%If trim(status_company) = trim(i) Then%>selected<%End If%>>&nbsp;<%=arr_status_comp(i)%>&nbsp;</option>
			  <%Next%> 
			</select>
		 </td>
	 <td align="<%=align_var%>" valign=top>&nbsp;<span id="word13" name=word13><!--סטטוס--><%=arrTitles(13)%></span>&nbsp;</td>	 
	</tr>	 
	 </form>	    
</table>
</td>
</tr>
<tr><td colspan="3" height="15" nowrap></td></tr> 
<tr><td colspan="3" align=center dir="<%=dir_var%>">
<table cellpadding=0 cellspacing=0 width=80% align=center>
<tr><td width=50% nowrap align=center>
<input type=button class="but_menu" value="<%=arrButtons(2)%>" onClick="window.close();" style="width:80" ID="Button2" NAME="Button2">
</td>
<td width=50% nowrap align=center>
<input type=button class="but_menu" value="<%=arrButtons(1)%>" onClick="return CheckFields('submit')" style="width:80" ID="Button1" NAME="Button1">
</td></tr>
</table>
</td></tr>
</table>
</body>
<%set con=Nothing%>
</html>


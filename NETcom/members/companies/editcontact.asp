<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%   
 
  contactID = trim(Request("contactID"))
  companyID = trim(Request("companyID"))
  
  if Request.Form("CONTACT_NAME")<>nil then     
    
   CONTACT_NAME=sFix(trim(Request.Form("CONTACT_NAME")))     
   prefix_phone=sFix(trim(Request.Form("prefix_phone")))	
   phone=sFix(trim(Request.Form("phone")))
   prefix_fax=sFix(trim(Request.Form("prefix_fax")))
   fax=sFix(trim(Request.Form("fax")))
   prefix_cellular=sFix(trim(Request.Form("prefix_cellular")))
   cellular=sFix(trim(Request.Form("cellular")))
   email=sFix(trim(Request.Form("email")))   
   messangerName=sFix(trim(Request.Form("messangerName")))
   extension_phone = trim(Request.Form("extension_phone"))
   contact_address=trim(sFix(Request.Form("contact_address")))
   contact_city_name=trim(sFix(Request.Form("contact_city_name")))
   contact_zip_code=trim(sFix(Request.Form("contact_zip_code")))
   contact_desc = sFix(Left(trim(Request.Form("contact_desc")),500))
   contact_types = trim(Request.Form("contact_types"))
   
   if trim(contactID)<>"" then   		
      sqlUpd="UPDATE contacts SET CONTACT_NAME='"& CONTACT_NAME &_
      "', email='"& email  &"',phone='"& phone &"', fax='"& fax &_
      "', cellular='"& cellular &"', date_update=getDate()" &_
      ", messanger_name='"& messangerName &"', contact_address = '" & contact_address &_
      "', contact_city_name = '" & contact_city_name & "', contact_zip_code = '" & contact_zip_code &_    
      "', contact_desc = '" & contact_desc &_    
      "' WHERE contact_ID="& contactID 
      'Response.Write(sqlUpd)
      'Response.End
      con.ExecuteQuery(sqlUpd)
      
		'--insert into changes table
		sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
		" SELECT 'איש קשר', ' שם:'  + IsNULL(CONTACT_NAME, '') + ' נייד: ' + IsNULL(cellular, ''), CONTACT_ID, 'עדכון', getDate()," & UserID & _
		" FROM dbo.CONTACTS WHERE (CONTACT_ID = "& contactID &")"
		con.executeQuery(sqlstr)
      
   else               
      sqlIns="SET DATEFORMAT MDY; SET NOCOUNT ON; INSERT INTO contacts("&_
      " company_ID,CONTACT_NAME,email,phone,fax,cellular,date_update,messanger_name,"&_
      " contact_address,contact_city_name,contact_zip_code,contact_desc,Organization_ID) VALUES ("&_
      companyID & ", '"& CONTACT_NAME & "','" & email & "','" & phone &_
      "','"& fax & "','" & cellular & "', getDate(),'" & messangerName & "','" & contact_address & "','" &_
      contact_city_name & "','" & contact_zip_code & "','" & contact_desc & "'," & OrgID & ");"&_
      " SELECT @@IDENTITY AS NewID"
      'Response.Write(sqlIns)
      'Response.End
     set rs_tmp = con.getRecordSet(sqlIns)
		contactID = rs_tmp.Fields("NewID").value
	 set rs_tmp = Nothing	
	 
    '--insert into changes table
    sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
    " SELECT 'איש קשר', ' שם:'  + IsNULL(CONTACT_NAME, '') + ' נייד: ' + IsNULL(cellular, ''), CONTACT_ID, 'הוספה', getDate()," & UserID & _
    " FROM dbo.CONTACTS WHERE (CONTACT_ID = "& contactID &")"
     con.executeQuery(sqlstr)	 
    End if     
    
    ' עדכון סיווגי אנשי הקשר
    sqlstr="Delete FROM contact_to_types WHERE contact_ID = " & contactID
    con.executeQuery(sqlstr)	
	
	If  Len(contact_types) > 0 Then
		arrTypes = Split(contact_types & ",", ",")
		numOfTypes = Ubound(arrTypes)
	End If	
	
	If IsArray(arrTypes) And numOfTypes > 0 Then
		For i=0 To numOfTypes		
			If IsNumeric(arrTypes(i)) Then	
			sqlstr="Insert Into contact_to_types (contact_ID,type_id) values ("&contactID&","&arrTypes(i)&" )"			
			con.executeQuery(sqlstr)	
			End If		
		Next
	End If
	
If trim(wizard_id) <> "" Then%>
   <SCRIPT LANGUAGE=javascript>
   <!--
        window.opener.document.location.href= "contact_wizard.asp?companyId=<%=companyID%>&contactID=<%=contactID%>";
        self.close();
   //-->
  </SCRIPT>
<%Else%>
   <SCRIPT LANGUAGE=javascript>
   <!--
       window.opener.document.location.href="contact.asp?companyId=<%=companyID%>&contactID=<%=contactID%>";
       self.close();
   //-->
  </SCRIPT>
<%End If%> 
<%End If%>  
<% If trim(contactID)<>"" then 
   sql="SELECT contact_ID,email,prefix_phone,phone,prefix_cellular,prefix_fax,fax,"&_
   " cellular,CONTACT_NAME,messanger_name,date_update,contact_city_name, contact_address,"&_
   " contact_desc FROM contacts WHERE contact_ID=" & contactID
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
      contact_address = listContact("contact_address")
      contact_zip_code = listContact("contact_zip_code")
      contact_city_Name = listContact("contact_city_Name")
      contact_desc = listContact("contact_desc")
	  If Len(contact_desc) > 70 Then
  		  contact_desc_short = Left(contact_desc,70) & ".."
  	  Else
  		  contact_desc_short = contact_desc
  	  End If            
    End if   
    set  listContact = Nothing
  End If       
  
	If isNumeric(trim(contactId)) Then
		types=""
		sqlstr = "get_contact_types '" & contactId & "'"
		set rstype = con.getRecordSet(sqlstr)		   						
		If not rstype.eof Then
			types = rstype.getString(,,",",",")						
		End If		    
		set rstype=Nothing		
		If Len(types) > 0 Then
			types = Left(types,(Len(types)-1))
		End If
		
		contact_types=""
		sqlstr = "get_contact_types_id '" & contactId & "'"
		set rstype = con.getRecordSet(sqlstr)		   						
		If not rstype.eof Then
			contact_types = rstype.getString(,,",",",")						
		End If		    
		set rstype=Nothing		
		If Len(contact_types) > 0 Then
			contact_types = Left(contact_types,(Len(contact_types)-1))
		End If		
	End If	   

	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 12 Order By word_id"				
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
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">

var DetailsList,ContactList
var changed = false; 

function fieldChanged() { 
     changed = true; 
} 
var oPopup = window.createPopup();
function activityDropDown(obj)
{
    oPopup.document.body.innerHTML = activity_Popup.innerHTML; 
    oPopup.document.charset="windows-1255";
    oPopup.show(-206, 17, 220, 84, obj);    
}

function CheckFields(subButton)
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
		if (window.document.all("email").value=='')
		{
			<%
				If trim(lang_id) = "1" Then
					str_confirm = Space(80) & ",שים לב\n\n ,לא  הזנת מייל ולכן " & trim(Request.Cookies("bizpegasus")("ContactsOne")) & " לא יכלל ברשימות הדיוור\n\n" & Space(55) & "? האם ברצונך להמשיך"
				Else
					str_confirm = "Pay attention !"&Space(80)&" \n\n You didn't insert the email, so the " & trim(Request.Cookies("bizpegasus")("ContactsOne")) & " will not be included in the mailing list. \n\n Are you sure want to proceed ?"& Space(55)
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
			  
 
    subButton.disabled = true;
    //window.opener.name = "company";                 
    //document.form_contact.target = "company";
    document.form_contact.submit();        
    //window.close();
   return false; 
}

function checkEmail(addr)
{
	addr = addr.replace(/^\s*|\s*$/g,"");
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
<body style="margin:0px;background:#E5E5E5">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top">
<tr><td class="page_title">&nbsp;<%If trim(contactID)<>"" Then%><span id=word1 name=word1><!--עדכן--><%=arrTitles(1)%></span>&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%><%Else%><span id="word2" name=word2><!--הוסף--><%=arrTitles(2)%></span> <%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%><%End If%>&nbsp;</td></tr>         	
<tr><td height="10" nowrap></td></tr>
<tr>
  <td align="<%=align_var%>" valign="top" style="padding-right:10px; padding-left:10px">
  <FORM name="form_contact" ACTION="editcontact.asp" METHOD="post" ID="form_contact">   
   <input type="hidden" name="companyID" id="companyID" value="<%=companyID%>">
   <input type="hidden" name="contactID" id="contactID" value="<%=contactID%>">
    <table width="100%" border="0" cellpadding="2" cellspacing="0" dir="<%=dir_var%>">	        
      <tr>
         <td width="100%" align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">           
         <input type=text dir="<%=dir_obj_var%>" name="CONTACT_NAME" value="<%=vFix(CONTACT_NAME)%>" maxlength=50 style="width:180" class="Form">           
         </td>
         <td width="100" nowrap align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<span id="word14" name=word14><!--שם מלא--><%=arrTitles(14)%></span>&nbsp;<font style="color:#FF0000;font-size:11px">*</font></td>
      </tr>     
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">		 
         <INPUT type=image src="../../images/icon_find.gif" name=messangers_list id="messangers_list" onclick='window.open("messangers_list.asp","MessangersList","left=300,top=100,width=250,height=350,scrollbars=1");return false;' title="<%=arrTitles(19)%>">&nbsp;
         <input type=text dir="<%=dir_obj_var%>" name="messangerName" value="<%=vFix(messangerName)%>" style="width:180" maxlength=50 class="Form" ID="messangerName">         
         </td>         
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<span id="word15" name=word15><!--תפקיד--><%=arrTitles(15)%></span>&nbsp;</td>
      </tr>
       <tr>
        <td align="<%=align_var%>"><input type=text dir="<%=dir_obj_var%>" name="contact_address" value="<%=vFix(contact_address)%>"  style="width:180" class="Form" ID="contact_address"></td>
        <td align="<%=align_var%>">&nbsp;<span id="word5" name=word5><!--כתובת--><%=arrTitles(5)%></span>&nbsp;</td>
     </tr>             
     <tr>
         <td align="<%=align_var%>" nowrap>
         <INPUT type=image src="../../images/icon_find.gif" onclick='window.open("cities_list.asp?fieldID=contact_city_Name","CitiesList","left=100,top=50,width=450,height=450,scrollbars=1");return false;' title="<%=arrTitles(19)%>">&nbsp;                    
         <input type=text dir="<%=dir_obj_var%>" class="Form" name="contact_city_Name" id="contact_city_Name" value="<%=contact_city_Name%>" style="width:180">
         </td>
         <td align="<%=align_var%>">&nbsp;<span id="Span3" name=word6><!--עיר--><%=arrTitles(6)%></span>&nbsp;</td>
      </tr>         
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" dir=ltr>             
         <input type=text dir="ltr" name="contact_zip_code" value="<%=vFix(contact_zip_code)%>" size="20" maxlength=20 class="Form" ID="contact_zip_code">
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<!--מיקוד--><%=arrTitles(16)%>&nbsp;</td>
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
         <input type=text name="email" dir=ltr value="<%=vFix(email)%>" style="width:180" maxlength=50 class="Form">              
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;Email&nbsp;</td>
      </tr> 
      <tr>
           <td align="<%=align_var%>">
           <INPUT type=image src="../../images/icon_find.gif" name=types_list id="types_list" onclick='window.open("types_list.asp?contactID=<%=contactID%>&contact_types=" + contact_types.value,"TypesList","left=300,top=100,width=250,height=350,scrollbars=1");return false;' title="<%=arrTitles(19)%>">&nbsp;
           <span class="Form_R" dir="<%=dir_obj_var%>" style="width:180;line-height:16px" name="types_names" id="types_names"><%=types%></span>
           <input type=hidden name=contact_types id=contact_types value="<%=contact_types%>">
           </td>
           <td align="<%=align_var%>" valign=top>&nbsp;<span id="Span1" name=word11><!--קבוצה--><%=arrTitles(11)%></span>&nbsp;</td>
      </tr>  
      <tr>
      <td align="<%=align_var%>" valign=top>
      <textarea class="Form" dir="<%=dir_obj_var%>" style="width:280;" rows=5 name="contact_desc" id="contact_desc"><%=trim(contact_desc_short)%></textarea>      
      </td>
      <td align="<%=align_var%>" valign=top>&nbsp;<span id="Span4" name=word12><!--פרטים נוספים--><%=arrTitles(12)%></span>&nbsp;</td>
      </tr>     
      <tr><td colspan=2 height="20" nowrap></td></tr>
	  <tr><td colspan="2" align=center dir=<%=dir_var%>>
	  <input type=button class="but_menu" value="<%=arrButtons(2)%>" onClick="window.close();" style="width:80" ID="Button2" NAME="Button2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <input type=button class="but_menu" value="<%=arrButtons(1)%>" onClick="return CheckFields(this)" style="width:80" ID="Button1" NAME="Button1">
      </td></tr>
    </FORM>          
</table>
</body>
<%set con=Nothing%>
</html>


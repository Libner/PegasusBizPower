<!--#include file="../../connect.asp"-->
<% server.ScriptTimeout=10000 %>
<% Response.Buffer = True %>
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%sort = trim(Request.QueryString("sort"))
	ContactId = trim(Request("ContactId"))
	companyID = trim(Request("companyID"))
	LinkedContactid=trim(Request("LinkedContactid"))  
	ContactLinksType=Request.Form("ContactLinksType")
	
'	response.Write LinkedContactid &":"& Request.Form("ContactLinksType")
'response.end

 
   if Request.Form("company_name") <> nil then  
	
	company_name = trim(sFix(Request.Form("company_name")))
	address = trim(sFix(Request.Form("address")))
	email = trim(sFix(Request.Form("email")))
	cityName = trim(sFix(Request.Form("cityName")))
	zip_code = trim(sFix(Request.Form("zip_code")))
	url = trim(sFix(Request.Form("url")))
	phone1 = trim(sFix(Request.Form("phone1")))
	phone2 = trim(sFix(Request.Form("phone2")))
	fax1 = trim(sFix(Request.Form("fax1")))
	fax2 = trim(sFix(Request.Form("fax2")))		
	company_email = trim(sFix(Request.Form("company_email")))
	status = trim(Request.Form("status"))
	If trim(status) = "" Then
		status = "2"
	End If		
	company_desc = sFix(Left(trim(Request.Form("company_desc")),500))
	

   if trim(companyID)<>"" then
      sqlUpd="UPDATE companies SET company_name='" & company_name &_
      "', date_update=getDate() , company_desc='" & company_desc & "'" &_
      ", address='"& address & "', city_Name='" & cityName & "', zip_code='" & zip_code &_  
      "', url='" & url & "'" & ", prefix_phone1='"& prefix_phone1 & "', prefix_phone2='"& prefix_phone2 &_
      "', prefix_fax1='"& prefix_fax1 &"', prefix_fax2='"& prefix_fax2 &"'" &_
      ", phone1='"& phone1 &"', phone2='"& phone2 &"', fax1='"& fax1 &_
      "', fax2='"& fax2 &"', email='"& company_email &"', User_Id=" & UserId &_   
      "  WHERE company_ID="& companyID
     'Response.Write(sqlUpd)
     'Response.End
      con.ExecuteQuery(sqlUpd)
   else      
      sqlIns="SET NOCOUNT ON; INSERT INTO COMPANIES (company_name,date_update,private_flag,company_desc"&_      
      ", address,url,prefix_phone1,prefix_phone2,prefix_fax1,prefix_fax2"&_
      ", phone1,phone2,fax1,fax2,zip_code,city_Name,email,status,User_Id,Organization_ID "&_      
      ")  VALUES ( '"& company_name &"',getDate(), 0,'"&company_desc & "','" &_      
      address &"','"& url &"','"& prefix_phone1 &"','"& prefix_phone2 &"','"&_
      prefix_fax1 &"','"& prefix_fax2 &"','"& phone1 &"','"& phone2 &"','"&_
      fax1 &"','"& fax2 &"','"& zip_code &"','"& cityName &"','" & company_email & "','" & status & "',"&_
      UserID & "," & OrgID & "); SELECT @@IDENTITY AS NewID"
      'Response.Write(sqlIns)
      'Response.End     
	  set rs_tmp = con.getRecordSet(sqlIns)
		companyID = rs_tmp.Fields("NewID").value
	  set rs_tmp = Nothing	  
	end if	
    '����� ��� ��� ��� ���� - ������ ���� �� ���� ��� ���� ������
	ElseIf Request.Form("companyID")=nil And Request.Form("company_name") = nil and Request.Form("CONTACT_NAME")<>nil Then 
		sqlstr = "Select company_id From companies WHERE Organization_ID = " & orgID & " And private_flag = 1"
		set rs_private = con.getRecordSet(sqlstr)
		if not rs_private.eof then		
			companyID = rs_private(0)	
		end if
		set rs_private = Nothing	
	End If		
' REM: add User ====================================================
if Request.Form("CONTACT_NAME")<>nil then     
     if IsNumeric(LinkedContactid) and Request.Form("ContactLinksType")<>"" then
 sqlDel="delete from ContactLinks where  (Contact_Id="&LinkedContactid &" and Contact_Linked="& contactId &") or (Contact_Id="& contactId &" and Contact_Linked="&  LinkedContactid &")"
 con.ExecuteQuery(sqlDel)
 sqlInsert="Insert into ContactLinks (Contact_Id,Contact_Linked,Contact_Type) values("& LinkedContactid &","& contactId &","& Request.Form("ContactLinksType") &")"
 con.ExecuteQuery(sqlInsert)

 end if
   contact_name=sFix(trim(Request.Form("contact_name")))
   first_name_E=sFix(trim(Request.Form("first_name_E")))
   last_name_E =sFix(trim(Request.Form("last_name_E")))
   phone=sFix(trim(Request.Form("phone")))
   fax=sFix(trim(Request.Form("fax")))
   cellular=sFix(trim(Request.Form("cellular")))
   phone_additional=sFix(trim(Request.Form("phone_additional")))
   email=sFix(trim(Request.Form("email")))   
   messangerName=sFix(trim(Request.Form("messangerName")))
   extension_phone = trim(Request.Form("extension_phone"))
   contact_address=trim(sFix(Request.Form("contact_address")))
   contact_city_name=trim(sFix(Request.Form("contact_city_name")))
   contact_zip_code=trim(sFix(Request.Form("contact_zip_code")))
   contact_desc = sFix(Left(trim(Request.Form("contact_desc")),500))
   contact_types = trim(Request.Form("contact_types"))
   responsible_id = trim(Request.Form("responsible_id"))
   is_NewLetters= Request.Form("is_NewLetters")
   if is_NewLetters="on" then
   is_NewLetters=1
   else
   is_NewLetters=0
   end if
   if ContactId="" and  trim(Request.Cookies("bizpegasus")("Edit_IsNewLetters"))="0" then
   is_NewLetters=1
   end if
   
    '===Added by Mila 25/10/2018 for internet tour order from site pegasus
    '==== Inserrt new contact - by default - true.
    is_NotPermit_TourOrder=trim(Request.Form("is_NotPermit_TourOrder"))
   if is_NotPermit_TourOrder="on" or ContactId="" then
        is_NotPermit_TourOrder=1
   else
        is_NotPermit_TourOrder=0
   end if

   if trim(ContactId)<>"" then  
          set st=con.GetRecordSet("select is_NewLetters from CONTACTS where (Contact_Id="& contactId &")")
     if not  st.Eof then
     is_NewLetters_old=trim(st("is_NewLetters"))
     end if
     set st=Nothing
     if is_NewLetters_old="True" then is_NewLetters_old="1"
      if is_NewLetters_old="False" then is_NewLetters_old="0"
     
		'--insert into changes table
		'sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
		'" SELECT '��� ���', ' ��:'  + IsNULL(CONTACT_NAME, '') + ' ����: ' + IsNULL(cellular, ''), CONTACT_ID, '�����', getDate()," & UserID & _
		'" FROM dbo.CONTACTS WHERE (CONTACT_ID = "& contactID &")"
		'con.executeQuery(sqlstr) 
		'response.Write is_NewLetters_old & "::" & is_NewLetters
		'response.end
		if cInt(is_NewLetters_old)=cInt(is_NewLetters) then
'	response.Write is_NewLetters_old & "::" & is_NewLetters
'	response.end
		else
'		response.Write "not="
'	response.Write is_NewLetters_old & "::" & is_NewLetters
'	response.end
   if  trim(Request.Cookies("bizpegasus")("Edit_IsNewLetters"))="1" then
      		sqlstr = "INSERT INTO ChangesNewsLetters (Object_Title,  Change_Date, [User_ID],NewsLetter_Value,Contact_Id,Change_Type) " & _
		" SELECT   IsNULL(CONTACT_NAME, '') + ' ' + IsNULL(cellular, ''), getDate()," & UserID & _
		"," &  is_NewLetters & "," & contactID &",'�����'" & _
		" FROM dbo.CONTACTS WHERE (CONTACT_ID = "& contactID &")"
	'response.Write "sqlstr="& sqlstr
	'	response.end
		con.executeQuery(sqlstr) 
		'response.Write "="
	end if	
      end if
    '  response.end
   else
      sqlIns="SET DATEFORMAT MDY; SET NOCOUNT ON; INSERT INTO contacts (company_ID, CONTACT_NAME,"&_
      " first_name_E, last_name_E, email, phone, fax, cellular, phone_additional,date_update, messanger_name, contact_address, contact_city_name,"&_
      " contact_zip_code, contact_desc, responsible_id, User_Id, Organization_ID,is_NewLetters,is_NewLetters_DateUpdate) VALUES ("&_
      companyID & ",'" & CONTACT_NAME & "','" & first_name_E & "','" & last_name_E & "','" & email & "','" & phone &_
      "','"& fax & "','" & cellular & "','" & phone_additional & "', getDate(),'" & messangerName & "','" & contact_address & "','" &_
      contact_city_name & "','" & contact_zip_code & "','" & contact_desc & "','" & responsible_id & "'," & UserID & "," & OrgID & "," & is_NewLetters &",getDate());"&_
      " SELECT @@IDENTITY AS NewID"
   'Response.Write(sqlIns)
   'Response.End
     set rs_tmp = con.getRecordSet(sqlIns)
		ContactId = rs_tmp.Fields("NewID").value
	 set rs_tmp = Nothing	
	 
    '--insert into changes table
  '  sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
   ' " SELECT '��� ���', ' ��:'  + IsNULL(CONTACT_NAME, '') + ' ����: ' + IsNULL(cellular, ''), CONTACT_ID, '�����', getDate()," & UserID & _
   ' " FROM dbo.CONTACTS WHERE (CONTACT_ID = "& contactID &")"
   '  con.executeQuery(sqlstr)	 
     '	sqlstr = "INSERT INTO ChangesNewsLetters (Object_Title,  Change_Date, [User_ID],NewsLetter_Value,Contact_Id,Change_Type) " & _
	'	" SELECT   IsNULL(CONTACT_NAME, '') + ' ' + IsNULL(cellular, ''), getDate()," & UserID & _
	'	"," &  is_NewLetters & "," & contactID &",'�����'" & _
	'	" FROM dbo.CONTACTS WHERE (CONTACT_ID = "& contactID &")"
'response.Write "sqlstr1="& sqlstr
'response.end
	'	con.executeQuery(sqlstr) 
    End if      		
      sqlUpd="UPDATE contacts SET CONTACT_NAME='"& CONTACT_NAME & "', first_name_E = '" & first_name_E & _
      "', last_name_E = '" & last_name_E & "', email='"& email  &"',phone='"& phone &"', fax='"& fax &_
      "', cellular='"& cellular &"', date_update=getDate()" &_
       ", phone_additional='"& phone_additional  &_    
      "', messanger_name='"& messangerName &"', contact_address = '" & contact_address &_
      "', contact_city_name = '" & contact_city_name & "', contact_zip_code = '" & contact_zip_code &_    
      "', contact_desc = '" & contact_desc & "', responsible_id = '" & responsible_id &_   
      "', company_id = "&companyID &", User_Id = " & UserId  &_   
      ", is_NotPermit_TourOrder = " & is_NotPermit_TourOrder   &_ 
      " WHERE contact_ID="& ContactId 
      'Response.Write(sqlUpd)
      'Response.End
      con.ExecuteQuery(sqlUpd)
   
    if  trim(Request.Cookies("bizpegasus")("Edit_IsNewLetters"))="1" then
    sqlUpd="UPDATE contacts SET is_NewLetters="& is_NewLetters &" WHERE contact_ID="& ContactId 
      con.ExecuteQuery(sqlUpd)
    end if
    
    ' ����� ������ ���� ����
    sqlstr="DELETE FROM contact_to_types WHERE contact_ID = " & ContactId
    con.executeQuery(sqlstr)	
	
	If  Len(contact_types) > 0 Then
		arrTypes = Split(contact_types & ",", ",")
		numOfTypes = Ubound(arrTypes)
	End If	
	
	If IsArray(arrTypes) And numOfTypes > 0 Then
		For i=0 To numOfTypes		
			If IsNumeric(arrTypes(i)) Then	
			sqlstr="INSERT INTO contact_to_types (contact_ID,type_id) VALUES ("&ContactId&","&arrTypes(i)&" )"			
			con.executeQuery(sqlstr)	
			End If		
		Next
	End If
	
If trim(wizard_id) <> "" Then%>
   <SCRIPT LANGUAGE=javascript>
   <!--
       window.opener.document.location.href= "contact_wizard.asp?companyId=<%=companyID%>&ContactId=<%=ContactId%>";
       self.close();
             
       
   //-->
  </SCRIPT>
<%Else%>
<%if LinkedContactid="" then%>
<script LANGUAGE=javascript>
       window.opener.document.location.href="contact.asp?companyId=<%=companyID%>&ContactId=<%=ContactId%>";
       self.close();

</script>
<%else%>
   <SCRIPT LANGUAGE=javascript>
   <!--
    //   window.opener.document.location.href="contact.asp?companyId=<%=companyID%>&ContactId=<%=ContactId%>";
    //  alert("ffffff")
      window.opener.opener.document.location.reload();
      window.opener.close();
      self.close();
   // alert(window.opener.opener.document.location.reload)
    //   self.close();
   //-->
  </SCRIPT>
  <%End If%>
  
<%End If%>
<%End If%>  
  
<%
  if trim(lang_id) = "1" then
	   arr_status_comp = Array("","�����","����","����","����")
  else
	   arr_status_comp = Array("","new","active","close","appeal")
  end if 
  
  If trim(ContactId)<>"" then  
	found_contact = false
	set listContact=con.GetRecordSet("EXEC dbo.contacts_contact_details @ContactId=" & ContactId & ", @OrgID=" & OrgID)
	if not listContact.EOF then 
		ContactId = cLng(listContact("contact_ID"))
		companyID = cLng(listContact("company_ID"))
		CONTACT_NAME = trim(listContact("CONTACT_NAME"))
		first_name_E = trim(listContact("first_name_E"))
		last_name_E = trim(listContact("last_name_E"))
		contacter = trim(CONTACT_NAME)
		email = trim(listContact("email"))
		phone = trim(listContact("phone"))
		extension_phone = trim(listContact("extension_phone"))
		cellular = trim(listContact("cellular"))
		phone_additional= trim(listContact("phone_additional"))
		fax = trim(listContact("fax"))
		messangerName = trim(listContact("messanger_name"))
		contact_address = trim(listContact("contact_address"))
		contact_zip_code = trim(listContact("contact_zip_code"))
		contact_city_Name = trim(listContact("contact_city_Name"))
		contact_desc = trim(listContact("contact_desc"))
		responsible_id = listContact("responsible_id")
		is_NewLetters = listContact("is_NewLetters")

		is_NotPermit_TourOrder = listContact("is_NotPermit_TourOrder")' Added by Mila 25/10/2018 for internet tour order from site pegasus

		found_contact = true     
		End if     
	  
		types=""
		set rstype = listContact.nextRecordSet()
		If not rstype.eof Then
			types = rstype.getString(,,",",",")						
		End If		    
		set rstype=Nothing		
		If Len(types) > 0 Then
			types = Left(types,(Len(types)-1))
		End If
		
		contact_types=""
		set rstype = listContact.nextRecordSet()		
		If not rstype.eof Then
			contact_types = rstype.getString(,,",",",")						
		End If		    
		set rstype=Nothing		
		If Len(contact_types) > 0 Then
			contact_types = Left(contact_types,(Len(contact_types)-1))
		End If
		set  listContact = Nothing		 
	End If	    
	
	If trim(is_companies) = 0 Then
  		sqlstr = "SELECT company_id FROM companies WHERE Organization_ID = " & orgID & " AND private_flag = 1"
		set rs_private = con.getRecordSet(sqlstr)
		if not rs_private.eof then		
			compPrivID = rs_private(0)	
		end if
		set rs_private = Nothing	
		
		If Not found_contact And trim(companyID) = "" Then
			companyID = compPrivID
		End If
	End If	
  
	found_company = false
	If trim(companyID)<>"" then   
	set pr=con.GetRecordSet("EXEC dbo.companies_company_details @CompanyID="& companyID & ", @OrgID=" & OrgID)
	if not pr.EOF then	
		company_name   = trim(pr("company_name"))
  		company_name_E = trim(pr("company_name_E"))
		address	      = trim(pr("address"))
		address2	  = trim(pr("address2"))
		street_number = trim(pr("street_number"))
		apartment	  = trim(pr("apartment"))
		post_box 	  = trim(pr("post_box"))
		cityName	  = trim(pr("city_Name"))
		zip_code	  = trim(pr("zip_code"))
		prefix_phone1 = trim(pr("prefix_phone1"))
		prefix_phone2 = trim(pr("prefix_phone2"))
		prefix_fax1	  = trim(pr("prefix_fax1"))
		prefix_fax2	  = trim(pr("prefix_fax2"))
		phone1	      = trim(pr("phone1"))
		phone2	      = trim(pr("phone2"))
		fax1          = trim(pr("fax1"))
		fax2	      = trim(pr("fax2"))
		company_email = trim(pr("email"))
		url	          = trim(pr("url"))
		date_update	  = trim(pr("date_update")	)
		status_company = trim(pr("status"))
		company_desc   = trim(pr("company_desc"))
		private_flag   = trim(pr("private_flag"))
		If Len(company_desc) > 70 Then
  			company_desc_short = Left(company_desc,70) & ".."
  		Else
  			company_desc_short = company_desc
  		End If
		found_company = true	 
	End if   
	Else
		private_flag = "0" 
	End if  
		 
		sqlstr = "SELECT word_id,word FROM translate WHERE (lang_id = " & lang_id & ") AND (page_id = 12) ORDER BY word_id"				
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
	  
		sqlstr = "SELECT value FROM buttons WHERE (lang_id = " & lang_id & ") ORDER BY button_id"
		set rsbuttons = con.getRecordSet(sqlstr)
		If not rsbuttons.eof Then
			arrButtons = ";" & rsbuttons.getString(,,",",";")		
			arrButtons = Split(arrButtons, ";")
		End If
		set rsbuttons=nothing	  %>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">

<SCRIPT LANGUAGE=javascript>
<!--
function check() 
{ 
   var alertstr = '';
   var cellular = document.getElementById("cellular").value;
   var phone = document.getElementById("phone").value;
  // alert (cellular)
   if(cellular != '' || phone != '')
   {
   if(window.XMLHttpRequest && !(window.ActiveXObject))
	 var xmlHTTP = new XMLHttpRequest();
    else
		if(window.ActiveXObject)
		//alert("window.ActiveXObject")
			var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP"); 
//  var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");

      xmlHTTP.open("POST", 'check_phone.asp', false);
      xmlHTTP.send('<?xml version="1.0" encoding="UTF-8"?><request><cellular>' + cellular + '</cellular><phone>' + phone + '</phone><ContactId><%=ContactId%></ContactId></request>');
      
    // window.document.write (xmlHTTP.ResponseText)

      
      if (xmlHTTP.responseText == "1")
      <%If trim(lang_id) = "1" Then%>
		alertstr = "���� ����� ����� ���� ���� ���� ��� ����� ������� \n"	
	  <%Else%>
	    alertstr = "The phone number, that you provided was found in the data base \n"	
	  <%End If%>	  
	}	 
  
   if (alertstr != "")
   {	
	  window.alert(alertstr);	
	 return false;
   }
 
   return true; 
}

//-->
</SCRIPT>

<script LANGUAGE="JavaScript">
function CheckFields(subButton)
{    
   
		if (document.getElementById("CONTACT_NAME").value=='')
		{
 		   <%
				If trim(lang_id) = "1" Then
					str_alert = "�� ���� ��"
				Else
					str_alert = "Please, insert the full name"
				End If	
			%>
 			window.alert("<%=str_alert%>");
 			document.getElementById("CONTACT_NAME").focus();
			return false;
		}
		<%if IsNumeric(LinkedContactid) then%>
		if (document.getElementById("ContactLinksType").value=='')
		{  <%
				If trim(lang_id) = "1" Then
					str_alert = "�� ����� ��� ���"
				Else
					str_alert = "Please, ckeck status"
				End If	
			%>
		
			window.alert("<%=str_alert%>");
			document.getElementById("ContactLinksType").focus();
			return false;
 		}
		<%end if%>
		
		if(check() == false)
		<%If trim(ContactId) = "" OR trim(ContactId) = "0" Then%>
			 return false;
		<%End If%>	 
			 		
	   if(document.getElementById("companyID").value == "")
       {
		if (document.getElementById("company_name").value=='')
		{
		   <%
				If trim(lang_id) = "1" Then
					str_confirm = "" & Space(22) & "��� ��, �� ����� �� ���� \n\n ?��� ������ ������ ��� ��� ��� ����"
				Else
					str_confirm = "Are you sure want to add customer without company?"
				End If	
			%>
 			var answer = window.confirm("<%=str_confirm%>");
 			if(answer == false)
 			{
 				document.getElementById("company_name").focus();
				return false;
			}	
		}		
		if (!checkEmail(document.getElementById("company_email").value) && document.getElementById("company_email").value != "")
		{
           <%
				If trim(lang_id) = "1" Then
					str_alert = "����� ���� �������� �� �����"
				Else
					str_alert = "The email address is not valid!"
				End If	
			%>
				window.alert("<%=str_alert%>");
				document.getElementById("company_email").focus();
				return false;
		}
		
    }
    if (document.getElementById("email").value=='')
	{
			<%
				If trim(lang_id) = "1" Then
					str_confirm = Space(80) & ",��� ��\n\n ,��  ���� ���� ���� " & trim(Request.Cookies("bizpegasus")("ContactsOne")) & " �� ���� ������� ������\n\n" & Space(55) & "? ��� ������ ������"
				Else
					str_confirm = "Pay attention !"&Space(80)&" \n\n You didn't insert the email, so the " & trim(Request.Cookies("bizpegasus")("ContactsOne")) & " will not be included in the mailing list. \n\n Are you sure want to proceed ?"& Space(55)
				End If	
			%>
			answer = window.confirm("<%=str_confirm%>");
			if(answer == false)
			{
				document.getElementById("email").focus();
				return false;
			}	
	}
	else if (!checkEmail(document.getElementById("email").value))
	{
			<%
				If trim(lang_id) = "1" Then
					str_alert = "����� ���� �������� �� �����"
				Else
					str_alert = "The email address is not valid!"
				End If	
			%>
			window.alert("<%=str_alert%>");
			return false;
	} 

	//subButton.disabled = true; 

    //window.opener.name = "contact";                 
    //document.form_contact.target = "contact";
    document.getElementById("form_contact").submit();  
    <%if IsNumeric(LinkedContactid) then%>      
//window.close();
   <%end if%>
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

function DoCal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}

var comp_list=null;
function openWindow()
{
	comp_list = window.open("companies_list.asp","List","width=550,height=380,top=50,left=50");
	return false;		
}

function ClearCompany()
{
	if(document.getElementById("companyId") != null)
	{
		document.getElementById("companyId").value = "";		
	}	
		
	if(document.getElementById("COMPANY_NAME") != null)	
	{
		document.getElementById("COMPANY_NAME").value = "";
		document.getElementById("COMPANY_NAME").readOnly = false;
		document.getElementById("COMPANY_NAME").className = "Form";
	}	
		
	if(document.getElementById("cityName") != null)
	{		
		document.getElementById("cityName").value = "";		
		document.getElementById("cityName").readOnly = false;
		document.getElementById("cityName").className = "Form";
	}	
		
	if(document.getElementById("address") != null)
	{
		document.getElementById("address").value = "";	
		document.getElementById("address").readOnly = false;
		document.getElementById("address").className = "Form";
	}			


	if(document.getElementById("phone1") != null)
	{
		document.getElementById("phone1").value = "";			
		document.getElementById("phone1").readOnly = false;
		document.getElementById("phone1").className = "Form";
	}			

		
	if(document.getElementById("phone2") != null)
	{
		document.getElementById("phone2").value = "";				
		document.getElementById("phone2").readOnly = false;
		document.getElementById("phone2").className = "Form";
	}
		
	if(document.getElementById("fax1") != null)
	{
		document.getElementById("fax1").value = "";			
		document.getElementById("fax1").readOnly = false;
		document.getElementById("fax1").className = "Form";
	}		
	
	if(document.getElementById("company_email") != null)	
	{
		document.getElementById("company_email").value = "";
		document.getElementById("company_email").readOnly = false;
		document.getElementById("company_email").className = "Form";
	}	
		
	if(document.getElementById("url") != null)	
	{
		document.getElementById("url").value = "";
		document.getElementById("url").readOnly = false;
		document.getElementById("url").className = "Form";
	}
	
	if(document.getElementById("company_desc") != null)	
	{
		document.getElementById("company_desc").value = "";
		document.getElementById("company_desc").style.height = "85px";
		document.getElementById("company_desc").style.lineHeight = "16px";
		document.getElementById("company_desc").readOnly = false;
		document.getElementById("company_desc").className = "Form";
	}
	
	if(document.getElementById("cities_list") != null)	
	{
		document.getElementById("cities_list").disabled = false;
		document.getElementById("cities_list").style.display = "inline";
		document.getElementById("cities_list").style.visible = true;
	}
	return false;
}
//-->
</script>  
</head>
<body style="margin:0;background:'#e6e6e6'" onload="window.focus()">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" dir="<%=dir_var%>">
<tr><td width="100%" class="page_title" dir="<%=dir_obj_var%>">&nbsp;<font style="color:#000000;"><b><%If trim(ContactId)<>"" Then%><!--����--><%=arrTitles(1)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%><%Else%><!--����--><%=arrTitles(2)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%><%End If%></b></font>&nbsp;</td></tr>
<tr><td width="100%">
<FORM name="form_contact" ACTION="newcontact.asp?ContactId=<%=ContactId%>&LinkedContactid=<%=LinkedContactid%>" METHOD="post" ID="form_contact">   
<table cellpadding=0 bgcolor="#e6e6e6" cellspacing=0 width=100% border=0 dir="ltr">
<tr>      
  <td align="<%=align_var%>" width="40%" nowrap valign="top" style="border: 1px solid #808080;border-left:none;border-top:none">  
  <input type="hidden" name="companyID" id="companyID" value="<%=companyID%>">  
  <input type="hidden" name="CONTACT_ID" id="CONTACT_ID" value="<%=ContactId%>"> 
   <table border="0" cellpadding="2" cellspacing="0" width="100%" dir="<%=dir_var%>" 
   <%If trim(is_companies) = "0" Then%> style="display: none;" <%End If%> >
      <tr><td colspan=2  align="<%=align_var%>" height=5 nowrap></td></tr>
      <tr>
         <td width="100%" align="<%=align_var%>">
         <INPUT type=image src="../../images/delete_icon.gif" name="delete_company" id="delete_company" onclick="return ClearCompany()" style="position:relative;top:3px">
         <INPUT type=image src="../../images/icon_find.gif" name="companies_list" id="companies_list" onclick="return openWindow()">
         <input type=text dir="<%=dir_obj_var%>" name="company_name" value="<%=vFix(company_name)%>"  style="width:200" <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> MaxLength=100 ID="Text1"></td>
         <td width="100" nowrap align="<%=align_var%>" style="padding-right:10px">&nbsp;<!--��--><%=arrTitles(4) & " " & Request.Cookies("bizpegasus")("CompaniesOne")%>&nbsp;<font style="color:#FF0000;font-size:11px">*</font></td> 
      </tr>            
      <tr>
        <td align="<%=align_var%>"><input type=text dir="<%=dir_obj_var%>" name="address" value="<%=vFix(address)%>"  style="width:200" <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="address"></td>
        <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--�����--><%=arrTitles(5)%>&nbsp;</td>
     </tr>             
     <tr>
         <td align="<%=align_var%>">
         <%If private_flag = "0" Then%>
         <INPUT type=image src="../../images/icon_find.gif" name="cities_list" id="cities_list" onclick='window.open("cities_list.asp","CitiesList","left=100,top=50,width=450,height=450,scrollbars=1");return false;'>&nbsp;
         <%End If%>
         <input type=text dir="<%=dir_obj_var%>" name="cityName" value="<%=vFix(cityName)%>" size="20" <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="cityName">         
         </td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--���--><%=arrTitles(6)%>&nbsp;</td>
      </tr>
      <tr>
          <td align="<%=align_var%>" dir=ltr>
		   <input type=text dir="ltr" name="zip_code" value="<%=vFix(zip_code)%>" size="10" maxlength=10 <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="zip_code">
           </td>
          <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--�����--><%=arrTitles(20)%>&nbsp;</td>
      </tr>                                
      <tr>
          <td align="<%=align_var%>" dir=ltr>
		   <input type=text dir="ltr" name="phone1" value="<%=vFix(phone1)%>" size="20" maxlength=20 <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="phone1">
           </td>
          <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--1 �����--><%=arrTitles(7)%>&nbsp;</td>
      </tr>                                
      <tr>  
          <td align="<%=align_var%>" dir=ltr>
		  <input type=text dir="ltr"  name="phone2" value="<%=vFix(phone2)%>" size="20" maxlength=20 <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="phone2">
          </td>
          <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--2 �����--><%=arrTitles(8)%>&nbsp;</td>
      </tr> 
       
      <tr>  
          <td align="<%=align_var%>" dir=ltr>            
			<input type=text dir="ltr"  name="fax1" value="<%=vFix(fax1)%>" size="20" maxlength=20 <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="fax1">
           </td>
          <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--���--><%=arrTitles(9)%>&nbsp;</td>
      </tr>                          
       <tr>
         <td align="<%=align_var%>">         
         <input dir=ltr type=text name="company_email" value="<%=vFix(company_email)%>" style="width:200" <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="company_email"></td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;Email&nbsp;</td>
      </tr>   
      <tr>
         <td align="<%=align_var%>">         
         <input dir=ltr type=text name="url" value="<%=vFix(url)%>" style="width:200" <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="url"></td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--���--><%=arrTitles(10)%>&nbsp;</td>
      </tr>
      <tr>
      <td align="<%=align_var%>">
      <textarea dir="<%=dir_obj_var%>" <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> style="width:250" name="company_desc" rows=5 ID="company_desc"><%=trim(company_desc)%></textarea>
      </td>
      <td align="<%=align_var%>" style="padding-right:10px" valign=top>&nbsp;<!--����� ������--><%=arrTitles(12)%>&nbsp;</td>
      </tr> 
      <tr><td colspan="2" height="5" nowrap></td></tr>                     
</table>
</td>
<td width="60%" nowrap valign="top" style="border: 1px solid #808080;border-left:none;border-top:none">
   <table border="0" cellpadding="0" cellspacing="0" width="90%" dir="<%=dir_var%>" align="center">
      <tr><td colspan=2  align="<%=align_var%>" height=5 nowrap></td></tr>
      <tr>
         <td width="65%" nowrap align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">           
         <input type=text dir="<%=dir_obj_var%>" name="CONTACT_NAME" id="CONTACT_NAME" value="<%=vFix(CONTACT_NAME)%>" maxlength=50 style="width:180" class="Form" ID="Text2">           
         </td>
         <td width="35%" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" bgcolor="#e6e6e6" valign="top"><font style="color:#FF0000;font-size:11px">*</font>&nbsp;<!--�� ���--><%=arrTitles(14)%>&nbsp;</td>
      </tr>
      <tr>
         <td width="65%" nowrap align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">           
         <input type=text dir="<%=dir_obj_var%>" name="first_name_E" value="<%=vFix(first_name_E)%>" maxlength="50" style="width:180px" class="Form" ID="first_name_E">           
         </td>
         <td width="35%" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;�� ����&nbsp;(������)</td>
      </tr>
      <tr>
         <td width="65%" nowrap align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">           
         <input type=text dir="<%=dir_obj_var%>" name="last_name_E" value="<%=vFix(last_name_E)%>" maxlength="50" style="width:180px" class="Form" ID="last_name_E">           
         </td>
         <td width="35%" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;�� �����&nbsp;(������)</td>
      </tr>            
      <tr><td colspan="2" align="<%=align_var%>" dir="<%=dir_obj_var%>" style="color:#FF0000; font-size:11px; padding: 5px; ">�� ����� ���� �������  ���� ���� �� �� ����� ������</td></tr>
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">		 
         <INPUT type=image src="../../images/icon_find.gif" name=messangers_list id="messangers_list" onclick='window.open("messangers_list.asp","MessangersList","left=300,top=50,width=250,height=380,scrollbars=1");return false;'>&nbsp;
         <input type=text dir="<%=dir_obj_var%>" name="messangerName" value="<%=vFix(messangerName)%>" style="width:180" maxlength=50 class="Form" ID="messangerName">         
         </td>         
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<!--�����--><%=arrTitles(15)%>&nbsp;</td>
      </tr>
      <tr>
        <td align="<%=align_var%>"><input type=text dir="<%=dir_obj_var%>" name="contact_address" value="<%=vFix(contact_address)%>"  style="width:180" class="Form" ID="contact_address"></td>
        <td align="<%=align_var%>">&nbsp;<!--�����--><%=arrTitles(5)%>&nbsp;</td>
     </tr>             
     <tr>
         <td align="<%=align_var%>" nowrap>
		 <INPUT type=image src="../../images/icon_find.gif" onclick='window.open("cities_list.asp?fieldID=contact_city_Name","CitiesList","left=100,top=50,width=450,height=450,scrollbars=1");return false'>&nbsp;                             
         <input type=text dir="<%=dir_obj_var%>" class="Form" name="contact_city_Name" id="contact_city_Name" value="<%=contact_city_Name%>">
         </td>
         <td align="<%=align_var%>">&nbsp;<!--���--><%=arrTitles(6)%>&nbsp;</td>
      </tr> 
      <tr>
          <td align="<%=align_var%>" dir=ltr>
		   <input type=text dir="ltr" name="contact_zip_code" value="<%=vFix(contact_zip_code)%>" size="10" maxlength=10 class="Form" ID="contact_zip_code">
           </td>
          <td align="<%=align_var%>">&nbsp;<!--�����--><%=arrTitles(20)%>&nbsp;</td>
      </tr>            
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" dir=ltr>             
         <input type=text dir="ltr" name="phone" value="<%=vFix(phone)%>" size="20" maxlength=20 class="Form" ID="phone">
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<!--�����--><%=arrTitles(16)%>&nbsp;</td>
      </tr>
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" dir=ltr>             
         <input type=text dir="ltr" name="cellular" value="<%=vFix(cellular)%>" size="20" maxlength=20 class="Form" ID="cellular">              
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<!--����� ����--><%=arrTitles(17)%>&nbsp;</td>
      </tr>
          <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" dir=ltr>             
         <input type=text dir="ltr" name="phone_additional" value="<%=vFix(phone_additional)%>" size="20" maxlength=20 class="Form" ID="Text3">              
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;����� ����&nbsp;</td>
      </tr>
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top" dir=ltr>                   
         <input type=text dir="ltr" name="fax" value="<%=vFix(fax)%>" size="20" maxlength=20 class="Form" ID="fax">
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<!--���--><%=arrTitles(18)%>&nbsp;</td>
      </tr>
      <tr>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">                        
         <input type=text name="email" id="email" dir=ltr value="<%=vFix(email)%>" style="width:180" maxlength=50 class="Form" ID="Text6">              
         </td>
         <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;Email&nbsp;</td>
      </tr> 
      <tr>
           <td align="<%=align_var%>">
           <INPUT type=image src="../../images/icon_find.gif" name=types_list id="types_list" onclick='window.open("types_list.asp?ContactId=<%=ContactId%>&contact_types=" + contact_types.value,"TypesList","left=300,top=100,width=250,height=350,scrollbars=1"); return false;'>&nbsp;
           <span class="Form_R" dir="<%=dir_obj_var%>" style="width:180;line-height:16px" name="types_names" id="types_names"><%=types%></span>
           <input type=hidden name=contact_types id=contact_types value="<%=contact_types%>">
           </td>
           <td align="<%=align_var%>" valign=top>&nbsp;<!--�����--><%=arrTitles(11)%>&nbsp;</td>
      </tr>  
      <tr>
      <td align="<%=align_var%>" valign=top>
      <textarea class="Form" dir="<%=dir_obj_var%>" style="width:250;" rows=5 name="contact_desc" id="contact_desc"><%=trim(contact_desc)%></textarea>      
      </td>
      <td align="<%=align_var%>" valign=top>&nbsp;<!--����� ������--><%=arrTitles(12)%>&nbsp;</td>
      </tr>  
      <tr>
        <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">                        
		<select name="responsible_id" dir="<%=dir_obj_var%>" class="norm" style="width:100%" ID="responsible_id">
			<option value="" id=word17 name=word17><!-- ��� --><%=arrTitles(19)%></option>
			<%set UserList=con.GetRecordSet("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME FROM Users WHERE ORGANIZATION_ID = " & OrgID & " ORDER BY FIRSTNAME + ' ' + LASTNAME")
			do while not UserList.EOF
			selUserID=UserList(0)
			selUserName=UserList(1)%>
			<option value="<%=selUserID%>" <%if trim(responsible_id)=trim(selUserID) then%> selected <%end if%>><%=selUserName%></option>
			<%
			UserList.MoveNext
			loop
			set UserList=Nothing%>
		</select>	          
		</td>
        <td align="<%=align_var%>" bgcolor="#e6e6e6" valign="top">&nbsp;<!--�����--><%=arrTitles(21)%>&nbsp;</td>
      </tr>
      <tr><td colspan=2 height=5></td></tr>    
   <%if  trim(Request.Cookies("bizpegasus")("Edit_IsNewLetters"))="1" then%>
       <tr>
      <td align="<%=align_var%>" valign=top>
      
      <input type="checkbox" dir="ltr" name="is_NewLetters" id="is_NewLetters" <%If is_NewLetters="True" or is_NewLetters="" then%> checked<%end if%>  >
      </td>
      <td align="<%=align_var%>" valign=top>&nbsp;������� ����� �����&nbsp;</td>
      </tr>  
      <%end if%>
        <tr><td colspan=2 height=5></td></tr>    
      <%if IsNumeric(LinkedContactid) then
        set listContact=con.GetRecordSet("select CONTACT_NAME from CONTACTS where contact_id=" & LinkedContactid )
  if not listContact.EOF then 
          contacter = trim(listContact("CONTACT_NAME"))
  end if
     set  listContact = Nothing  
       set TypeLink=con.GetRecordSet("select Contact_Type from ContactLinks where (Contact_Id="&LinkedContactid &" and Contact_Linked="& contactId &") or (Contact_Id="& contactId &" and Contact_Linked="&  LinkedContactid &")")
     if not  TypeLink.Eof then
     contacterType=trim(TypeLink("Contact_Type"))
     end if
     set TypeLink=Nothing
     
     
     %>
       <tr>
        <td align="<%=align_var%>" bgcolor="#FFD011" valign="top">                        
		<select name="ContactLinksType" dir="<%=dir_obj_var%>" class="norm" style="width:100%" ID="ContactLinksType">
			<option value="" ><!-- ��� -->��� ��� �����</option>
			<%set TypeList=con.GetRecordSet("SELECT Type_Id,Type_Name FROM ContactLinks_Type  ORDER BY Type_Name")
			do while not TypeList.EOF
			selTypeID=TypeList(0)
			selTypeName=TypeList(1)%>
			<option value="<%=selTypeID%>" <%if trim(contacterType)=trim(selTypeID) then%> selected <%end if%>><%=selTypeName%></option>
			<%
			TypeList.MoveNext
			loop
			set TypeList=Nothing%>
		</select>	          
		</td>
        <td align="<%=align_var%>" bgcolor="#FFD011" valign="top">&nbsp; ��� �&nbsp;-<%=contacter%></td>
      </tr>    
      <%end if%>
      <!-- is_NotPermit_TourOrder  Added by Mila 25/10/2018 for internet tour order from site pegasus-->
        <%if contactId<>"" then %>     
         <tr>
              <td align="right" valign=top>
                    <table cellpadding="3" cellspacing="0">
                    <tr>
                    <td <%If is_NotPermit_TourOrder="True" then%> style="background-color:red"<%end if%>>
                     <input type="checkbox" dir="ltr" name="is_NotPermit_TourOrder" id="is_NotPermit_TourOrder" <%If is_NotPermit_TourOrder="True" then%> checked <%end if%>  >
                    </td>
                    </tr>
                    </table>
              </td>
              <td align="<%=align_var%>" valign=top>&nbsp;���� ������ ���������&nbsp;</td>
        </tr>  
      <%end if%>   
</table></td></tr></table>
</form>
</td>
</tr>
<tr><td colspan="3" align=center dir=<%=dir_var%>>
<input type=button class="but_menu" value="<%=arrButtons(2)%>" onClick="window.close();" style="width:80" ID="Button2" NAME="Button2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=button class="but_menu" value="<%=arrButtons(1)%>" onClick="return CheckFields(this)" style="width:80" ID="Button1" NAME="Button1"></td></tr>
</table>
</body>
</html>
<%set con=Nothing%>
<% server.ScriptTimeout=10000 %>
<% Response.Buffer = False %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<% companyID = trim(Request("companyID"))
   contactID = trim(Request("contactID")) 
%>
<HTML>
<HEAD>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</HEAD>
<BODY class="body_admin">
<div id="div_save" bgcolor="#e8e8e8" name="div_save" style="position:absolute; left:0px; top:0px; width:100%; height:100%;" >  												
  <table bgcolor="#e8e8e8" height="100%" width="100%" cellspacing="2" cellpadding="2">  
     <tr><td bgcolor="#e8e8e8" align="center" >
         <table bgcolor="#ebebeb" border="0" height="100" width="400" cellspacing="1" cellpadding="1">
            <tr>  
              <td align="center" bgcolor="#d0d0d0">
              <font style="font-size:14px;color:#FF0000;"><b>מתבצעת שמירת נתונים</b></font>
              <br>
              <font style="font-size:14px;color:#000000;">...אנא המתן</font>
              </td>
            </tr>
         </table>
         </td>
     </tr>
  </table>
</div>
</BODY>
</HTML>
<%
 if Request.Form("companyID")=nil then  
	
	company_name=trim(sFix(Request.Form("company_name")))
	company_name_E=trim(sFix(Request.Form("company_name_E")))
	address=trim(sFix(Request.Form("address")))
	address2=trim(sFix(Request.Form("address2")))	
	street_number=trim(sFix(Request.Form("street_number")))
	apartment=trim(sFix(Request.Form("apartment")))	
	post_box=trim(sFix(Request.Form("post_box")))
	zip_code=trim(sFix(Request.Form("zip_code")))
	email=trim(sFix(Request.Form("email")))
	cityName=trim(sFix(Request.Form("cityName")))
	url=trim(sFix(Request.Form("url")))
	prefix_phone1=trim(sFix(Request.Form("prefix_phone1")))
	prefix_phone2=trim(sFix(Request.Form("prefix_phone2")))
	prefix_fax1=trim(sFix(Request.Form("prefix_fax1")))
	prefix_fax2=trim(sFix(Request.Form("prefix_fax2")))
	phone1=trim(sFix(Request.Form("phone1")))
	phone2=trim(sFix(Request.Form("phone2")))
	fax1=trim(sFix(Request.Form("fax1")))
	fax2=trim(sFix(Request.Form("fax2")))	
	email=trim(sFix(Request.Form("email")))
	status=trim(Request.Form("status"))
	If trim(status) = "" Then
		status = "2"
	End If	
	company_desc = sFix(Left(trim(Request.Form("company_desc")),500))
	
   if trim(companyID)<>"" then
      sqlUpd="UPDATE companies SET company_name='" & company_name &_
      "', date_update=getDate() , company_desc='" & company_desc & "'" &_
      ", address='"& address & "', city_Name='" & cityName & "', url='" & url & "'"&_
      ", prefix_phone1='"& prefix_phone1 &"', prefix_phone2='"& prefix_phone2 &_
      "', prefix_fax1='"& prefix_fax1 &"', prefix_fax2='"& prefix_fax2 &"'" &_
      ", phone1='"& phone1 &"', phone2='"& phone2 &"', fax1='"& fax1 &_
      "', fax2='"& fax2 &"', status='"& status &"', email='"& email &"'"&_   
      "  WHERE company_ID="& companyID
     'Response.Write(sqlUpd)
     'Response.End
      con.ExecuteQuery(sqlUpd)
   else      
      sqlIns="SET NOCOUNT ON; INSERT INTO COMPANIES(company_name,date_update,private_flag,company_desc"&_      
      ", address,url,prefix_phone1,prefix_phone2,prefix_fax1,prefix_fax2"&_
      ", phone1,phone2,fax1,fax2,zip_code,city_Name,email,status,User_Id,Organization_ID "&_      
      ")  VALUES ( '"& company_name &"',getDate(), 0,'"&company_desc & "','" &_      
      address &"','"& url &"','"& prefix_phone1 &"','"& prefix_phone2 &"','"&_
      prefix_fax1 &"','"& prefix_fax2 &"','"& phone1 &"','"& phone2 &"','"&_
      fax1 &"','"& fax2 &"','"& zip_code &"','"& cityName &"','" & email & "','" & status & "',"&_
      UserID & "," & OrgID & "); SELECT @@IDENTITY AS NewID"
      'Response.Write(sqlIns)
      'Response.End     
	  set rs_tmp = con.getRecordSet(sqlIns)
		companyID = rs_tmp.Fields("NewID").value
	  set rs_tmp = Nothing	  
	end if	

End If
' REM: add User ====================================================
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
   contact_desc = sFix(Left(trim(Request.Form("contact_desc")),500))
   contact_types = trim(Request.Form("contact_types"))
   
   if trim(contactID)<>"" then   		
      sqlUpd="UPDATE contacts SET CONTACT_NAME='"& CONTACT_NAME &_
      "', email='"& email  &"',phone='"& phone &"', fax='"& fax &_
      "', cellular='"& cellular &"', date_update=getDate()" &_
      ", messanger_name='"& messangerName &"', contact_address = '" & contact_address &_
      "', contact_city_name = '" & contact_city_name & "', contact_desc = '" & contact_desc &_    
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
      " contact_address,contact_city_name,contact_desc,Organization_ID) VALUES ("&_
      companyID & ", '"& CONTACT_NAME & "','" & email & "','" & phone &_
      "','"& fax & "','" & cellular & "', getDate(),'" & messangerName & "','" & contact_address & "','" &_
      contact_city_name & "','" & contact_desc & "'," & OrgID & "); SELECT @@IDENTITY AS NewID"
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
    
 end if  

If trim(wizard_id) <> "" Then%>
   <SCRIPT LANGUAGE=javascript>
   <!--
       document.location.href= "contact_wizard.asp?companyId=<%=companyID%>&contactID=<%=contactID%>";
   //-->
  </SCRIPT>
<%Else%>
   <SCRIPT LANGUAGE=javascript>
   <!--
       document.location.href="contact.asp?companyId=<%=companyID%>&contactID=<%=contactID%>";
   //-->
  </SCRIPT>
<%End If%>
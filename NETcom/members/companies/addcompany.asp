<% Server.ScriptTimeout=10000 %>
<% Response.Buffer = False %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%companyID = trim(Request("companyID"))%>
<%
if isNumeric(trim(UserID)) = true And IsNumeric(trim(OrgID)) then%>
<HTML>
<HEAD>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</HEAD>
<BODY class="body_admin">
<div id="div_save" bgcolor="#e8e8e8" name="div_save" style="position:absolute; left:0px; top:0px; width:100%; height:100%; " >  												
  <table bgcolor="#e8e8e8" height="100%" width="100%" cellspacing="2" cellpadding="2">  
     <tr><td bgcolor="#e8e8e8" align="center" >
         <table bgcolor="#ebebeb" border="0" height="100" width="400" cellspacing="1" cellpadding="1">
            <tr>  
              <td align="center" bgcolor="#d0d0d0">
              <font style="font-size:14px;color:#FF0000;"><b>מתבצעת שמירת נתונים</b></font>
              <br>
              <font style="font-size:14px;color:#000000;">... אנא המתן</font>
              </td>
            </tr>
         </table>
         </td>
     </tr>
  </table>
</div>
</BODY>
</HTML>
<%' REM: add company ====================================================%>
<%
 'For each item In Request.Form
  'Response.Write Request.Form(item) & " " & item & "<br>"
 'Next 
 'Response.End
%> 
<%  company_name=trim(sFix(Request.Form("company_name")))
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
      "', date_update=getDate() ,company_desc='" & company_desc & "'" &_
      ", address='"& address & "', city_Name='" & cityName & "', url='" & url & "'"&_
      ", prefix_phone1='"& prefix_phone1 &"', prefix_phone2='"& prefix_phone2 &_
      "', prefix_fax1='"& prefix_fax1 &"', prefix_fax2='"& prefix_fax2 &"'" &_
      ", phone1='"& phone1 &"', phone2='"& phone2 &"', fax1='"& fax1 &_
      "', fax2='"& fax2 &"', zip_code='"& zip_code &"', email='"& email &"'"&_   
      ",status = '" & status & "'  WHERE company_ID="& companyID
     'Response.Write(sqlUpd)
     'Response.End
      con.ExecuteQuery(sqlUpd)
   else      
      sqlIns="SET NOCOUNT ON; INSERT INTO companies(company_name,date_update,company_desc"&_      
      ", address,url,prefix_phone1,prefix_phone2,prefix_fax1,prefix_fax2"&_
      ", phone1,phone2,fax1,fax2,zip_code,city_Name,email,status,User_Id,Organization_ID "&_      
      ")  VALUES ( '"& company_name &"',getDate(), '" & company_desc & "','"&_      
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
If trim(wizard_id) <> "" Then
%>
<SCRIPT LANGUAGE=javascript>
  <!--
   document.location.href='company_wizard.asp?companyId=<%=companyID%>';
  //-->
</SCRIPT>
<%
Else
%>
<SCRIPT LANGUAGE=javascript>
  <!--
   document.location.href='company.asp?companyId=<%=companyID%>';
  //-->
</SCRIPT>
<%End If%>


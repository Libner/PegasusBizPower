<% server.ScriptTimeout=10000 %>
<% Response.Buffer = True %>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%  
    sort = Request.QueryString("sort")     
    companyID = trim(Request("companyID"))
    
    If Request.Form("company_name") <> nil Then
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
			",  address='"& address & "', city_Name='" & cityName & "', zip_code='" & zip_code &_  
			"', url='" & url & "'" & ", prefix_phone1='"& prefix_phone1 & "', prefix_phone2='"& prefix_phone2 &_
			"', prefix_fax1='"& prefix_fax1 &"', prefix_fax2='"& prefix_fax2 &"'" &_
			",  phone1='"& phone1 &"', phone2='"& phone2 &"', fax1='"& fax1 &_
			"', fax2='"& fax2 &"', email='"& email &"', status = '" & status &"', User_Id=" & UserId &_ 
			"   WHERE company_ID="& companyID
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
		End If
	
	
If trim(wizard_id) <> "" Then
%>
<SCRIPT LANGUAGE=javascript>
  <!--
		window.opener.document.location.href='company_wizard.asp?companyId=<%=companyID%>';
		self.close();
  //-->
</SCRIPT>
<%
Else
%>
<SCRIPT LANGUAGE=javascript>
  <!--
		window.opener.document.location.href='company.asp?companyId=<%=companyID%>';
		self.close();
  //-->
</SCRIPT>
<%End If%>
<%End If%>
<%
  
  if trim(lang_id) = "1" then
	arr_status_comp = Array("","עתידי","פעיל","סגור","פונה")
  else
	arr_status_comp = Array("","new","active","close","appeal")
  end if  
  
  If trim(companyID)<>"" then   
  sqlStr = "SELECT company_name,company_name_E,address,address2"&_
  ", street_number,apartment,post_box,city_Name,zip_code,prefix_phone1" &_
  ", prefix_phone2,prefix_fax1,prefix_fax2,phone1,phone2,fax1,fax2,url,email"&_
  ", date_update, status,company_desc FROM companies WHERE company_id="& companyID 
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
	cityName      =pr("city_Name")	
	zip_code	  =pr("zip_code")	
	prefix_phone1 =pr("prefix_phone1")
	prefix_phone2 =pr("prefix_phone2")
	prefix_fax1	  =pr("prefix_fax1")
	prefix_fax2	  =pr("prefix_fax2")
	phone1	      =pr("phone1")
	phone2	      =pr("phone2")
	fax1          =pr("fax1")
	fax2	      =pr("fax2")
	email         =pr("email")
	company_desc  =pr("company_desc")
	url	          =pr("url")	
	If trim(url) = "" Then
		url = "http://"
	End If
	date_update	  =pr("date_update")    
    status  =pr("status")	   
  end if  
 Else
	url = "http://"  : status = 2
 End if   
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 11 Order By word_id"				
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
<SCRIPT LANGUAGE=javascript>
<!--
function check() 
{ 
   var alertstr = "";
   var name = window.document.all("company_name").value;
 
   if(name != "")
   {
	  var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");
      xmlHTTP.open("POST", 'check_company.asp', false);
      xmlHTTP.send('<?xml version="1.0" encoding="UTF-8"?><request><cName>' + name + '</cName></request>');
      
      //window.alert(xmlHTTP.ResponseText)
      if (xmlHTTP.ResponseText == "1")
      <%If trim(lang_id) = "1" Then%>
		alertstr = "שם <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%> נמצא יותר מפעם אחד במאגר הנתונים \n"	
	  <%Else%>
	    alertstr = "The <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%> name, that you provided was found in the data base \n"	
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
 
		if (window.document.all("company_name").value=='')
		{
		   <%
				If trim(lang_id) = "1" Then
					str_alert = "נא למלא שם"
				Else
					str_alert = "Please, insert the name"
				End If	
			%>
 			window.alert("<%=str_alert%>");
 			window.document.all("company_name").focus();
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
		if (!checkEmail(document.all("email").value) && document.all("email").value != "")
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
		if(window.document.all("company_id").value=='' && (check() == false))
			return false;		
		   
		subButton.disabled = true;
		//window.opener.name = "company";                 
		//document.form_company.target = "company";
		document.form_company.submit();        
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
//-->
</script>  
</head>
<body style="margin:0;background:'#e6e6e6'">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" dir="<%=dir_var%>">
<tr><td class="page_title" align=center>&nbsp;<%if trim(companyId)<>"" then%><font style="color:#000000;"><b><span id=word1 name=word1><!--עדכן--><%=arrTitles(1)%></span>&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></b></font><%elseIf trim(companyId) = "" Then%><font style="color:#000000;"><b><span id="word2" name=word2><!--הוספת--><%=arrTitles(2)%></span> <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></b><%end if%>&nbsp;</td></tr>         	
<tr><td height=10 nowrap></td></tr>
<tr><td width="100%" dir="<%=dir_var%>">
<table cellpadding=1 bgcolor="#e6e6e6" cellspacing=0 width=100% border=0 style="border-collpase:collapse;" dir="<%=dir_var%>">
<FORM name="form_company" ACTION="editcompany.asp?companyID=<%=companyID%>" METHOD="post" ID="form_company">   
  <input type="hidden" name="COMPANY_ID" id="COMPANY_ID" value="<%=companyID%>">  
      <tr>
         <td width="75%" align="<%=align_var%>"><input type=text dir="<%=dir_obj_var%>" name="company_name" value="<%=vFix(company_name)%>"  style="width:200" class="Form" MaxLength=100 ID="Text1"></td>
         <td width="25%" align="<%=align_var%>" style="padding-right:10px">&nbsp;<!--שם--><%=arrTitles(4)%>&nbsp;<font style="color:#FF0000;font-size:11px">*</font></td> 
      </tr>            
      <tr>
        <td align="<%=align_var%>"><input type=text dir="<%=dir_obj_var%>" name="address" value="<%=vFix(address)%>"  style="width:200" class="Form" ID="address"></td>
        <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--כתובת--><%=arrTitles(5)%>&nbsp;</td>
     </tr>             
     <tr>
         <td align="<%=align_var%>">
         <INPUT type=image src="../../images/icon_find.gif" onclick='window.open("cities_list.asp","CitiesList","left=100,top=50,width=450,height=450,scrollbars=1");return false;' title="<%=arrTitles(14)%>">&nbsp;
         <input type=text dir="<%=dir_obj_var%>" name="cityName" value="<%=vFix(cityName)%>"  style="width:200" class="Form" ID="cityName">         
         </td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--עיר--><%=arrTitles(6)%>&nbsp;</td>
      </tr>  
      <tr>
          <td align="<%=align_var%>" dir=ltr>
		   <input type=text dir="ltr" name="zip_code" value="<%=vFix(zip_code)%>" size="10" maxlength=10 class="Form" ID="zip_code">
           </td>
          <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--מיקוד--><%=arrTitles(15)%>&nbsp;</td>
      </tr>                            
      <tr>
          <td align="<%=align_var%>" dir=ltr>
		   <input type=text dir="ltr" name="phone1" value="<%=vFix(phone1)%>" size="20" maxlength=20 class="Form" ID="phone1">
           </td>
          <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--1 טלפון--><%=arrTitles(7)%>&nbsp;</td>
      </tr>                                
      <tr>  
          <td align="<%=align_var%>" dir=ltr>
		  <input type=text dir="ltr"  name="phone2" value="<%=vFix(phone2)%>" size="20" maxlength=20 class="Form" ID="phone2">
          </td>
          <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--2 טלפון--><%=arrTitles(8)%>&nbsp;</td>
      </tr>  
      <tr>  
          <td align="<%=align_var%>" dir=ltr>            
			<input type=text dir="ltr"  name="fax1" value="<%=vFix(fax1)%>" size="20" maxlength=20 class="Form" ID="fax1">
           </td>
          <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--פקס--><%=arrTitles(9)%>&nbsp;</td>
      </tr>                          
       <tr>
         <td align="<%=align_var%>">         
         <input dir=ltr type=text name="email" value="<%=vFix(email)%>" style="width:200" class="Form"  ID="email"></td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;Email&nbsp;</td>
      </tr>   
      <tr>
         <td align="<%=align_var%>">         
         <input dir=ltr type=text name="url" value="<%=vFix(url)%>" style="width:200" class="Form"  ID="url"></td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--אתר--><%=arrTitles(10)%>&nbsp;</td>
      </tr>
      <tr>
      <td align="<%=align_var%>">
      <textarea dir="<%=dir_obj_var%>" class=Form style="width:330" name="company_desc" rows=5 ID="company_desc"><%=trim(company_desc)%></textarea>
      </td>
      <td align="<%=align_var%>" style="padding-right:10px" valign=top>&nbsp;<!--פרטים נוספים--><%=arrTitles(12)%>&nbsp;</td>
      </tr>
      <tr>
      <td align="<%=align_var%>">
		    <select id="status" name="status" dir="<%=dir_obj_var%>" class="norm">
			  <%For i=1 To Ubound(arr_status_comp)%>
			   <option value="<%=i%>" <%If trim(status) = trim(i) Then%>selected<%End If%>>&nbsp;<%=arr_status_comp(i)%>&nbsp;</option>
			  <%Next%> 
			</select>
		 </td>
	 <td align="<%=align_var%>" style="padding-right:10px" valign=top>&nbsp;<!--סטטוס--><%=arrTitles(13)%>&nbsp;</td>	 
	</tr>	 
    <tr><td colspan=2 height="15" nowrap></td></tr>
    <tr>
		 <td colspan=2 align="center" nowrap dir=<%=dir_var%>>
		 <input type=button class="but_menu" value="<%=arrButtons(2)%>" onClick="window.close();" style="width:80" ID="Button2" NAME="Button2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		 <input type=button class="but_menu" value="<%=arrButtons(1)%>" onClick="return CheckFields(this)" style="width:80" ID="Button1" NAME="Button1">
		 </td>		 
	</tr>
    </FORM>                        
</table>
</td>
</tr>
</table>
</body>
<%set con=Nothing%>
</html>


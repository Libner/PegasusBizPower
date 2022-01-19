<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%
	Response.CharSet = "windows-1255"  
	OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
	UserID=trim(trim(Request.Cookies("bizpegasus")("UserID")))
	CurrUserName=trim(trim(Request.Cookies("bizpegasus")("UserName")))
	lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
	SURVEYS = trim(Request.Cookies("bizpegasus")("SURVEYS"))
	COMPANIES = trim(Request.Cookies("bizpegasus")("COMPANIES"))  
	if lang_id = "1" then
		arr_Status = Array("","עתידית","הסתיימה","הוכנס סיכום","נדחתה")	
	else
		arr_Status = Array("","Future","Done","Summary added","Postponed")	
	end if	
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"  :  align_var = "left"  :  dir_obj_var = "ltr"
	Else
		dir_var = "ltr"  :  align_var = "right"  :  dir_obj_var = "rtl"
	End If
		'------------------------------------ Check organization sms   ---------------------		
	sql = "SELECT sms_Text FROM ORGANIZATIONS WHERE ORGANIZATION_ID = " & OrgID
	set rs_org = con.getRecordSet(sql)	
	If not rs_org.eof Then
		smsText = trim(rs_org(0))
	End If
	set rs_org = Nothing	
	
	If isNumeric(smsText) Then
		smsText = cLng(smsText)
	End If
	'------------------------------------ End of Check organization sms    ---------------------		
	companyid=trim(Request("companyId"))
	contactid=trim(Request("contactID"))

	If trim(companyid) <> "" Then
		sqlstr = "SELECT company_Name FROM companies WHERE company_Id = " & cLng(companyid)
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			companyname = trim(rs_pr("company_Name"))
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(contactid) <> "" Then
		sqlstr = "SELECT contact_Name,cellular FROM contacts WHERE contact_Id = " & cLng(contactid)
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			contactname = trim(rs_pr("contact_Name"))
			cellular=trim(rs_pr("cellular"))
		End If
		set rs_pr = Nothing
	End If	

	If Request.Form("add") <> nil Then	
	SMSTypeId= trim(Request.Form("SMS_type_id"))  
	companyId=trim(Request.Form("companyId"))
    contactID=trim(Request.Form("contactID"))
	sms_phone=trim(Request.Form("sms_phone"))
	cellular=trim(Request.Form("cellular"))
	SMS_content=trim(Request.Form("SMS_content"))
	smsStatusId="0"
		sqlstr="SET DATEFORMAT DMY; SET NOCOUNT ON; Insert Into sms_to_contact (company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send) "&_
			" values (" & companyID & "," & contactID & "," & SMSTypeId & "," & smsStatusId & ",'" & cellular & "','" & sms_phone & "','" & SMS_content  & "'," & UserID & ",GetDate()); SELECT @@IDENTITY AS NewID"		
		'Response.Write sqlstr
		'Response.End
			set rs_tmp = con.getRecordSet(sqlstr)
				smsID = rs_tmp.Fields("NewID").value
			set rs_tmp = Nothing						
	
	'insert into [sms_to_contact]
	%>
	<%	'sendUrl = "http://www.micropay.co.il/ExtApi/ScheduleSms.php" 
		'	'sendUrl = strLocal & "netcom/members/sms/ScheduleSms.asp"	
		'	getUrl = strLocal & "netcom/members/sms/alertSms.asp"			
		 '
		'	Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP") 
		'	xmlhttp.open "POST", sendUrl, false 
		'	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
		'	xmlhttp.send "uid=1269&un=atraf&msg=" &  Server.URLEncode(SMS_content) & "&charset=iso-8859-8" & _
		'	"&from=" & sms_phone & "&post=2&list=" & Server.URLEncode(cellular) & _
		'	"&desc=" & Server.URLEncode(orgName & "_" & smsId) & "&nu=" & Server.URLEncode(getUrl)
		'	strResponse = xmlhttp.responseText 
		'	Set xmlhttp = Nothing 
		'	
		'	tid=0
		'	error=""
		'	
		'	If InStr(strResponse, "ERROR") > 0 Then
		'		error = Mid(strResponse, 6)	
		'	ElseIf InStr(strResponse, "OK") > 0 Then	
		'		tid	 = trim(Mid(strResponse, 3))
		'		If isNumeric(tid) Then
		'			tid = cLng(tid)
		'		Else
		'		    tid = 0
		'		 End If				 
		'	End If		%>
	<SCRIPT LANGUAGE=javascript>
	<!--	
		opener.focus();
		var loc_str = new String(opener.window.location.href);
		if(loc_str.search("default.asp")>0)
		{
			opener.window.location.href = "default.asp?date_=<%=meeting_date%>&participant_id=<%=participant_id%>";
		}
		else
		{
			opener.window.location.reload(true);
		}	
		self.close();
	//-->
	</SCRIPT>
<%end if
			
	
	
				
	 %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
function checkTel(cell)
{
  if(cell == '')
  {
	 window.alert("!טלפון נייד לא קיים .לא ניתן לשלוח SMS ");
	closeWin();
    }
 if (cell.length!=10)
	 {
	 	window.alert('!מספר טלפון >10 ספרות');
		closeWin();

	 }  
	 	 		var cellTxt=cell
	 	
if(!/^(0[23489]([- ]{0,1})[^0\D]{1}\d{6})$|^(05[0247]{1}([- ]{0,1})[0-9]{7})$/.test(cellTxt))
					{
				window.alert("מספר טלפון לא חוקי");
		closeWin();
			}
			else
			{
				if(/^(050)|(052)|(054)|(057)/.test(cellTxt))
					isMobile=true
				else
				return false;
				
			}
    
}
//start of pop-ups related functions//
function textLimit(field, maxlen) {
if (field.value.length > maxlen + 1)
alert('הקלט נחתך!');
if (field.value.length > maxlen)
field.value = field.value.substring(0, maxlen);
} 
	function openCompaniesList()
	{
		window.open('../appeals/companies_list.asp','winList','top=50, left=50, width=740, height=600, scrollbars=1');
		return false;
	}	
	
	function removeCompany()
	{
		window.document.getElementById("companyID").value = "";
		window.document.getElementById("CompanyName").value = "";
		
		window.document.getElementById("project_id").value = "";
		window.document.getElementById("projectName").value = "";
		
		window.document.getElementById("contactID").value = "";
		window.document.getElementById("contactName").value = "";		
		
		window.document.getElementById("contacter_body").style.display = "none";
		window.document.getElementById("project_body").style.display = "none";	 
		return false;
	}
	


	function openContactsList()
	{
		var companyId = document.getElementById("companyID").value;
		window.open('../appeals/contacts_list.asp?companyId=' + companyId,'winList','top=50, left=50, width=740, height=600, scrollbars=1');
		return false;
	}	
	
	
//end of pop-ups related functions//

function closeWin()
{
	opener.focus();
	self.close();
}	

function CheckFields()
{

    if(window.document.formSendSMS.SMS_type_id.value == '')
  {
	 window.alert("! נא בחר סוג ההודעה");
	 window.document.formSendSMS.SMS_type_id.focus();
	 return false;
  }     
 
  if(window.document.formSendSMS.sms_phone.value == '')
  {
	 window.alert("! נא להכניס טלפון שולח ");
	 window.document.formSendSMS.sms_phone.focus();
	 return false;
    }     
 if (document.getElementById("sms_phone").value.length!=10)
	 {
	 	window.alert('!נא למלא טלפון 10 ספרות');
		document.getElementById("sms_phone").focus();
		return false;

	 }
	 		var phoneTxt=document.getElementById("sms_phone").value
if(!/^(0[23489]([- ]{0,1})[^0\D]{1}\d{6})$|^(05[0247]{1}([- ]{0,1})[0-9]{7})$/.test(phoneTxt))
					{
				window.alert("מספר טלפון לא חוקי");
				window.document.getElementById("sms_phone").focus()
				return false;
			}
			else
			{
				if(/^(050)|(052)|(054)|(057)/.test(phoneTxt))
					isMobile=true
				else
				return false;
				
			}
	
   if(window.document.formSendSMS.SMS_content.value == '')
  {
	 window.alert("! נא להכניס  טקסט ההודעה");
	 window.document.formSendSMS.SMS_content.focus();
	 return false;
  } 
   document.formSendSMS.submit();               	 
   return true;
}




//-->
</script>  
</head>
<body style="margin:0px;background:#e6e6e6" onload="self.focus();checkTel('<%=cellular%>')">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" dir="<%=dir_var%>" ID="Table1">
	
<tr>	     
	<td class="page_title"  dir="<%=dir_obj_var%>">&nbsp;שליחת SMS&nbsp;</td>  
</tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" width="100%" align=center dir="<%=dir_var%>" ID="Table2">
<FORM name="formSendSMS" ACTION="AddSMS.asp" METHOD="post" ID="formSendSMS">
<tr>
   <td align="<%=align_var%>" width=100% nowrap bgcolor="#e6e6e6" valign="top" style="border: 1px solid #808080;border-top:none">
    <input type=hidden name=add id=add value="1">   
     <input type=hidden name=contactId id="contactId" value="<%=contactId%>">   
     <input type=hidden name=companyId id="companyId" value="<%=companyId%>"> 
     <input type=hidden name=cellular id="cellular" value="<%=cellular%>">     
    
   <table border="0" cellpadding="0" align=center cellspacing="2" width=100% ID="Table3">          
   <tr>
   <td align="<%=align_var%>" colspan=4 style="padding:10px">
   <table cellpadding=1 cellspacing=1 border=0 width=100% dir="<%=dir_var%>"  ID="Table4">
   <tr><td align="<%=align_var%>" width=100%><%'=companyname%> &nbsp;<%=contactname%></td>
   <td align="right" width="140" nowrap dir="rtl"><b>שם</b></td>
   </tr>
<tr><td align="<%=align_var%>" width=100%><%=cellular%></td>
   <td align="right" width="140" nowrap dir="rtl"><b>טלפון נייד</b></td>
   </tr>

    <tr>
	<td align="<%=align_var%>" width=100%>
 <select name="SMS_type_id"  dir="rtl" class="norm" style="width:250" ID="MS_type_id">
    <option value="" id="" name="">בחר</option>
    <%set SMSType=con.GetRecordSet("SELECT SMS_type_id, SMS_type_name  FROM SMS_types")
    do while not SMSType.EOF
    SMSTypeId=SMSType(0)
    SMSTypeName=SMSType(1)%>
    <option value="<%=SMSTypeId%>"  ><%=SMSTypeName%></option>
    <%
    SMSType.MoveNext
    loop
    set SMSType=Nothing%>
     </select>	</td>
	<td align="<%=align_var%>" width=140 nowrap><b><!--תאריך-->סוג ההודעה <br>(לצורך תיעוד בלבד)</b></td>
   </tr>
<tr>
	<td align="right" width="100%" valign="top"><input dir="rtl" type="text" class="texts" 
	style="width:100;background-color:#c0c0c0;color:#808080" id="sms_phone" name="sms_phone" value="0500000000"  readonly maxlength="10"></td>
	<td align="right" width="140" nowrap dir="rtl"><b>טלפון שולח</b></td>
</tr>
   <tr>   
  


	  <tr>             
	<td align="<%=align_var%>" nowrap  dir="<%=dir_obj_var%>">
	<textarea id="SMS_content" name="SMS_content" dir="<%=dir_obj_var%>" onkeyup="textLimit(this, <%=smsText%>);" style="height:40;width:250;line-height:120%;" class="Form" ><%=SMS_content%></textarea>
	</td>	
	<td align="<%=align_var%>" width=140 nowrap valign=top><b>תוכן ההודעה<BR>(<%=smsText%> תווים )</b></td>
	</tr> 
</table></td></tr>
</table></td></tr>
      
</form>
<tr><td align=center colspan="2" height=20 nowrap></td></tr>
<tr><td align=center colspan="2">
<table cellpadding=0 cellspacing=0 width=100% ID="Table7">
<tr>
<td align=center width=50%>
<input type=button value="ביטול" class="but_menu" style="width:120" onclick="closeWin();" ID="Button2" NAME="Button2">
</td>
<td align=center width=50%>
<A class="but_menu" href="#" style="width:120;line-height:120%;padding:4px" onClick="return CheckFields()"><!--הוסף-->שלח</a>
</td>
</tr>
</table>
</td></tr>
<tr><td align=center colspan="2" height=10 nowrap></td></tr>
</table>
</div>
</body>
<%set con=Nothing%>
</html>
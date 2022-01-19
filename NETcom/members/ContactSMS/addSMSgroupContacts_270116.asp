<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<script>
function ContentOnchange(pId)
	{
	//alert(pId.value)
	//document.all("message_content").value="uuuuuu"
		if(window.XMLHttpRequest && !(window.ActiveXObject))
	 var xmlHTTP = new XMLHttpRequest();
    else
		if(window.ActiveXObject)
			var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");
	if (!xmlHTTP)
		var xmlHTTP = new ActiveXObject("Msxml2.XMLHTTP");
			if(pId.length > 0)
			{	 
				xmlHTTP.open("POST", 'GetSmsText.asp', false);
				xmlHTTP.send('<?xml version="1.0" encoding="UTF-8"?><request><typeId>' + pId.value + '</typeId></request>');	 
				result = new String(xmlHTTP.responseText);
				document.getElementById("SMS_content").value=result

		}	
	}
</script>
<%Response.CharSet = "windows-1255"
	orgName="pegasus"
	OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
	UserID=trim(trim(Request.Cookies("bizpegasus")("UserID")))
	CurrUserName=trim(trim(Request.Cookies("bizpegasus")("UserName")))
	lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
	SURVEYS = trim(Request.Cookies("bizpegasus")("SURVEYS"))
	COMPANIES = trim(Request.Cookies("bizpegasus")("COMPANIES"))  
'send  sms
	If Request.Form("add") <> nil Then	
	SMSTypeId= trim(Request.Form("SMS_type_id"))  
	sms_phone=trim(Request.Form("sms_phone"))
	SMS_content=trim(Request.Form("SMS_content"))
	SmsId=trim(Request("SmsId"))
	smsStatusId="0"
'response.Write 	SmsId
'response.end
		sqlstr = "Select sms_to_contact.Company_Id,sms_to_contact.Contact_Id,cellular from sms_to_contact left join Contacts on sms_to_contact.Contact_id=contacts.Contact_Id WHERE  len(cellular)>0 and SMS_Id IN (" & SmsId & ")"
'	response.Write sqlstr
'	response.end
	set rs_Contacts = con.getRecordSet(sqlstr)	
	do while not rs_Contacts.eof
	companyID=rs_Contacts("Company_Id")
	contactID=rs_Contacts("contact_id")
	cellular=rs_Contacts("cellular")
	response.Write left(cellular,3)&"<BR>"
	if len(cellular)<10 then
	smsStatusId=3 ' מספר שגוי
			sqlstr="SET DATEFORMAT DMY; SET NOCOUNT ON; Insert Into sms_to_contact (company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send) "&_
			" values (" & companyID & "," & contactID & "," & SMSTypeId & "," & smsStatusId & ",'" & cellular & "','" & sms_phone & "','" & SMS_content  & "'," & UserID & ",GetDate()); SELECT @@IDENTITY AS NewID"		
		'Response.Write sqlstr &"<Br>"
		'Response.End
			set rs_tmp = con.getRecordSet(sqlstr)
				smsID = rs_tmp.Fields("NewID").value
			set rs_tmp = Nothing						

	elseif  left(cellular,3)<>"050" and left(cellular,3)<>"052" and left(cellular,3)<>"054"  and left(cellular,3)<>"077"   then
	smsStatusId=3 ' מספר שגוי
			sqlstr="SET DATEFORMAT DMY; SET NOCOUNT ON; Insert Into sms_to_contact (company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send) "&_
			" values (" & companyID & "," & contactID & "," & SMSTypeId & "," & smsStatusId & ",'" & cellular & "','" & sms_phone & "','" & SMS_content  & "'," & UserID & ",GetDate()); SELECT @@IDENTITY AS NewID"		
		'Response.Write sqlstr &"<Br>"
		'Response.End
			set rs_tmp = con.getRecordSet(sqlstr)
				smsID = rs_tmp.Fields("NewID").value
			set rs_tmp = Nothing						

	else
	smsStatusId=0
	sendUrl = "http://www.micropay.co.il/ExtApi/ScheduleSms.php" 
		'	'sendUrl = strLocal & "netcom/members/sms/ScheduleSms.asp"	
		'	getUrl = strLocal & "netcom/members/sms/alertSms.asp"			
		 '
		' http://www.micropay.co.il/ExtApi/ScheduleSms.php?get=1&uid=10&un=test&msg=הודעת בדיקה&list=0545370070&charset=iso-8859-8&from=035555555

		' getUrl=""
			  '  xmlhttp.send "uid=2575&un=pegasus&msg=" &  Server.URLEncode(SMS_content) & "&charset=iso-8859-8" & _
			'	"&from=" & sms_phone & "&post=2&list=" & Server.URLEncode(cellular) & _
	
	
			Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP") 
			xmlhttp.open "POST", sendUrl, false 
			xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
		    xmlhttp.send "uid=2575&un=pegasus&msg=" &  Server.URLEncode(SMS_content) & "&charset=iso-8859-8" & _
			"&from=" & sms_phone & "&post=2&list=" & Server.URLEncode(cellular) & _
			"&desc=" & Server.URLEncode(orgName & "_" & smsId) 
			strResponse = xmlhttp.responseText 
			Set xmlhttp = Nothing 
		'	response.Write "strResponse="&strResponse
		'	response.end
		'	
			tid=0
			error=""
		'	
			If InStr(strResponse, "ERROR") > 0 Then
				error = Mid(strResponse, 6)	
				smsStatusId=2
			ElseIf InStr(strResponse, "OK") > 0 Then	
			smsStatusId=1
				tid	 = trim(Mid(strResponse, 3))
				If isNumeric(tid) Then
					tid = cLng(tid)
				Else
				    tid = 0
				 End If				 
			End If	
				sqlstr="SET DATEFORMAT DMY; SET NOCOUNT ON; Insert Into sms_to_contact (company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send) "&_
			" values (" & companyID & "," & contactID & "," & SMSTypeId & "," & smsStatusId & ",'" & cellular & "','" & sms_phone & "','" & SMS_content  & "'," & UserID & ",GetDate()); SELECT @@IDENTITY AS NewID"		
		'Response.Write sqlstr
		'Response.End
			set rs_tmp = con.getRecordSet(sqlstr)
				smsID = rs_tmp.Fields("NewID").value
			set rs_tmp = Nothing						
	end if
	

	rs_Contacts.MoveNext
	loop
	
		
	'insert into [sms_to_contact]
	%>
	
	<SCRIPT LANGUAGE=javascript>
	<!--	
		/*opener.focus();
		var loc_str = new String(opener.window.location.href);
		if(loc_str.search("default.asp")>0)
		{
			opener.window.location.href = "default.asp?date_=<%=meeting_date%>&participant_id=<%=participant_id%>";
		}
		else
		{
			opener.window.location.reload(true);
		}	*/
		opener.window.location.reload(true);
		self.close();
	//-->
	</SCRIPT>
<%end if
			
	
	
				


'end of send sms




  SmsId=trim(Request("SmsId"))
   	arrAppeals = Split(SmsId, ",")   	
  'response.Write SmsId
    

		If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"  :  align_var = "left"  :  dir_obj_var = "ltr"
	Else
		dir_var = "ltr"  :  align_var = "right"  :  dir_obj_var = "rtl"
	End If
		'------------------------------------ Check organization sms   ---------------------		
	sql = "SELECT sms_Text,SMS_Phone FROM ORGANIZATIONS WHERE ORGANIZATION_ID = " & OrgID
	set rs_org = con.getRecordSet(sql)	
	If not rs_org.eof Then
		smsText = trim(rs_org(0))
		sms_phone= trim(rs_org(1))
	End If
	set rs_org = Nothing	
	
	If isNumeric(smsText) Then
		smsText = cLng(smsText)
	End If
	


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
function textLimit(field, maxlen) {
if (field.value.length > maxlen + 1)
alert('הקלט נחתך!');
if (field.value.length > maxlen)
field.value = field.value.substring(0, maxlen);
} 
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
 /*if (document.getElementById("sms_phone").value.length!=10)
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
				
			}*/
	
   if(window.document.formSendSMS.SMS_content.value == '')
  {
	 window.alert("! נא להכניס  טקסט ההודעה");
	 window.document.formSendSMS.SMS_content.focus();
	 return false;
  } 
  var str=window.document.formSendSMS.SMS_content.value;
  var n = str.indexOf("'"); 
  var n1 = str.indexOf('"'); 
if (n>=0)
{
alert("שימו לב, תוכן ההודעה מכיל סימנים מיוחדים \n אנא הסרו אותם מההודעה")
return false;
}
if (n1>=0)
{
alert("שימו לב, תוכן ההודעה מכיל סימנים מיוחדים \n אנא הסרו אותם מההודעה")
return false;
}
   document.formSendSMS.submit();               	 
   return true;
}



function checkTel(cell)
{
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
<FORM name="formSendSMS" ACTION="AddSMSgroupContacts.asp" METHOD="post" ID="formSendSMS">
  <input type=hidden name=add id=add value="1">   
<input type=hidden name=SmsId id=SmsId value="<%=SmsId%>">
<table border="0" cellpadding="1" cellspacing="0" width="100%" align=center dir="<%=dir_var%>" ID="Table2">
  <tr><td align=center colspan=2  dir=rtl><b>סה''כ SMS לשליחה <%=Ubound(arrAppeals)+1%></b></td></tr>
   <tr>
	<td align="<%=align_var%>" width=100%>
 <select name="SMS_type_id"  dir="rtl" class="norm" style="width:250" ID="MS_type_id" onchange='javascript:ContentOnchange(this);'>
    <option value="" id="" name="">בחר</option>
    <%set SMSType=con.GetRecordSet("SELECT SMS_type_id, SMS_type_name  FROM SMS_types order by SMS_type_name")
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
	style="width:100;background-color:#c0c0c0;color:#808080" id="sms_phone" name="sms_phone" value="<%=sms_phone%>" maxlength="10" readonly></td>
	<td align="right" width="140" nowrap dir="rtl"><b>טלפון שולח</b></td>
</tr>
	  <tr>             
	<td align="<%=align_var%>" nowrap  dir="<%=dir_obj_var%>">
	<textarea id="SMS_content" name="SMS_content" dir="<%=dir_obj_var%>" onkeyup="textLimit(this, <%=smsText%>);" style="height:40;width:250;line-height:120%;" class="Form" ><%=SMS_content%></textarea>
	</td>	
	<td align="<%=align_var%>" width=140 nowrap valign=top><b>תוכן ההודעה<BR>(<%=smsText%> תווים )</b></td>
	</tr> 
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
</table> 
</form>
</body>
</html>
<%@ Language=VBScript%>
<%Response.Buffer = False%>
<%ScriptTimeout=6000%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%
  UserID = trim(trim(Request.Cookies("bizpegasus")("UserID"))) 
  OrgID = trim(trim(Request.Cookies("bizpegasus")("OrgID")))   
  smsId = trim(Request("smsId")) 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=iso-8859-8">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css" />
<style>
.normalB
{
    FONT-WEIGHT: bold;
    FONT-SIZE: 11pt;
    COLOR: #2e2380;
    FONT-FAMILY: Arial
}
</style>
</head>
<body onload="window.focus();">
<%strPhonesList = ""
	 sqlstr = "Exec dbo.get_sms_list @OrgID = '" & OrgID & "', @smsId = '" & smsId & "'"
	 Set rs_tmp = con.getRecordSet(sqlstr)	
	 If Not rs_tmp.Eof Then
		totalCount = rs_tmp.RecordCount
		strPhonesList = rs_tmp.getString(,,",",",")
	 End If
	 Set rs_tmp = Nothing
	 If Len(strPhonesList) > 0 Then
		strPhonesList = Left(strPhonesList,(Len(strPhonesList)-1))
	 End If
  
 	sqlStr = "SELECT COUNT(PEOPLE_ID) FROM sms_to_people WHERE (ORGANIZATION_ID="& OrgID &") " & _
	" AND (sms_id = " & smsId & ")"
	'Response.Write sqlStr
	set rsc = con.getRecordSet(sqlStr)
	If not rsc.eof Then
		totalAllCount = trim(rsc(0))
	Else
		totalAllCount = 0	
	End If
	set rsc = Nothing
  
  smsLimit = 0	  
  sql = "Select Sms_Limit, ORGANIZATION_NAME From Organizations WHERE ORGANIZATION_ID = " & OrgID
  set rs_org = con.getRecordSet(sql)	
  If not rs_org.eof Then
	smsLimit = trim(rs_org(0))
	orgName = trim(rs_org(1))
  End If
  set rs_org = Nothing	
	
  If isNumeric(smsLimit) Then
	smsLimit = cLng(smsLimit)
  End If	%>
<%If Request.QueryString("send_end") = nil Then
			sqlstr = "SELECT sms_phone, sms_desc, sms_text FROM sms_sends WHERE (sms_id=" & smsId &	") AND (ORGANIZATION_ID=" & OrgID & ")"
			set rs_sms = con.GetRecordSet(sqlStr)
			If Not rs_sms.Eof Then
				sms_phone = trim(rs_sms("sms_phone"))
				sms_desc = trim(rs_sms("sms_desc"))
				sms_text = trim(rs_sms("sms_text"))
			End If	
			Set rs_sms = Nothing
			
			sendUrl = "http://www.micropay.co.il/ExtApi/ScheduleSms.php" 
			'sendUrl = strLocal & "netcom/members/sms/ScheduleSms.asp"	
			getUrl = strLocal & "netcom/members/sms/alertSms.asp"			
		 	'Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP") 
			'xmlhttp.open "POST", sendUrl, false 
			'xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
			'xmlhttp.send "uid=1269&un=atraf&msg=" &  Server.URLEncode(sms_text) & "&charset=iso-8859-8" & _
			'"&from=" & sms_phone & "&post=2&list=" & Server.URLEncode(strPhonesList) & _
			'"&desc=" & Server.URLEncode(orgName & "_" & smsId) & "&nu=" & Server.URLEncode(getUrl)
			'strResponse = xmlhttp.responseText 
			'Set xmlhttp = Nothing 
		 
		 
			Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP") 
			xmlhttp.open "POST", sendUrl, false 
			xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded" 
			xmlhttp.send "uid=2575&un=pegasus&msglong=" &  Server.URLEncode(sms_text) & "&charset=iso-8859-8" & _
			"&from=" & sms_phone & "&post=2&list=" & Server.URLEncode(strPhonesList) & _
			"&desc=" & Server.URLEncode(orgName & "_" & smsId) & "&nu=" & Server.URLEncode(getUrl)
			strResponse = xmlhttp.responseText 
			Set xmlhttp = Nothing 
			
			tid=0
			error=""
			
			If InStr(strResponse, "ERROR") > 0 Then
				error = Mid(strResponse, 6)	
			ElseIf InStr(strResponse, "OK") > 0 Then	
				tid	 = trim(Mid(strResponse, 3))
				If isNumeric(tid) Then
					tid = cLng(tid)
				Else
				    tid = 0
				 End If				 
			End If		
			
			If tid > 0 Then 'שליחה הצליחה		    
				smsRemainder  = cLng(smsLimit) - cLng(totalCount)
				con.executeQuery("UPDATE ORGANIZATIONS SET SMS_Limit = '" & smsRemainder & "' WHERE ORGANIZATION_ID = " & OrgID)	
				con.executeQuery("UPDATE sms_sends SET tid = " & tid & ", sms_date = getDate(), error ='" & vFix(error) & "', status = 1 WHERE sms_id = " & smsId)	
				con.executeQuery("UPDATE sms_to_people SET date_send = getDate() WHERE sms_id = " & smsId)	
			Else
				con.executeQuery("UPDATE sms_sends SET error ='" & vFix(error) & "', sms_date = getDate() WHERE sms_id = " & smsId)	
			End If			%>
			<script language=javascript>
			<!--
						document.location.href = "sendSms.asp?send_end=1&smsId=<%=smsId%>";
			//-->
			</script>
			<%			 
   End If
   Set con=Nothing %>  
<table cellpadding=0 cellspacing=0 width="100%">
<%If Request.QueryString("send_end") = nil Then ' end of sending%>
<tr><td height="20"></td></tr>	
<tr>
	<td align="center" nowrap class="normalB" dir=rtl><font color="red">תהליך שליחת sms מתבצע כעת, התהליך יארך מספר דקות</font></td>
</tr>		
<tr>
	<td  align="center" nowrap class="normalB" dir=rtl><font color="red">אין לסגור חלון זה ואין להפעיל ישומים אחרים במחשב בזמן התהליך.</font></td>
</tr>	
<%End If%>
<tr><td height="30" nowrap></td></tr>
<%If Request.QueryString("send_end") <> nil Then ' end of sending%>
<tr><td align="center" class="td_subj">שליחת הודעות הסתיימה</td></tr>
<%ElseIf totalCount >= smsLimit And Request.QueryString("send_end") <> nil Then%>			
<tr><td align="center" class="td_subj">כיוון שיתרתך לשליחת הודעות נגמרה שליחת הודעות הסתיימה</td></tr>
<%ElseIf Request.QueryString("send_end") = nil Then%>
<tr><td align="center" class="td_subj">שליחת הודעות החלה</td></tr>
<tr><td height="10" nowrap></td></tr>
<tr><td align="center" class="td_subj" dir=rtl> סה"כ sms לשליחה <font color="#023296"><b><%=totalAllCount%></b></font></td></tr>
<%End If%>
<tr><td height="20" nowrap></td></tr>
<tr>
<td nowrap align="center"><input class="but_menu" style="width:90"  type="button" onclick="javascript:window.close()" 
value="סגור חלון" id=button3 name=button3></td>
</tr>
</table>	
</body>
</html>
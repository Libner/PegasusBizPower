<%@ Language=VBScript%>
<%Response.Buffer = False%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<BODY onload="window.focus();" style="margin:0px">
<%
   UserID = trim(Request.Cookies("bizpegasus")("UserID"))
   OrgID = trim(trim(Request.Cookies("bizpegasus")("ORGID")))
   smsId = trim(Request.Form("smsId"))
   groups_id = trim(Request.Form("groups_id"))	
   
   If (Not IsNumeric(smsId) Or trim(smsId) = "" Or IsNULL(smsId)) And isNumeric(OrgID) Then
		If Len(groups_id) > 0 Then
			sqlStr = "SELECT COUNT(PEOPLE_ID) FROM sms_peoples WHERE (ORGANIZATION_ID = "& OrgID &") " & _
			" AND GROUP_ID IN ("& groups_id &")"				
			set rsc = con.getRecordSet(sqlStr)
			If not rsc.eof Then
				totalCount = trim(rsc(0))
			Else
				totalCount = 0	
			End If
			set rsc = Nothing	
		Else
			totalCount = 0
		End If
		badCount = 0
		
	 ElseIf IsNumeric(OrgID) And IsNumeric(smsId) And Len(groups_id) > 0 Then
		sqlStr = "SELECT COUNT(PEOPLE_ID) FROM sms_to_people WHERE (ORGANIZATION_ID="& OrgID &") AND " & _
		" GROUP_ID IN ("& groups_id &") AND sms_id = " & smsId & " And (date_send is NULL) "
		set rsc = con.getRecordSet(sqlStr)
		If not rsc.eof Then
			totalCount = trim(rsc(0))
		Else
			totalCount = 0	
		End If
		set rsc = Nothing
   End If	
	  
	If IsNumeric(OrgID) And Len(groups_id) > 0 Then			
			sql = "SELECT Sms_Limit FROM Organizations WHERE ORGANIZATION_ID = " & OrgID
			set rs_org = con.getRecordSet(sql)	
			If not rs_org.eof Then
				smsLimit = trim(rs_org(0))
			End If
			set rs_org = Nothing	
				
			If isNumeric(smsLimit) Then
				smsLimit = cLng(smsLimit)
			End If	
	Else
			blockedCount = 0 : badCount = 0
    End If	
	
  If ((totalCount)>0) And ((smsLimit - (totalCount))=>0) Then	
   If Not IsNumeric(smsId) Or Len(smsId) = 0 Or IsNULL(smsId) then 	
	If trim(Request.Form("editFlag")) <> "" Then			
			sms_phone = sFix(trim(Request.Form("sms_phone")))
			sms_desc = sFix(trim(Request.Form("sms_desc")))
			sms_text = sFix(trim(Request.Form("sms_text")))		
			
			sqlstr = "Exec dbo.insert_sms @OrgID = '" & OrgID & "', @UserID = '" & UserId & "', @smsPhone = '" & sms_phone &_
			"', @smsDesc = '" & sms_desc & "', @smsText = '" & sms_text & "', @groups_id = '" & groups_id & "'"
			Set rs_tmp = con.getRecordSet(sqlstr)		
			If Not rs_tmp.eof Then
				smsId = rs_tmp(0)
			End If
			Set rs_tmp = Nothing					
		End If					
	End If
End If	%>
<script language="javascript">
<!--
	 function GoSend(btn)
	 {			
		btn.disabled = true;
		strUrl = "sendSMS.asp?send_groupsId=<%=groups_id%>&smsId=<%=smsId%>";		
		document.location.href = strUrl;
		window.opener.document.location.href = "default.asp";
	 }
//-->
</script>
<table cellpadding=2 cellspacing=2 width="100%" border=0>
<tr><td height="10" nowrap></td></tr>
<tr><td align="center" class="td_subj" dir=ltr> סה"כ הודעות לשליחה  בקבוצות שנבחרו <font color="#023296"><b><%=totalCount%></b></font></td></tr>
<tr><td align="center" class="td_subj" dir=ltr><font color="#023296" dir=ltr><b><%=smsLimit - (totalCount)%></b></font> יתרה לאחר השליחה</td></tr>
<tr><td height="10" nowrap></td></tr>
<%If ((totalCount)>0) And ((smsLimit - (totalCount))=>0) Then%>
<tr><td nowrap align="center"><input class="but_menu" style="width: 90px" type="button" onclick="GoSend(this)" value="שלח הפצה"></td></tr>
<%Else%>
<tr><td align="center" class="td_subj" dir=ltr><font color=red><b>לא ניתן לבצע את השליחה</b></font></td></tr>
<%If ((smsLimit - totalCount)<=0) Then%>
<tr><td align="center" class="td_subj" dir=ltr><font color=red><b>04 - לקנית יתרה נוספת ניתן להתקשר ל 8770282</font></td></tr>
<%End if%>
<%End if%>
</table>
</body>
</html>
<%Set con=Nothing%>
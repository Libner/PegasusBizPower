<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%smsId=trim(Request("smsId"))		
	 groups_id=trim(session("wgroupID"))
	 If trim(groups_id) <> "" Then
			sqlstr="Select GROUPNAME FROM sms_groups WHERE GROUP_ID IN (" & groups_id & ")"
			set rsn = con.getRecordSet(sqlstr)
			If not rsn.eof Then
				groupname = trim(rsn.getString(,," ; "," ; "))
			End if
			set rsn = Nothing	
		
			If Len(groupname) > 1 Then
				groups_list = Left(groupname,Len(groupname)-1) 
			End If
			sqlstr="SELECT COUNT(PEOPLE_ID) FROM sms_peoples WHERE GROUP_ID IN (" & groups_id & ") AND IsNull(PEOPLE_CELL, '') <> ''"
			set rs_count = con.getRecordSet(sqlstr)
			If not rs_count.eof Then
				count_all = trim(rs_count(0))
			End If		
			set rs_count = Nothing
			count_all = "(" & count_all & " נמענים)"
		End If
		
	'------------------------------------ Check organization sms limit  ---------------------		
	sql = "SELECT Sms_Limit FROM ORGANIZATIONS WHERE ORGANIZATION_ID = " & OrgID
	set rs_org = con.getRecordSet(sql)	
	If not rs_org.eof Then
		smsLimit = trim(rs_org(0))
	End If
	set rs_org = Nothing	
	
	If isNumeric(smsLimit) Then
		smsLimit = cLng(smsLimit)
	End If
	'------------------------------------ End of Check organization sms limit  ---------------------			
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css" />
</head>
<script language="javascript">
<!--
function CheckField() {
 
  if(window.document.dataForm.groups_id.value == '')
  {
	 window.alert("! חובה לבחור נמענים להפצה");	
	 return false;
  }  
  if(window.document.dataForm.sms_phone.value == '')
  {
	 window.alert("! נא להכניס טלפון שולח הפצה");
	 window.document.dataForm.sms_phone.focus();
	 window.document.dataForm.sms_phone.scrollIntoView();
	 return false;
  }   
  if(window.document.dataForm.sms_desc.value == '')
  {
	 window.alert("! נא להכניס נושא של הפצה");
	 window.document.dataForm.sms_desc.focus();
	 window.document.dataForm.sms_desc.scrollIntoView();
	 return false;
  } 
  if(window.document.dataForm.sms_text.value == '')
  {
	 window.alert("! נא להכניס תוכן של הפצה");
	 window.document.dataForm.sms_text.focus();
	 window.document.dataForm.sms_text.scrollIntoView();
	 return false;
  }   
  else
  {
	if(window.document.dataForm.sms_text.value.length >250)
	{
		window.alert("! ניתן להכניס תוכן של הפצה לא יותר מ250 תווים");
		window.document.dataForm.sms_text.focus();
		window.document.dataForm.sms_text.scrollIntoView();
		return false;
	}  
  } 
  if (window.confirm("?האם ברצונך לשלוח הודעה לנמענים שנבחרו"))
  {
	  save_win = window.open("","save_win","scrollbars=1,toolbar=0,top=100,left=100,width=450,height=250,align=center,resizable=1")
	  window.document.dataForm.target = "save_win";
	  window.document.dataForm.submit();
  }
  return false;
}

function popupcal(elTarget){
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

function GetEnglish()
{
	var ch=event.keyCode;
	event.returnValue = ch < 126;
}
 
function openGroups()
{
	window.open("groups_choose.asp?groups_id="+window.document.all("groups_id").value,"","left=250,top=100,height=350,width=390,status=no,toolbar=no,scroll=no,menubar=no,location=no,scrollbars=yes")
}

function checkLen(object, maxChars, divNum, langFlip, langFlipDec)
{
	if (langFlip)
	{
	var smsT=object.value.length//<%=smsText%> 
		//charToDec = 70 - maxChars;
		charToDec = smsT - maxChars;
		if (object.value.search(/[^\x00-\x7E]/) >= 0) ;
		else
		{
			if (langFlipDec) charToDec -= langFlipDec;
			//maxChars = 126 - charToDec;
			maxChars = smsT - charToDec;
		}
	}	

	elementId = "charCount" + divNum;
	//document.getElementById(elementId).style.display="block";	

	if (object.value.length > maxChars) 
	{
		object.value = object.value.substring(0,maxChars); 
		object.focus();
	}

	leftChars = maxChars - object.value.length;
	if (leftChars < 0) leftChars = 0;	
	document.getElementById(elementId).innerHTML = "נותרו לכם <font color='red'>" + leftChars + "</font> תווים";
	
	return true;	
}
//-->
</script>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 46%>
<%numOfLink = 1%>
<!--#include file="../../top_in.asp"-->
<div align="<%=align_var%>"><center>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<tr><td class="page_title" dir="<%=dir_obj_var%>">&nbsp;שליחת SMS&nbsp;-&nbsp;<span class="td_subj" style="color:#6E6DA6">יתרת sms לשליחה&nbsp;<font color=red><%=smsLimit%></font></font>&nbsp;</td></tr>
<tr><td width="100%">
<table cellpadding=0 cellspacing=0 width=100% bgcolor="#E6E6E6">
<tr><td height=10 nowrap></td></tr>
<tr><td>
<FORM name="dataForm" ACTION="save_sms.asp" METHOD="post" ID="dataForm">
<table align=center border="0" cellpadding="1" cellspacing="1" width="650">
<tr>
	<td align="<%=align_var%>" height=15 nowrap>&nbsp;</td>	
</tr>
<tr>
<td align="<%=align_var%>">
<table cellpadding=0 cellspacing=0 width="100%">
<tr>
	<td width=100% align="<%=align_var%>" id="groups_name" name="groups_name" dir="<%=dir_obj_var%>">&nbsp;<span dir="<%=dir_obj_var%>"><%=groups_list%></span>&nbsp;<font dir="<%=dir_obj_var%>" color=red><%=count_all%></font></td>
	<input dir="ltr" type="hidden" id="groups_id" name="groups_id" value="<%=vFix(groups_id)%>">
	<td align="<%=align_var%>" width="110" dir="<%=dir_obj_var%>" nowrap>	
	<input type=button class="but_menu" value="בחר קבוצות" onclick="return openGroups();" style="width:100">
	</td>
</tr>
</table>
</td></tr>
<tr>
<td align="<%=align_var%>">
<table cellpadding=0 cellspacing=0 width="100%">
<tr>
	<td align="<%=align_var%>" width="100%" valign="top"><input dir="<%=dir_obj_var%>" type="text" class="texts" style="width:350" 
	id="sms_desc" name="sms_desc" value="<%=vFix(smsdesc)%>" maxlength="200"></td>
	<td align="<%=align_var%>" width="110" nowrap dir="<%=dir_obj_var%>"><b>שם הפצה<br>(לצורך תיעוד בלבד)</b></td>
</tr>
</table>
</td>
</tr>
<tr>
<td align="<%=align_var%>">
<table cellpadding=0 cellspacing=0 width="100%" border="0" ID="Table1">
<tr>
	<td align="<%=align_var%>" width="100%" valign="top"><input dir="<%=dir_obj_var%>" type="text" class="texts" 
	style="width:350" id="sms_phone" name="sms_phone" value="<%=vFix(smsphone)%>" maxlength="10"></td>
	<td align="<%=align_var%>" width="110" nowrap dir="<%=dir_obj_var%>"><b>טלפון שולח</b></td>
</tr>
</table>
</td>
</tr>
<tr>
<td align="<%=align_var%>">
<table cellpadding=0 cellspacing=0 width="100%">
<tr>
	<td align="<%=align_var%>" width=100%><div id="charCount1"></div>
	<textarea dir="<%=dir_obj_var%>"  class="texts" style="width:350;height:100px" id="sms_text" name="sms_text" 
	 MaxLength="250<%'=smsText%>" onkeyup="return checkLen(this, 250<%'=smsText%>, 1, 1);" 
	 onchange="return checkLen(this, 250<%'=smsText%>, 1, 1);"><%=vFix(smstext)%></textarea>
	<!--input dir="<%=dir_obj_var%>" type="text" class="texts" style="width:350;height:100px" id="sms_text" name="sms_text" 
	value="<%=vFix(smstext)%>" MaxLength="250<%'=smsText%>" onkeyup="return checkLen(this, 250<%'=smsText%>, 1, 1);" onchange="return checkLen(this, <%=smsText%>, 1, 1);"-->
	</td>
	<td align="<%=align_var%>" width="110" nowrap dir="<%=dir_obj_var%>" valign="top"><b>תוכן ההודעה</b></td>
</tr>
</table>
</td>
</tr>
<tr><td height="10"><input type="hidden" name="smsId" value="<%=smsId%>"><input type="hidden" name="editFlag" value="yes"></td></tr>
<tr><td colspan="2">
<table width=100% border=0 cellspacing=0 dir="<%=dir_var%>">
<tr><td width=48% align="<%=align_var%>">
<INPUT class="but_menu" style="width:90" type="button" value="ביטול" id=button2 name=button2 onclick="document.location.href='default.asp'"></td>
<td width=4% nowrap></td>
<td width=48%  dir="<%=dir_var%>">
<input class="but_menu" style="width:90" type="button" value="אישור" onClick="return CheckField();"></td>
</tr>
</table>
</td></tr>
</table></form>
</td></tr></table>
</td></tr>
<tr><td height=10 nowrap></td></tr>
</table></center></div>
</body>
</html>
<%set con=nothing%>
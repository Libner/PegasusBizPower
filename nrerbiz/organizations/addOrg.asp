<!--#include file="../../netcom/connect.asp"-->
<!--#include file="../../netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<% Session.LCID = 1037 %>
<html>
<head>
	<title>Bizpower Administration</title>
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
	<meta charset="windows-1255">
	<link href="../../admin.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript" type="text/javascript">
<!--
function CheckFields()
{
	if (document.frmMain.ORGANIZATION_NAME.value=='')
	{
		window.alert("'חובה למלא את השדה 'שם ארגון ");
		document.frmMain.ORGANIZATION_NAME.focus();		
		return false;				
	}
	if (document.frmMain.Email_Limit.value=='')
	{
		window.alert("'חובה למלא את השדה 'יתרת מיילים נוכחית ");
		document.frmMain.Email_Limit.focus();		
		return false;				
	}
	if (document.frmMain.Email_Limit_Month.value=='')
	{
		window.alert("'חובה למלא את השדה 'יתרת מיילים חודשית ");
		document.frmMain.Email_Limit_Month.focus();		
		return false;				
	}		
	document.frmMain.submit();
	return true;			
}

function GetNumbers()
{
    var ch=event.keyCode;	
    	
	if(!(ch >= 48 && ch <= 57) || ch == 46 || ch == 13 || ch == 10)
		event.returnValue = false;
	else
		event.returnValue = true;
}

function CheckDel(field,orgId) 
//For picture deleting
{
  if(confirm("?האם ברצונך למחוק את התמונה") == true )
  {
	document.location.href='addOrg.asp?ORGANIZATION_ID='+orgId+'&'+field+'=1';
	return true;
  }
  return false;
}
function ifFieldEmpty(field){
	  if (field.value=='')
		{
		alert('חובה לבחור את התמונה');
		return false;
		}
	  else
		return true;
}

function check_all_bars(objChk,parentID)
{
	input_arr = document.getElementsByTagName("input");	
	for(i=0;i<input_arr.length;i++)	{
		
		if(input_arr(i).type == "checkbox")
		{
			currparentId = "";
			objValue = new String(input_arr(i).id);			
			value_arr = objValue.split("!");
			currparentId =  value_arr[1];
			if(currparentId == parentID)
			{
				//input_arr(i).disabled = objChk.checked;
				input_arr(i).checked = objChk.checked;
			}	
		}	
	}
	return true;
}
   function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../netcom/calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
   }
//-->
</script>  
<%
OrgID=trim(Request("ORGANIZATION_ID"))
'response.end
if request.form("ORGANIZATION_NAME")<>nil then 'after form filling	
	 	 
	 ORGANIZATION_NAME = sFix(Request.form("ORGANIZATION_NAME"))
	 Email_Limit = sFix(request.form("Email_Limit"))
	 Date_Add = sFix(request.form("Date_Add")) 
	 Email_Limit_Month = sFix(request.form("Email_Limit_Month")) 
	 SmsLimit = sFix(request.form("SmsLimit")) 
	 SmsText = sFix(request.form("SmsText")) 
	 SmsPhone = sFix(request.form("SmsPhone")) 
	 If trim(Request.Form("RowsInList")) <> "" Then
		RowsInList = trim(Request.Form("RowsInList"))
	 Else
		RowsInList = "30"	
	 End If	
	 
	 LANG_ID = Request.Form("LANG_ID") 
	 If Request.Form("FILE_UPLOAD") <> nil Then
		FILE_UPLOAD  = 1
	 Else
	 	FILE_UPLOAD  = 0
	 End If	
	 If Request.Form("UseBizLogo") <> nil Then
		UseBizLogo  = 1
	 Else
	 	UseBizLogo  = 0
	 End If
	 If Request.Form("UseCaptcha") <> nil Then
		UseCaptcha  = 1
	 Else
	 	UseCaptcha  = 0
	 End If	 
	 	 
	 if OrgID=nil or OrgID="" then 'new record in DataBase		    
		sqlStr = "SET NOCOUNT ON; SET DATEFORMAT DMY; Insert Into ORGANIZATIONS (ORGANIZATION_NAME,ACTIVE, "&_
		" Email_Limit,RowsInList,LANG_ID,FILE_UPLOAD, UseBizLogo,Date_Add,Email_Limit_Month, UseCaptcha, Sms_Limit,SMS_TEXT) Values ('" &_
		ORGANIZATION_NAME & "',1,'" & Email_Limit & "'," & RowsInList & "," & LANG_ID & ",'" & FILE_UPLOAD &_
		"','" & UseBizLogo & "','" & Date_Add & "','" & Email_Limit_Month & "','" & UseCaptcha & "','" & SmsLimit & "','"& SmsText &"'); SELECT @@IDENTITY AS NewID"
		'Response.Write sqlStr
		'Response.End
       	set rs_tmp = con.getRecordSet(sqlStr)
			OrgID = rs_tmp.Fields("NewID").value
		set rs_tmp = Nothing		
				
		'הוספת לקוח פרטי - לקוח אחד פר ארגון אשר כל אנשי קשר ללא חברה מקושרים אליו
		If trim(LANG_ID) = "1" Then
		arrComp = Array("ללא חברה")
		Else
		arrComp = Array("No company")
		End If
		For i=0 To Ubound(arrComp)
			sqlstr = "Insert into companies (Organization_ID,company_name,private_flag) values (" &_
			trim(OrgID) & ",'" & sFix(arrComp(i)) & "',1)"
			con.executeQuery(sqlstr)	
		Next
		
		'הוספת תפקידי ברירת המחדל
		If trim(LANG_ID) = "1" Then
		arrMessangers = Array("מנהל מערכת")
		Else
		arrMessangers = Array("System administrator")
		End If
		For i=0 To Ubound(arrMessangers)
			sqlstr = "Insert into jobs (Organization_ID,job_name,hour_pay) values (" &_
			trim(OrgID) & ",'" & sFix(arrMessangers(i)) & "',0)"
			con.executeQuery(sqlstr)	
		Next
		
		'הוספת תפקידי ברירת המחדל
		If trim(LANG_ID) = "1" Then
		arrMessangers = Array("מכירות","שיווק","כספים","אדמיניסטרציה","מנכ''ל","מנהל",_
		"אחר","טכני","כח אדם","הדרכה","מפתח","רכש")
		Else
		arrMessangers = Array("Technical","Marketing","H.r.","Manager","Financial","Another")
		End If
		For i=0 To Ubound(arrMessangers)
			sqlstr = "Insert into messangers (Organization_ID,item_Name) values (" &_
			trim(OrgID) & ",'" & sFix(arrMessangers(i)) & "')"
			con.executeQuery(sqlstr)	
		Next
		
		'הוספת סוגי פעילויות ברירת המחדל
		If trim(LANG_ID) = "1" Then
		arrActTypes = Array("מעקב הצעת מחיר","פגישה","פנית לקוח","שיחת טלפון",_
		"שיחת מכירה","שליחת הצעת מחיר","שליחת חומר","תלונות לקוח")
		Else
		arrActTypes = Array("request","telephone call","another")
		End If
		For i=0 To Ubound(arrActTypes)
			sqlstr = "Insert into activity_types (Organization_ID,activity_type_name) values (" &_
			trim(OrgID) & ",'" & sFix(arrActTypes(i)) & "')"
			con.executeQuery(sqlstr)	
		Next
		
		'הוספת קבוצת ארגונים ברירת המחדל
		If trim(LANG_ID) = "1" Then
		arrOrgTypes = Array("אופנה","אלקטרוניקה וחשמל","ביטוח","בית וגן","בנייה",_
		"בנקים","בריאות וטיפוח","בתי תוכנה","הייטק","חינוך","יועצים","כימיקלים",_
		"מזון","ממשלה","נוער","ספורט","רכב ותחבורה","שווק והפצה","תעשייה","תרופות",_
		"מדיה","ביוטכנולוגיה","שרותים","מזון בריאות","מכשור רפואי","בתי קפה","מסעדות",_
		"יחצנים / ליינים","מועדונים","משרדי פרסום","דאנס בר","בר","משקאות","טיפולי גוף",_
		"מכללות","בתי ספר","אריזות","מוצרי פלסטיק","קבוצה כללית")
		Else
		arrOrgTypes = Array("Electronics","Biotechnology","Banks","Hi-tech","Food",_
		"Medicines")		
		End If

		For i=0 To Ubound(arrOrgTypes)
			sqlstr = "Insert into contact_type (Organization_ID,type_name) values (" &_
			trim(OrgID) & ",'" & sFix(arrOrgTypes(i)) & "')"
			con.executeQuery(sqlstr)	
		Next	
		
		If trim(LANG_ID) = "1" Then
			LANGU = "heb"
		Else
			LANGU = "eng"
		End If	
		'הוספת טופס לפרויקט	
		sqlstr = "Insert Into Products (ORGANIZATION_ID,PRODUCT_TYPE,LANGU) values ("&trim(OrgID) & ",1,'"&LANGU&"')"
		con.executeQuery(sqlstr)
		
	 else
		sqlStr = "SET DATEFORMAT DMY; Update Organizations set Organization_Name='" & ORGANIZATION_NAME &_
		"', Email_Limit = '" & Email_Limit & "', RowsInList = " & RowsInList & ", File_Upload='" & File_Upload &_
		"', UseBizLogo='" & UseBizLogo & "', Date_Add = '" & Date_Add & "', Email_Limit_Month = '" &_
		Email_Limit_Month & "', UseCaptcha = '" & UseCaptcha & "', Sms_Limit = '" & SmsLimit  & "', Sms_Text = '" & SmsText  & "', Sms_Phone = '" & SmsPhone & "' Where ORGANIZATION_ID=" & OrgID
		'Response.Write sqlStr
		'Response.End
		con.GetRecordSet (sqlStr)			
    end if 
    
    sqlstr = "Delete FROM bar_organizations WHERE organization_id = " & OrgID
	con.executeQuery(sqlstr)
					
	sqlstr="Select bar_id from bars order by bar_order"
	set rs_bars = con.getRecordSet(sqlstr)
	While not rs_bars.eof
		barID = trim(rs_bars(0))
		barTitle = Request.Form("title_bar"&barID)		
		'barVisible = Request.Form("is_visible"&barID)
		If trim(Request.Form("is_visible"&barID)) = "on" Then
			barVisible = 1
		Else
			barVisible = 0
		End If			
				
		If barVisible = 1 Then			
			sqlstr = "Insert Into bar_organizations values (" & barID & "," & OrgID & ",'" & sFix(barTitle) & "','" & barVisible & "')"
			con.executeQuery(sqlstr)
		Else
		    barVisible = 0	
		    		
			sqlstr = "Insert Into bar_organizations values (" & barID & "," & OrgID & ",'" & sFix(barTitle) & "','" & barVisible & "')"
			con.executeQuery(sqlstr)			
			
			sqlstr = "Select User_ID FROM USERS WHERE ORGANIZATION_ID = " & OrgID
			set rs_users = con.getRecordSet(sqlstr)
			while not rs_users.eof
			    userID = trim(rs_users(0))			  				
			    
			    sqlstr="Delete From bar_users WHERE ORGANIZATION_ID = " & OrgID & " And User_ID = " & userID & " AND Bar_ID = " & barID
				con.executeQuery(sqlstr)
			
				sqlstr = "Insert Into bar_users values (" & barID & "," & orgID & "," & userID & ",'" & barVisible & "')"
				con.executeQuery(sqlstr)				
									
				rs_users.moveNext
			wend
			set rs_users = nothing
		End If				
		
		
	rs_bars.moveNext
	Wend
	set rs_bars = Nothing
	
	sqlstr = "Select User_ID FROM USERS WHERE ORGANIZATION_ID = " & OrgID
	set rs_users = con.getRecordSet(sqlstr)
	while not rs_users.eof
	    userID = trim(rs_users(0))
	    
	    sqlstr="Select bar_id, is_visible from bar_users WHERE ORGANIZATION_ID = " & OrgID & " And User_ID = " & UserID & " Order by bar_id"
		set rs_bars = con.getRecordSet(sqlstr)
		While not rs_bars.eof
		barID = trim(rs_bars(0))
		barVisible = trim(rs_bars(1))
		
		If barID = "1" Then	
			COMPANIES = barVisible		
		ElseIf barID = "2" Then
			SURVEYS = barVisible
		ElseIf barID = "3" Then		
			EMAILS = barVisible
		ElseIf barID = "4" Then	
			WORK_PRICING = barVisible
		ElseIf barID = "5" Then	
			PROPERTIES = barVisible
		ElseIf barID = "6" Then		
			TASKS = barVisible	
		ElseIf barID = "29" Then		
			CASH_FLOW = barVisible					
		End If					
		rs_bars.moveNext
		Wend
		set rs_bars = Nothing   		
		
		sqlStr = "UPDATE USERS SET WORK_PRICING='"&WORK_PRICING&"',SURVEYS='"&SURVEYS&"',EMAILS='"&EMAILS&_
		"', COMPANIES='"&COMPANIES&"', TASKS='"&TASKS&"', CASH_FLOW='"&CASH_FLOW&"' WHERE USER_ID="&userID
		'Response.Write sqlStr
		'Response.End
		con.GetRecordSet (sqlStr)	
	rs_users.moveNext
	wend
	set rs_users = nothing
				
	
	con.executeQuery("Delete FROM Titles WHERE Organization_ID = " & OrgID)
	obj="Select count(*) as j from objects"
	set rs=con.getRecordSet(obj)
	j=rs("j")
	For i=1 to j
		title_one=Request.Form("title_one"& i)
		title_multi=Request.Form("title_multi"& i)
		If title_one<>"" Or title_multi<>"" then
			sqlobj = "Insert into  Titles (Organization_ID,object_id,title_organization_one,title_organization_multi)" &_
			"values (" & trim(OrgID) & "," & i & ",'" & title_one & "','" & title_multi & "')"
			'Response.Write sqlobj &"<BR>"
			'response.end
			con.GetRecordSet(sqlobj)
		End if
	Next
    
Response.Redirect ("default.asp")     
end if
%>
<body bgcolor="#FFFFFF">
<div align="right"> 
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="5" class="page_title">דף ניהול&nbsp;<%if OrgID<>nil then%>עדכון<%else%>הוספת<%end if%>&nbsp;ארגון&nbsp;</td>
  </tr>
  <tr>
     <td width="5%" align="center"><a class="button_admin_1" onclick="javascript:CheckFields(); return false;"  href="">עדכון</a></td>  
     <td width="5%" align="center"><a class="button_admin_1" href="default.asp">חזרה לדף ארגונים</a></td>     
     <td width="5%" align="center"><a class="button_admin_1" href="../../nrerbiz/choose.asp">חזרה לדף ניהול ראשי</a></td>  
     <td width="*%" align="center"></td> 
  </tr>
</table>
<FORM name="frmMain" ACTION="addOrg.asp?ORGANIZATION_ID=<%=OrgID%>" METHOD="post" 
onSubmit="return CheckFields()" ID="frmMain" style="margin: 0px; padding:0px;" >
<table align=center border="0" cellpadding="3" cellspacing="1" width="100%">
<tr><td colspan="2" height="5"></td></tr>
<%
langId = Request.QueryString("langId")

if OrgID<>nil and OrgID<>"" then
  sqlStr = "SELECT ORGANIZATION_NAME,WORK_PRICING,SURVEYS,EMAILS,COMPANIES,TASKS,CASH_FLOW,Email_Limit, Sms_Limit, "&_
  " RowsInList,LANG_ID,isNULL(FILE_UPLOAD,0) as FILE_UPLOAD,isNULL(UseBizLogo,0) as UseBizLogo,"&_
  " Email_Limit_Month, Date_Add, isNULL(UseCaptcha,0) as UseCaptcha,SMS_TEXT,SMS_Phone From ORGANIZATIONS where ORGANIZATION_ID=" & OrgID  
  ''Response.Write sqlStr
  set rs_Org = con.GetRecordSet(sqlStr)  
	if not rs_Org.eof then
		ORGANIZATION_NAME = rs_Org("ORGANIZATION_NAME")		
		WORK_PRICING = rs_Org("WORK_PRICING")
		SURVEYS = rs_Org("SURVEYS")
		EMAILS = rs_Org("EMAILS")
		COMPANIES = rs_Org("COMPANIES")
		TASKS = rs_Org("TASKS")
		CASH_FLOW = rs_Org("CASH_FLOW")			
		Email_Limit = rs_Org("Email_Limit")
		SmsLimit = rs_Org("Sms_Limit")
		RowsInList = rs_Org("RowsInList")
		LANG_ID = rs_Org("LANG_ID")
		FILE_UPLOAD = rs_Org("FILE_UPLOAD")
		UseBizLogo = rs_Org("UseBizLogo")
		Email_Limit_Month = rs_Org("Email_Limit_Month")
		Date_Add = rs_Org("Date_Add")
		UseCaptcha =  rs_Org("UseCaptcha")
		SMSText=rs_Org("SMS_TEXT")
		SMSPhone=rs_Org("SMS_Phone")
		
	end if
	set rs_Org = nothing
else
		RowsInList = 30	
		Email_Limit_Month = 1000
		Email_Limit = 1000
		SmsLimit = 0
		Date_Add = Date()
		LANG_ID = langID
		UseBizLogo = "1" : UseCaptcha = "0"
end if

If isNumeric(trim(OrgID)) Then
	sqlpg="SELECT ORGANIZATION_LOGO FROM ORGANIZATIONS WHERE ORGANIZATION_ID="&OrgID&""
	'Response.Write sqlpg
	'Response.End
	set pg=con.getRecordSet(sqlpg)	
	perSize=pg.Fields("ORGANIZATION_LOGO").ActualSize
	set pg = Nothing
End If
	
If Request.QueryString("delScreen") <> nil Then
	con.executeQuery("UPdate ORGANIZATIONS set ORGANIZATION_LOGO = null where ORGANIZATION_ID = " & OrgID)
	Response.Redirect "addOrg.asp?ORGANIZATION_ID=" & OrgID
End If
%>
<tr><td colspan="4" bgcolor="#808080" height="1"></td></tr>
<tr>
	<td align=right bgcolor="#E6E6E6" colspan=2>
	<select dir="rtl" name="LANG_ID" style="width:250px;" <%If trim(OrgID) <> "" Then%> disabled <%Else%> onChange="document.location.href='addOrg.asp?langId='+this.value" <%End If%>>
	<option value=1 <%If trim(LANG_ID) = "1" Then%> selected <%End If%>>עברית</option>
	<option value=2 <%If trim(LANG_ID) = "2" Then%> selected <%End If%>>אנגלית</option></select></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" colspan=2>שפת האתר&nbsp;</td>
</tr>
<tr>
	<td align=right class="10normalB" bgcolor="#E6E6E6" colspan=2><input type=text dir="rtl" name="ORGANIZATION_NAME" value="<%=vfix(ORGANIZATION_NAME)%>" size="50" style="width:250px;"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" colspan=2>שם ארגון&nbsp;</td>
</tr>
<tr>
	<td align=right class="10normalB" bgcolor="#E6E6E6" colspan=2><input type=text dir="rtl" name="Email_Limit" value="<%=vfix(Email_Limit)%>" style="width:250px;" maxlength=10 onKeyPress="GetNumbers()"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" colspan=2>יתרת מיילים נוכחית&nbsp;</td>
</tr>
<tr>
	<td align=right class="10normalB" bgcolor="#E6E6E6" colspan=2><input type=text dir="rtl" name="Email_Limit_Month" value="<%=vfix(Email_Limit_Month)%>" maxlength="10" style="width:250px;" onKeyPress="GetNumbers()" ID="Email_Limit_Month"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" colspan=2>יתרת מיילים חודשית&nbsp;</td>
</tr>
<tr>
	<td align=right class="10normalB" bgcolor="#E6E6E6" colspan=2><font color="red">יתרה נוכחית מתעדכנת כל חודש ההחל מתאריך זה</font>
	<input type=image src='../../netcom/images/calend.gif' onclick='return popupcal(this.form.elements("Date_Add"));'>&nbsp;<input type=text dir="rtl" name="Date_Add" id="Date_Add" value="<%=vFix(Date_Add)%>" style="width:80px;" ReadOnly onclick="return popupcal(this);">
	</td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" colspan=2>תאריך הצטרפות&nbsp;</td>
</tr>
<tr>
	<td align=right class="10normalB" bgcolor="#E6E6E6" colspan=2><input type=text dir="rtl" name="SmsLimit" value="<%=vFix(SmsLimit)%>" size="50" style="width:250px;" onKeyPress="GetNumbers()" ID="SmsLimit"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" colspan="2" dir="rtl">&nbsp;יתרת SMS&nbsp;</td>
</tr>
<tr>
	<td align=right class="10normalB" bgcolor="#E6E6E6" colspan=2><input type=text dir="rtl" name="SmsText" value="<%=vFix(SmsText)%>" size="10" style="width:50px;" onKeyPress="GetNumbers()" ID="SmsText"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" colspan="2" dir="rtl">&nbsp;טקסט ההודעה - (מספר תווים)  &nbsp;</td>
</tr>
<tr>
	<td align=right class="10normalB" bgcolor="#E6E6E6" colspan=2><input type=text dir="rtl" name="SmsPhone" value="<%=vFix(SmsPhone)%>" size="10" style="width:80px;" onKeyPress="GetNumbers()" ID="SmsPhone"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" colspan="2" dir="rtl">&nbsp; טלפון ברירת מחדל להודעות  SMS   &nbsp;</td>
</tr>

<tr>
	<td align=right class="10normalB" bgcolor="#E6E6E6" colspan=2><input type=text dir="rtl" name="RowsInList" value="<%=vFix(RowsInList)%>" size="50" style="width:250px;" onKeyPress="GetNumbers()" ID="Text1"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" colspan=2>מספר שורות בדף&nbsp;</td>
</tr>
<tr>
	<td align=right class="10normalB" bgcolor="#E6E6E6" colspan=2><input type=checkbox dir="ltr" name="FILE_UPLOAD" <%If trim(FILE_UPLOAD) = "1" Then%> checked <%End If%> ID="Checkbox1"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" colspan=2>מורשה לצרף קבצים לטפסים&nbsp;&nbsp;</td>
</tr>
<tr>
	<td align=right class="10normalB" bgcolor="#E6E6E6" colspan=2><input type=checkbox dir="ltr" name="UseBizLogo" <%If trim(UseBizLogo) = "1" Then%> checked <%End If%> ID="Checkbox2"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" dir="rtl" nowrap width="220" colspan=2>&nbsp;&nbsp;הצג לוגו Bizpower&nbsp;&nbsp;</td>
</tr>
<tr>
	<td align=right class="10normalB" bgcolor="#E6E6E6" colspan=2><input type=checkbox dir="ltr" name="UseCaptcha" <%If trim(UseCaptcha) = "1" Then%> checked <%End If%> ID="UseCaptcha"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" dir="rtl" nowrap width="220" colspan=2>&nbsp;&nbsp;אימות חתימה בטפסים&nbsp;&nbsp;</td>
</tr>
<tr><td colspan="4" bgcolor="#808080" height="1"></td></tr>
<tr><td class="11normalB" align=right bgcolor="#E6E6E6" colspan=4>הגדרת תפריט והרשאות&nbsp;</td></tr>
<% 
If trim(langId) = "2" Then
	lang_ = "_eng"
Else
	lang_ = ""
End If

sql_obj="Select dbo.bars.bar_id, dbo.bars.bar_title, dbo.bars.parent_id, bars_1.bar_order AS Parent_Order " &_
" From dbo.bars LEFT OUTER JOIN dbo.bars bars_1 ON dbo.bars.parent_id = bars_1.bar_id " &_
" WHERE dbo.bars.parent_id IS NOT NULL And dbo.bars.in_use = 1 Order by Parent_Order, dbo.bars.bar_Order"
set rs_obj=con.getRecordSet(sql_obj)
If not rs_obj.eof Then		
	arr_bars_d1 = Split(rs_obj.Getstring(,,",",";") , ";")
End If
set rs_obj = Nothing

If OrgID<>nil and OrgID<>"" then
    sql_obj="Select bar_title From dbo.bar_organizations_table('"&OrgID&"') WHERE PARENT_ID IS NOT NULL Order by Parent_Order, bar_Order"
   ' response.Write sql_obj
  'response.end
    set rs_obj=con.getRecordSet(sql_obj)
	If not rs_obj.eof Then		
		arr_bars1 = Split(rs_obj.Getstring(,,",",";") , ";")
	Else
		arr_bars1 = array()
		redim arr_bars1(Ubound(arr_bars_d1))
	End If
	set rs_obj = Nothing
	
	sql_obj="Select is_visible From dbo.bar_organizations_table('"&OrgID&"') WHERE PARENT_ID IS NOT NULL Order by Parent_Order, bar_Order"
    set rs_obj=con.getRecordSet(sql_obj)
	If not rs_obj.eof Then		
		arr_visible1 = Split(rs_obj.Getstring(,,",",";") , ";")
	Else
		arr_visible1 = array()
		redim arr_visible1(Ubound(arr_bars_d1))
	End If
	set rs_obj = Nothing
Else
	sql_obj="Select bar_title"&trim(lang_)&" From bars WHERE PARENT_ID IS NOT NULL "&_
	" And in_use = 1 Order By (Select barsp.bar_order From dbo.bars barsp Where "&_
	" barsp.bar_id = dbo.bars.parent_id), bar_Order"
    set rs_obj=con.getRecordSet(sql_obj)
	If not rs_obj.eof Then		
		arr_bars1 = Split(rs_obj.Getstring(,,",",";") , ";")
	Else
		arr_bars1 = array()
		redim arr_bars1(Ubound(arr_bars_d1))
	End If
	set rs_obj = Nothing
	
	arr_visible1 = array()
	redim arr_visible1(Ubound(arr_bars_d1))
	For i=0 To Ubound(arr_visible1)
		arr_visible1(i) = 1
	Next	
End If
%>

<%For i=0 To Ubound(arr_bars_d1)-1
    arr_bar = Split(arr_bars_d1(i),",")
    barID = arr_bar(0)
    barTitle = arr_bar(1)
    parent_id = arr_bar(2)  
	If trim(barID) = "43" Then
		barTitle = barTitle & " <font color=red>(הרשאות טפסים)</font>"
		bgcolor = "#FFBFBF"
	ElseIf trim(barID) = "10" Then
		barTitle = barTitle & ""
		bgcolor = "#FFBFBF"	
	Else
		bgcolor = "#E6E6E6"
	End If      
    If isNumeric(parent_id) Then
		If OrgID<>nil and OrgID<>"" then
			sql_obj="Select bar_title, is_visible  from BAR_ORGANIZATIONS WHERE bar_id = " & parent_id & " AND ORGANIZATION_ID="&OrgID
			'Response.Write sql_obj
			'Response.End
			set rs_obj=con.getRecordSet(sql_obj)
			If not rs_obj.eof Then		
				parentTitle = rs_obj(0)
				parentVisible = rs_obj(1)	
			End If
			set rs_obj = Nothing	
		Else
			sql_obj="Select bar_title"&lang_&" from bars WHERE bar_id = " & parent_id
			'Response.Write sql_obj
			'Response.End
			set rs_obj=con.getRecordSet(sql_obj)
			If not rs_obj.eof Then		
				parentTitle = rs_obj(0)
				parentVisible = 1
			End If
			set rs_obj = Nothing	
		End If
	If old_parent <> parentTitle Then	

	sql_obj="Select bar_title from bars WHERE parent_id IS NULL AND bar_id = " & parent_id
	'Response.Write sql_obj
	set rs_obj=con.getRecordSet(sql_obj)
	If not rs_obj.eof Then		
		parentTitleD = rs_obj(0)			
	End If
	set rs_obj = Nothing		
%>
<%If old_parent <> parentTitle Then%>
<tr><td colspan="4" bgcolor="#808080" height="1"></td></tr>
<%End If%>
<tr bgcolor="#C0C0C0">
	<td align=right width=100%><input type=checkbox dir="ltr" name="is_visible<%=parent_id%>" <%If trim(parentVisible) = "1" Then%> checked <%End If%> ID="is_visible<%=parent_id%>" onclick="return check_all_bars(this,'<%=parent_id%>')" ></td>
	<td class="10normalB" align=right width=300 nowrap><input type=text dir="rtl" name="title_bar<%=parent_id%>" value="<%=parentTitle%>" style="width:290px;" ID="title_bar<%=parent_id%>"></td>
	<th class="11normalB" align=right nowrap><%=parentTitleD%>&nbsp;</th>
</tr>
<%
	End If	
	End If
%>
<tr bgcolor="<%=bgcolor%>">
	<td class="10normalB" align="right" dir="rtl" width="100%"><input type=checkbox dir="ltr" name="is_visible<%=barID%>" id="<%=barID%>!<%=parent_id%>" <%If trim(arr_visible1(i)) = "1" And Not(trim(OrgID) = "" And trim(barID) = "43") Then%> checked <%End If%> ID="is_visible<%=barID%>">
	 <%If trim(barID) = "10" Then%><font color=red>(שימו לב: בארגונים ללא מודול חברות לא ניתן להזין אנשי קשר עם שיוך לחברה. מודול זה מתאים לארגונים שעובדים עם לקוחות פרטיים בלבד!)</font><%End If%>
	</td>
	<td class="10normalB" align=right width=300 nowrap><input type=text dir="rtl" name="title_bar<%=barID%>" value="<%=arr_bars1(i)%>" style="width:290px;" ID="title_bar<%=barID%>"></td>
	<th class="10normalB" align=right dir=rtl>&nbsp;&nbsp;&nbsp;<%=barTitle%></th>
</tr>
<% old_parent = parentTitle %>
<%Next%>
<tr><td colspan="4" bgcolor="#808080" height="1"></td></tr>
<% 
if OrgID<>nil and OrgID<>"" then
   sql_obj="Select * from view_objects WHERE ORGANIZATION_ID="&OrgID &" order by title"
   set rs_obj=con.getRecordSet(sql_obj)
else
    sql_obj="Select * from objects order by title"
end if
  set rs_obj=con.getRecordSet(sql_obj) 
%>
<tr><td class="11normalB"  align=right bgcolor="#E6E6E6" colspan=4>הגדרת אוביקטים&nbsp;</td></tr>
<tr>
<td align=right bgcolor="#E6E6E6" class="10normalB">רבים</td>
<td align=right bgcolor="#E6E6E6" class="10normalB">יחיד</td>
<td align=right bgcolor="#DDDDDD" colspan=2 class="10normalB">שם אוביקט</td>
</tr>
<%if not rs_obj.eof then 
		do while not rs_obj.eof
			%>
<tr>
<td align=right bgcolor="#DDDDDD" class="10normalB"><input type=text dir="rtl" name=title_multi<%=rs_obj("object_ID")%> value="<%=vfix(rs_obj("title_organization_multi"&lang_))%>" size="50" style="width:250px;" ></td>
<td align=right bgcolor="#DDDDDD" class="10normalB"><input type=text dir="rtl" name=title_one<%=rs_obj("object_ID")%> value="<%=vfix(rs_obj("title_organization_one"&lang_))%>" size="50" style="width:250px;" ></td>
<td align=right bgcolor="#DDDDDD" colspan=2 class="normalB"><%=trim(rs_obj("title"))%></td>
</tr>
<%	rs_obj.MoveNext
		loop
	set rs_obj=nothing%>
	<%end if%>
<tr><td colspan="4" bgcolor="#808080" height="1"></td></tr></table>
</form>
<% If trim(OrgID) <> "" Then %>
<form ACTION="Aimgadd.asp?C=1&F=ORGANIZATION_LOGO&ORGANIZATION_ID=<%=OrgID%>" METHOD="post" 
ENCTYPE="multipart/form-data" ID="Form1" style="margin: 0px; padding:0px;" >
<table cellpadding=0 cellspacing=0 width=100%>
<tr><td colspan=2 align=right class="10normalB" bgcolor="#DDDDDD">pxרוחב מקסימלי  350 X pxגודל מומלץ: גובה 60 </td></tr>
		<%If perSize = 0 then %>
		<tr>
			<td align="right" width="530" nowrap bgcolor="#dbdbdb">&nbsp;</td>
			<td align="center" width="100" nowrap rowspan="2" class="10normalB" bgcolor="#DDDDDD"><b>Logo</b>
			</td>
		</tr>
		<tr valign=middle>
			<td align="right" bgcolor="#dbdbdb" valign="middle">
			<table border="0" cellspacing="0" cellpadding="0" align="right">
				<tr valign=middle>        
				<td align="right" class="td_admin_1" style="padding-right:2px;"><INPUT type="submit" class="but_browse" onclick="return ifFieldEmpty('UploadFile2')" value="Logo העלאת"></td>
				<td align="right" class="td_admin_1" style="padding-right:2px;"><INPUT TYPE="FILE" NAME="UploadFile2" size=30 ID="File1"></td>
				</tr>
			</table>       
			</td>
		</tr>
		<tr><td align="right" colspan=2 nowrap bgcolor="#dbdbdb">&nbsp;</td></tr>
		<%Else%>
		<tr>
			<td align=right  bgcolor="#dbdbdb"><img id="imgPict" name="imgPict" src="../../Netcom/GetImage.asp?DB=ORGANIZATION&amp;FIELD=ORGANIZATION_LOGO&amp;ID=<%=OrgID%>" border="0" hspace=2 ></td>						
			<td align=center rowspan="2" bgcolor="#b0b0b0" nowrap valign=top><font style="font-size:10pt;color:#ffffff;"><b>Logo</b></font></td>
		</tr>
		<tr>
			<td align=right  bgcolor="#b0b0b0" >
			<table border="0" cellspacing="0" cellpadding="0" align="right" ID="Table2">
				<tr valign=middle>        
				<td align="right" class="td_admin_1" style="padding-right:2px;"><INPUT type="button" class="but_browse"  ONCLICK="return CheckDel('delScreen','<%=OrgID%>');" value="Logo מחיקת " id=button1 name=button1></td>
				<td align="right" class="td_admin_1" style="padding-right:2px;"><INPUT type="submit" class="but_browse" onclick="return ifFieldEmpty('UploadFile2')" value="Logo החלפת"></td>
				<td align="right" class="td_admin_1" style="padding-right:2px;"><INPUT TYPE="FILE" NAME="UploadFile2" size=30 ID="File2"></td>
				</tr>
			</table>
			</td>
		</tr>
		<%End If%>		
</table>
</form>		
<%End If%>
</div>
</body>
</html>
<%set con = nothing%>
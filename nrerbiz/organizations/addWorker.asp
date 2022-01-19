<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
</head>

<script LANGUAGE="JavaScript">
<!--
function CheckFields()
{
	if ( document.frmMain.FIRSTNAME.value=='' ||
	     document.frmMain.LOGINNAME.value=='' || 
	     document.frmMain.PASSWORD.value==''  ||
	     document.frmMain.JOB_ID.value==''  ||
	     document.frmMain.email.value==''
	     )		   
		{
			window.alert("'*' אנא מלאו כל השדות הצוינות בסימן")
			return false;				
		}
	else if (!checkEmail(document.all("email").value) && document.all("email").value != "")
		{
          	window.alert("כתובת דואר אלקטרוני לא חוקית");
			document.all("email").focus();
			return false;
		}				
	return true;
	
}

	function GetNumbers ()
	{
		var ch=event.keyCode;
		event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
	} 
	
	function changeJob(selObj)
	{
		if(selObj.value == "")
		{
			selObj.focus();
			window.alert("נא לבחור תפקיד");
			return false;
		}
		
		selVal = new String(selObj.value);
		selVal = selVal.split("!");
		window.document.all("Month_Min_Order").value = selVal[1];
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
	
	function checkEmail(addr)
	{
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
	
	function CheckResp(selObject)
	{
		var chkGrps = selObject.value;
		if(chkGrps != "")
			window.document.all("resp_edit").style.display = "inline";
		else
			window.document.all("resp_edit").style.display = "none";
	}

//-->
</script>  

<%
OrgID=request.querystring("ORGANIZATION_ID")
USER_ID=request.querystring("USER_ID")

if OrgID<>nil and OrgID<>"" then
	sqlStr = "select ORGANIZATION_NAME,isNULL(Voice_Operation,0) as Voice_Operation From ORGANIZATIONS where ORGANIZATION_ID=" & OrgID  
	'Response.Write sqlStr
	'Response.End
	set rs_Org = con.GetRecordSet(sqlStr)  
	if not rs_Org.eof then
		ORGANIZATION_NAME = rs_Org("ORGANIZATION_NAME")
		Voice_Operation = rs_Org("Voice_Operation")
	end if
	set rs_Org = nothing
  
  	sqlStr = "Select Top 1 Bar_Id from bar_organizations WHERE Bar_ID = 43 And IsNull(is_visible,0) = 1 and ORGANIZATION_ID= "& OrgID
	'Response.Write sqlStr
	set rs_groups = con.GetRecordSet(sqlStr)
	If not rs_groups.eof Then
		is_groups = "1"
	Else
		is_groups = "0"	
	End If
	set rs_groups = Nothing		
end if

errLogin=false
errEmail=false

if request.form("FIRSTNAME")<>nil then 'after form filling
	 	 
	 FIRSTNAME = sFix(request.form("FIRSTNAME"))
	 LASTNAME  = sFix(trim(Request.Form("LASTNAME")))
	 LOGINNAME = sFix(request.form("LOGINNAME"))
	 PASSWORD  = sFix(request.form("PASSWORD"))
     'office    = sFix(trim(Request.Form("office")))
     phone     = sFix(trim(Request.Form("phone")))
     mobile    = sFix(trim(Request.Form("mobile")))
     fax       = sFix(trim(Request.Form("fax")))
     email     = sFix(trim(Request.Form("email")))
     Month_Min_Order  = sFix(trim(Request.Form("Month_Min_Order")))
     If Request.Form("EDIT_APPEAL") <> nil Then
		EDIT_APPEAL = sFix(trim(Request.Form("EDIT_APPEAL")))     
     Else
		EDIT_APPEAL = "0"
     End If
     job_pay    = trim(Request.Form("job_id"))
     job_pay = Split(job_pay,"!")
     If IsArray(job_pay) Then
		job_id = job_pay(0)
	 Else
		job_id = "NULL"
	 End If	 
     
     If Request.Form("chief") <> nil Then
		chief = "1"  
     Else
		chief = "0"
     End If    
     
     If Request.Form("Add_Meetings") <> nil Then
		Add_Meetings = "1"  
     Else
		Add_Meetings = "0"
     End If     
       If Request.Form("Add_Insurance") <> nil Then
		Add_Insurance = "1"  
     Else
		Add_Insurance = "0"
     End If     

     COMPANIES = 0 : SURVEYS = 0 : EMAILS = 0 : WORK_PRICING = 0 : PROPERTIES = 0 : TASKS = 0 : CASH_FLOW = 0
	 if USER_ID=nil or USER_ID="" then 'new record in DataBase
		    sqlStr = "select LOGINNAME from USERS where LOGINNAME = '" & sFix(LOGINNAME) & "'"
		    set rs_LOGINNAME = con.GetRecordSet(sqlStr)
		    if  not rs_LOGINNAME.EOF then
		        errLogin=true
		    end if
		    set rs_LOGINNAME = nothing
		    sqlStr = "select EMAIL from USERS where EMAIL='" & email & "' AND ORGANIZATION_ID = " & OrgID
		    set rs_EMAIL = con.GetRecordSet(sqlStr)
		    if  not rs_EMAIL.EOF then
		        errEmail=true
		    end if
		    set rs_EMAIL = Nothing
		    'Response.Write ("<BR>errLogin="& errLogin)
		    'Response.Write ("<BR>errEmail="& errEmail)
		    'Response.End
		    if errLogin=false AND errEmail=false  then
				sqlStr = "SET NOCOUNT ON; Insert Into USERS (FIRSTNAME,LASTNAME,LOGINNAME,PASSWORD,ACTIVE,"&_
				"ORGANIZATION_ID,job_id,TELEPHONE,MOBILE,FAX,EMAIL,Month_Min_Order,WORK_PRICING,SURVEYS,EMAILS,"&_
				"COMPANIES,TASKS,CASH_FLOW,CHIEF,EDIT_APPEAL,Voice_Operator,Add_Meetings,Add_Insurance) " &_
				" values ('" & FIRSTNAME &"','"& LASTNAME &"','"& LOGINNAME &"','"& PASSWORD &"',1,"& OrgID &","&_
				job_id &",'"& phone &"','"& mobile &"','"& fax &"','"& email &"','" & Month_Min_Order & "','" &_
				WORK_PRICING& "','" &SURVEYS& "','" &EMAILS& "','" &COMPANIES& "','" & TASKS & "','" &_
				CASH_FLOW & "'," & CHIEF & ",'" & EDIT_APPEAL & "','" & Voice_Operator & "','" & Add_Meetings & "','"& Add_Insurance &"'); SELECT @@IDENTITY AS NewID"
				'Response.Write sqlStr
				'Response.End
				set rs_tmp = con.getRecordSet(sqlStr)
					USER_ID = rs_tmp.Fields("NewID").value
				set rs_tmp = Nothing							  							  
			 elseif errLogin=true then
				 LOGINNAME = "" %>
			     <SCRIPT LANGUAGE=javascript>
			     <!--
				   alert('שם משתמש זה בשימוש, בחר שם משתמש אחר')
				   history.back();
			      //-->
			     </SCRIPT>			 			 
			 <%Response.End
			   elseif errEmail=true then
				  email=""%>
			      <SCRIPT LANGUAGE=javascript>
			      <!--
				     alert('כתובת דואר אלקטרוני שציינת'+'\n'+'רשומה כבר במערכת')
				     history.back();
			      //-->
			      </SCRIPT>			 			 				
			  <%Response.End
			    end if			
		    
	 else			   
				 sqlStr = "select LOGINNAME from USERS where USER_ID <> "& USER_ID &" and LOGINNAME = '" & sFix(LOGINNAME) & "'"
				 set rs_LOGINNAME = con.GetRecordSet(sqlStr)
				 if not rs_LOGINNAME.EOF then
		             errLogin=true
		         end if
		         set rs_LOGINNAME = nothing
		    
				 sqlStr = "select EMAIL from USERS where USER_ID <> "& USER_ID &" and EMAIL='" & email & "' AND ORGANIZATION_ID = " & OrgID
		         set rs_EMAIL = con.GetRecordSet(sqlStr)
		         if not rs_EMAIL.EOF then
		             errEmail=true
		         end if
		         set rs_EMAIL = Nothing
				 
				 if errLogin=false AND errEmail=false  then
					sqlStr = "UPDATE USERS set FIRSTNAME='" & FIRSTNAME &"',LASTNAME='"& LASTNAME &_
					"', LOGINNAME='"& LOGINNAME &"', PASSWORD = '"& PASSWORD &"',job_id="& job_id &_
					",  TELEPHONE='"& phone &"',MOBILE='"& mobile &"',FAX='"& fax &"',EMAIL='"& email &_
					"', WORK_PRICING='" & WORK_PRICING & "',SURVEYS='" & SURVEYS & "',EMAILS='" & EMAILS &_
					"', COMPANIES='" & COMPANIES & "', TASKS='" & TASKS & "', CASH_FLOW='" & CASH_FLOW &_
					"', Month_Min_Order='" & Month_Min_Order & "', CHIEF = " & CHIEF & ", EDIT_APPEAL = '" & EDIT_APPEAL &_
					"', Voice_Operator='" & Voice_Operator & "', Add_Meetings = '" & Add_Meetings & "' , Add_Insurance = '" & Add_Insurance & "'  WHERE USER_ID=" & USER_ID
					'Response.Write sqlStr
					'Response.End 
					con.GetRecordSet (sqlStr)					     
				 elseif errLogin=true then
					LOGINNAME = ""%>
				    <SCRIPT LANGUAGE=javascript>
				     <!--
					   alert('שם משתמש זה בשימוש, בחר שם משתמש אחר')
					   history.back();
				     //-->
				   </SCRIPT>			 			 
				 <%Response.End
				   elseif errEmail=true then
				   email=""%>
			       <SCRIPT LANGUAGE=javascript>
			       <!--
						alert('כתובת דואר אלקטרוני שציינת'+'\n'+'רשומה כבר במערכת') 
						history.back();
			       //-->
			       </SCRIPT>			 			 				
			    <%Response.End
			      end if			    
    end if
	sqlstr = "Delete FROM bar_users WHERE organization_id = " & OrgID & " AND USER_ID = " & USER_ID
	'Response.Write sqlstr
	'Response.End
	con.executeQuery(sqlstr)
			
	sqlstr="Select bar_id from bars Order by bar_id"
	set rs_bars = con.getRecordSet(sqlstr)
	While not rs_bars.eof
		barID = trim(rs_bars(0))
		barVisible = Request.Form("is_visible"&barID)
		If trim(barVisible) = "on" Then
			barVisible = 1
		Else
			barVisible = 0
		End If	
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
		sqlstr = "Insert Into bar_users values (" & barID & "," & OrgID & "," & USER_ID & ",'" & barVisible & "')"
		con.executeQuery(sqlstr)
	rs_bars.moveNext
	Wend
	set rs_bars = Nothing   
	
	sqlStr = "UPDATE USERS SET WORK_PRICING='"&WORK_PRICING&"',SURVEYS='"&SURVEYS&"',EMAILS='"&EMAILS&_
	"', COMPANIES='"&COMPANIES&"', TASKS='"&TASKS&"', CASH_FLOW='"&CASH_FLOW&"', EDIT_APPEAL = '" & EDIT_APPEAL &_
	"'  WHERE USER_ID="&USER_ID
	'Response.Write sqlStr
	'Response.End
	con.GetRecordSet (sqlStr)
	
	'<--------------------------------------- שיוך לקבוצות ------------------------------------------------------>
	sqlstr="Delete FROM Users_to_Groups WHERE USER_ID="&USER_ID
	con.executeQuery(sqlstr)
	
	groups = trim(Request.Form("group_id")) & ","
	groups = Split(groups, ",")
	If IsArray(groups) And Ubound(groups) > 0 Then
		For i=0 To Ubound(groups)
			If trim(groups(i)) <> "" And IsNumeric(groups(i)) = true Then
				sqlstr="Insert Into Users_to_Groups (USER_ID,group_id) values ("&USER_ID&","&groups(i)&" )"
				con.executeQuery(sqlstr)
			End If 
		Next
	End If
	'<--------------------------------------- שיוך לקבוצות ------------------------------------------------------>
	
	'<--------------------------------------- אחראי בקבוצות ------------------------------------------------------>
	sqlstr="Delete FROM Responsibles_to_Groups WHERE Responsible_ID="&USER_ID
	con.executeQuery(sqlstr)
	
	r_groups = trim(Request.Form("R_GROUP_ID")) & ","
	r_groups = Split(r_groups, ",")
	If IsArray(r_groups) And Ubound(r_groups) > 0 Then
		For i=0 To Ubound(r_groups)
			If trim(r_groups(i)) <> "" And IsNumeric(r_groups(i)) = true Then
				sqlstr="Insert Into Responsibles_to_Groups (Responsible_ID,GROUP_ID) values ("&USER_ID&","&r_groups(i)&" )"
				con.executeQuery(sqlstr)
			End If 
		Next
	End If
	'<--------------------------------------- אחראי בקבוצות ------------------------------------------------------>
	'<----------------------------------------הוספת הראשת טפסים לקבוצות שנוספו ----------------------------------->
	If Request.Form("group_id") <> nil Then
	sqlstr = "Select DISTINCT Product_ID, Group_ID From Users_To_Products WHERE Organization_ID = " & OrgID & " And Group_ID IN (" & trim(Request.Form("group_id")) & ")"
	set rs_products = con.getRecordSet(sqlstr)
	while not rs_products.eof
	    sqlCheck = "Select Top 1 Product_ID From Users_To_Products WHERE User_ID = " & USER_ID &_
	    " And Product_ID = " & rs_products(0) & " And Group_ID = " & rs_products(1)
	    set rs_check = con.getRecordSet(sqlCheck)
	    If rs_check.eof Then
	    sqlInsert = "Insert Into Users_To_Products values (" & USER_ID & "," & rs_products(1) & "," & OrgID & "," & rs_products(0) & ")"
	    con.executeQuery(sqlInsert)
	    End If
	    set rs_check = Nothing
		rs_products.moveNext
	wend
	set rs_products = Nothing
	End If
	'<----------------------------------------הוספת הראשת טפסים לקבוצות שנוספו ----------------------------------->
	'<----------------------------------------למחוק הרשאה לטפסים מקבוצות שנמחקו ---------------------------------->
	sqlDelete = "Delete FROM Users_To_Products WHERE  User_ID = " & USER_ID &_
	" And Group_ID Not IN (Select Group_ID FROM Users_to_Groups WHERE USER_ID="&USER_ID&")"
	con.executeQuery(sqlDelete)
	'<----------------------------------------למחוק הרשאה לטפסים מקבוצות שנמחקו ---------------------------------->	
	%>
	<script language=javascript>
	<!--
		document.location.href = "workers.asp?ORGANIZATION_ID=<%=OrgID%>";
	//-->
	</script>
	<%				 
end if
%>
<body bgcolor="#FFFFFF" <%If trim(is_groups) = "1" Then%> onload="CheckResp(document.all('R_GROUP_ID'))" <%End If%>>
<div align="right">
<FORM name="frmMain" ACTION="addWorker.asp?USER_ID=<%=USER_ID%>&ORGANIZATION_ID=<%=OrgID%>" METHOD="post" onSubmit="return CheckFields()" ID="Form1">
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="6" class="page_title" dir=rtl><%=ORGANIZATION_NAME%>&nbsp;-&nbsp;<%if USER_ID<>nil then%>עדכון<%else%>הוספת<%end if%>&nbsp;עובד</td>
  </tr>
  <tr>
     <td width="5%" align="center"><input type=submit class="button_admin_1" value="עדכון"></td>       
     <td width="5%" align="center"><a class="button_admin_1" href="workers.asp?ORGANIZATION_ID=<%=OrgID%>" target=_self>חזרה לדף עובדים</a></td>          
     <td width="5%" align="center"><a class="button_admin_1" href="default.asp" target=_self>חזרה לדף ארגונים</a></td>     
     <td width="5%" align="center"><a class="button_admin_1" href="../../nrerbiz/choose.asp" target=_self>חזרה לדף ניהול ראשי</a></td>  
     <td width="*%" align="center"></td> 
  </tr>
</table> 
<table align=center border="0" cellpadding="3" cellspacing="1" width="100%">
<tr><td colspan="4" height="15"></td></tr>
<%
	if USER_ID<>nil and USER_ID<>"" then
	sqlStr = "select FIRSTNAME,LASTNAME, LOGINNAME,PASSWORD,OFFICE,TELEPHONE,MOBILE,FAX,EMAIL,Month_Min_Order,chief, "&_
	" job_id,WORK_PRICING,SURVEYS,EMAILS,COMPANIES,TASKS,CASH_FLOW,CHIEF,ISNULL(EDIT_APPEAL,0),"&_
	" isNULL(Voice_Operator,0),IsNUll(Add_Meetings,0),IsNUll(Add_Insurance,0)   From USERS where USER_ID=" & USER_ID  
	''Response.Write sqlStr
	set rs_USERS = con.GetRecordSet(sqlStr)
		if not rs_USERS.eof then
			FIRSTNAME   = rs_USERS(0)
			LASTNAME    = rs_USERS(1)
			LOGINNAME   = rs_USERS(2)	
			PASSWORD_U  = rs_USERS(3)					
			office      = rs_USERS(4)
			phone       = rs_USERS(5)
			mobile      = rs_USERS(6)
			fax         = rs_USERS(7)
			email       = rs_USERS(8)
			Month_Min_Order  = rs_USERS(9)
			chief       = rs_USERS(10) 
			job_id    = rs_USERS(11)
			WORK_PRICING = rs_USERS(12)
			SURVEYS = rs_USERS(13)
			EMAILS = rs_USERS(14)
			COMPANIES = rs_USERS(15)
			TASKS = rs_USERS(16)
			CASH_FLOW = rs_USERS(17)
			CHIEF = rs_USERS(18)	
			EDIT_APPEAL = rs_USERS(19) 	
			Voice_Operator = rs_USERS(20) 	
			Add_Meetings = rs_USERS(21) 	
			Add_Insurance = rs_USERS(22) 	
		end if
		set rs_USERS = nothing	
		
	Groups=""		
	sqlstr="Select Group_Name From Users_to_Groups_view Where User_id = " & USER_ID & " Order By Group_id"
	set rssub = con.getRecordSet(sqlstr)		   
	If not rssub.eof Then
		Groups = rssub.getString(,,",",",")		
	Else	
		Groups = ""
	End If		    
	set rssub=Nothing
	If Len(Groups) > 0 Then
		Groups = Left(Groups,(Len(Groups)-1))
	End If	
	
	ResInGroups=""		
	sqlstr="Select Group_Name From Responsibles_to_Groups_view Where Responsible_id = " & USER_ID & " Order By Group_id"
	set rssub = con.getRecordSet(sqlstr)		   
	If not rssub.eof Then
		ResInGroups = rssub.getString(,,",",",")	
	Else
		ResInGroups=""
	End If		    
	set rssub=Nothing
	If Len(ResInGroups) > 0 Then
		ResInGroups = Left(ResInGroups,(Len(ResInGroups)-1))
	End If			

end if
%>
<tr><td colspan="4" bgcolor="#808080" height="1"></td></tr>
<tr>
	<td align=right bgcolor="#E6E6E6" width=100%><input type=text dir="rtl" name="FIRSTNAME" value="<%=vfix(FIRSTNAME)%>" size="50" style="width:250px;font-family:arial"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220">שם פרטי&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6"><input type=text dir="rtl" name="LASTNAME" value="<%=vfix(LASTNAME)%>" size="50" style="width:250px;font-family:arial"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220">שם משפחה&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6"><input type=text dir="ltr" name="LOGINNAME" value="<%=vfix(LOGINNAME)%>" size="50" style="width:250px;font-family:arial"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220">שם משתמש&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6"><input type=text dir="ltr" name="PASSWORD" value="<%=vfix(PASSWORD_U)%>" size="50" style="width:250px;font-family:arial"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220">סיסמה&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6">
	<select dir="rtl" name="JOB_ID" class="norm" onchange="return changeJob(this)" ID="JOB_ID" style="width:250px;font-family:arial">
	<option><%=String(20,"-")%>בחר תפקיד<%=String(20,"-")%></option>
	<%
		sqlstr = "select job_id,job_name,hour_pay from jobs WHERE ORGANIZATION_ID = " & OrgID & " Order By job_id"
		set rs_jobs = con.getRecordSet(sqlstr)
		while not rs_jobs.eof
	%>
	<option value="<%=trim(rs_jobs(0))%>!<%=trim(rs_jobs(2))%>" <%If trim(job_id) = trim(rs_jobs(0)) Then%> selected <%End If%>><%=rs_jobs(1)%></option>
	<%
		rs_jobs.moveNext
		wend
		set rs_jobs = nothing
	%>
	</select>
	</td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220">תפקיד&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6"><input type=text dir="ltr" name="phone" value="<%=vfix(phone)%>" size="50" style="width:250px;font-family:arial"></td>
	<td align="right" bgcolor="#DDDDDD" class="10normalB" nowrap width="220">טלפון&nbsp;&nbsp;</td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6"><input type=text dir="ltr" name="mobile" value="<%=vfix(mobile)%>" size="50" style="width:250px;font-family:arial"></td>
	<td align="right" bgcolor="#DDDDDD"class="10normalB" nowrap width="220">נייד&nbsp;&nbsp;</td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6"><input type=text dir="ltr" name="fax" value="<%=vfix(fax)%>" size="50" style="width:250px;font-family:arial"></td>
	<td align="right" bgcolor="#DDDDDD" class="10normalB" nowrap width="220">פקס&nbsp;&nbsp;</td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6"><input type=text dir="ltr" name="email" value="<%=vfix(email)%>" size="50" style="width:250px;font-family:arial"></td>
	<td align="right" bgcolor="#DDDDDD" class="10normalB" nowrap width="220">דואר אלקטרוני&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6"><input type=text dir="ltr" name="Month_Min_Order" value="<%=vfix(Month_Min_Order)%>" size="50" style="width:250px;font-family:arial" onkeypress="GetNumbers();" ID="Text1"></td>
	<td align="right" bgcolor="#DDDDDD" class="10normalB" nowrap width="220">עלות שעה&nbsp;&nbsp;</td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6"><input type="checkbox" dir="ltr" name="chief" <%If trim(chief) = "1" Then%> checked <%End If%> ID="chief"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" >מורשה מחיקה&nbsp;&nbsp;</td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6"><input type="checkbox" dir="ltr" name="Add_Meetings" <%If trim(Add_Meetings) = "1" Then%> checked <%End If%> ID="Add_Meetings"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" >אחראי פגישות&nbsp;&nbsp;</td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6"><input type="checkbox" dir="ltr" name="Add_Insurance" <%If trim(Add_Insurance) = "1" Then%> checked <%End If%> ID="Add_Insurance"></td>
	<td align="right" class="10normalB" bgcolor="#DDDDDD" nowrap width="220" >"ביטוח מול "דיוויד שילד&nbsp;&nbsp;</td>
</tr>
<% If trim(is_groups) = "1" Then %>
<tr><td colspan="4" bgcolor="#808080" height="1"></td></tr>
<tr><td class="11normalB" align=right bgcolor="#E6E6E6" colspan=4>הרשאות טפסים&nbsp;</td></tr>
<tr><td colspan="4" bgcolor="#808080" height="1"></td></tr>
<tr>
	<td align=right bgcolor="#E6E6E6" valign=top>.העובד יוכל להזין את כל הטפסים המשויכים לקבוצות אלו
	<select dir="rtl" name="GROUP_ID" class="norm" ID="GROUP_ID" style="width:250" size=3 multiple>
	<option value=""  <%if trim(Groups) = "" then%> selected <%End If%>><%=String(15,"-")%> ללא שיוך לקבוצה <%=String(15,"-")%></option>
	<%
        If trim(USER_ID) <> "" Then
            set groupList=con.GetRecordSet("SELECT group_ID,group_Name FROM Users_to_Groups_view WHERE ORGANIZATION_ID = " & trim(OrgID) & " AND User_ID = "&USER_ID&" Order By group_Name")
            do while not groupList.EOF
                prgroupID=groupList("group_ID")
                prgroupName=groupList("group_Name")%>
                <option value="<%=trim(prgroupID)%>" <%if InStr(groups,prgroupName) > 0 then%> selected <%end if%>><%=prgroupName%></option>
            <%groupList.MoveNext
            loop
            groupList.close
            set groupList=Nothing
            End If
            If trim(USER_ID) <> "" Then
            set groupList=con.GetRecordSet("SELECT group_ID,group_Name FROM Users_Groups WHERE ORGANIZATION_ID = " & trim(OrgID) & " AND group_ID NOT IN ( Select group_ID From Users_to_Groups WHERE User_ID = "&USER_ID&") Order By group_Name")
            Else
            set groupList=con.GetRecordSet("SELECT group_ID,group_Name FROM Users_Groups WHERE ORGANIZATION_ID = " & trim(OrgID) & " Order By group_Name")
            End If
            do while not groupList.EOF
                prgroupID=groupList("group_ID")
                prgroupName=groupList("group_Name")%>
                <option value="<%=trim(prgroupID)%>" ><%=prgroupName%></option>
            <%groupList.MoveNext
            loop
            groupList.close
            set groupList=Nothing%>
	</select>
	</td>
	<td align="right" valign="top" class="10normalB" bgcolor="#DDDDDD" nowrap width="220">&nbsp;שייך לקבוצות&nbsp;&nbsp;</td>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6" valign=top>.העובד יוכל לצפות או לעדכן את כל הטפסים שהוזנו ע"י העובדים בקבוצות אלו 
	<select dir="rtl" name="R_GROUP_ID" class="norm" ID="R_GROUP_ID" style="width:250" size=3 multiple onchange="CheckResp(this)">
	<option value="" id=word18 name=word18 <%If Len(ResInGroups) = 0 Then%> selected <%End If%>><%=String(15,"-")%> לא אחראי קבוצות <%=String(15,"-")%></option>
	<%
        If trim(USER_ID) <> "" Then
            set groupList=con.GetRecordSet("SELECT group_ID,group_Name FROM Responsibles_to_Groups_View WHERE ORGANIZATION_ID = " & trim(OrgID) & " AND Responsible_ID = "&USER_ID&" Order By group_Name")
            do while not groupList.EOF
                prgroupID=groupList("group_ID")
                prgroupName=groupList("group_Name")%>
                <option value="<%=trim(prgroupID)%>" <%if InStr(ResInGroups,prgroupName) > 0 then%> selected <%end if%>><%=prgroupName%></option>
            <%groupList.MoveNext
            loop
            groupList.close
            set groupList=Nothing
            End If
            If trim(USER_ID) <> "" Then
            set groupList=con.GetRecordSet("SELECT group_ID,group_Name FROM Users_Groups WHERE ORGANIZATION_ID = " & trim(OrgID) & " AND group_ID NOT IN ( Select group_ID From Responsibles_to_Groups WHERE Responsible_ID = "&USER_ID&") Order By group_Name")
            Else
            set groupList=con.GetRecordSet("SELECT group_ID,group_Name FROM Users_Groups WHERE ORGANIZATION_ID = " & trim(OrgID) & " Order By group_Name")
            End If
            do while not groupList.EOF
                prgroupID=groupList("group_ID")
                prgroupName=groupList("group_Name")%>
                <option value="<%=trim(prgroupID)%>" ><%=prgroupName%></option>
            <%groupList.MoveNext
            loop
            groupList.close
            set groupList=Nothing%>
	</select>
	</td>
	<td align="right" bgcolor="#DDDDDD" class="10normalB" nowrap width="220" valign=top>&nbsp;<span id="word17" name=word17>אחראי קבוצות</span>&nbsp;&nbsp;</td>
</tr>
<tbody name="resp_edit" id="resp_edit">
<tr>
	<td align=right bgcolor="#E6E6E6"><input type=radio name="EDIT_APPEAL" value="0" <%If trim(EDIT_APPEAL) = "0" Then%> checked <%End If%> ID="Radio1"></td>
	<td align="right" bgcolor="#DDDDDD" class="10normalB" nowrap width="220">&nbsp;אחראי מורשה צפיה בלבד&nbsp;</th>
</tr>
<tr>
	<td align=right bgcolor="#E6E6E6"><input type=radio name="EDIT_APPEAL" value="1" <%If trim(EDIT_APPEAL) = "1" Then%> checked <%End If%> ID="Radio2"></td>
	<td align="right" bgcolor="#DDDDDD" class="10normalB" nowrap width="220">&nbsp;אחראי מורשה עדכון&nbsp;</td>
</tr>
</tbody>
<%End If%>
<tr><td colspan="4" bgcolor="#808080" height="1"></td></tr>
<tr><td class="11normalB" align=right bgcolor="#E6E6E6" colspan=4>הרשאות מודולים&nbsp;<SPAN style="FONT-WEIGHT: 500">אנא סמן את המודולים שיוצגו בפני העובד</SPAN></td></tr>
<%sql_obj="Select bar_id,bar_title,parent_id, parent_title From dbo.bar_organizations_table('"&OrgID&"') "&_
  " WHERE PARENT_ID IS NOT NULL AND IS_VISIBLE = '1' Order By Parent_Order, bar_Order"
set rs_obj=con.getRecordSet(sql_obj)
While not rs_obj.eof
	barID = trim(rs_obj(0))
	barTitle = trim(rs_obj(1))
	barParent = trim(rs_obj(2))
	parentTitle = trim(rs_obj(3))
	
	If trim(barID) = "43" Then
		barTitle = barTitle & " <font color=red>(הרשאות טפסים)</font>"
		bgcolor = "#FFBFBF"
	ElseIf trim(barID) = "10" Then
		barTitle = barTitle & " <font color=red></font>"
		bgcolor = "#FFBFBF"	
	Else
		bgcolor = "#E6E6E6"
	End If  

	If USER_ID<>nil and USER_ID<>"" then	   	
		sql_obj="Select is_visible From dbo.bar_users_table('" & OrgID & "','" & USER_ID & "') "&_
		" WHERE bar_id =" & barID
		set rs_visible=con.getRecordSet(sql_obj)
		If not rs_visible.eof Then		
			barVisible = rs_visible("is_visible")
		Else
			barVisible = 0
		End If
		set rs_visible = Nothing				
	Else
		barVisible = 1
	End If

    If isNumeric(barParent) Then
		
	If old_parent <> parentTitle Then	

	If USER_ID<>nil and USER_ID<>"" then
		sql_obj="Select is_visible From dbo.bar_users_table('" & OrgID & "','" & USER_ID & "') WHERE bar_id = " & barParent
		'Response.Write sql_obj
		'Response.End
		set rs_parent=con.getRecordSet(sql_obj)
		If not rs_parent.eof Then					
			parentVisible = trim(rs_parent(0))			
		End If
		set rs_parent = Nothing		
	Else
		parentVisible = 1	
	End If	
%>
<%If old_parent <> parentTitle Then%>
<tr><td colspan="4" bgcolor="#808080" height="1"></td></tr>
<%End If%>
<tr>
	<td align=right width=100% bgcolor="#C0C0C0"><input type=checkbox dir="ltr" name="is_visible<%=barParent%>" <%If trim(parentVisible) = "1" Then%> checked <%End If%> ID="is_visible<%=barParent%>" onclick="return check_all_bars(this,'<%=barParent%>')"></td>
	<th bgcolor="#C0C0C0" class="11normalB" align=right nowrap><%=parentTitle%>&nbsp;</th>
</tr>
<%
	End If	
	End If
%>
<tr bgcolor="<%=bgcolor%>">
	<td align=right width=100%><input type=checkbox dir="ltr" name="is_visible<%=barID%>" id="<%=barID%>!<%=barParent%>" <%If trim(barVisible) = "1" Then%> checked <%End If%>></td>
	<th class="10normalB" align=right dir=rtl>&nbsp;<%=barTitle%></th>
</tr>
<% 
   old_parent = parentTitle
   rs_obj.moveNext
   Wend
   set  rs_obj = Nothing
%>
<tr><td colspan="4" bgcolor="#808080" height="1"></td></tr>
<tr><td colspan="4" bgcolor="#ffffff" height="1"></td></tr>
<tr><td colspan="2" >&nbsp;</td></tr>
</table>
</form>
</div>
</body>
<%
set con = nothing
%>
</html>

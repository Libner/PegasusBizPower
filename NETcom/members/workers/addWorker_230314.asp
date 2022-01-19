<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
if OrgID<>nil and OrgID<>"" then
  sqlStr = "select ORGANIZATION_NAME from ORGANIZATIONS where ORGANIZATION_ID=" & OrgID  
  ''Response.Write sqlStr
  set rs_Org = con.GetRecordSet(sqlStr)  
  if not rs_Org.eof then
	ORGANIZATION_NAME = rs_Org("ORGANIZATION_NAME")						
  end if
  set rs_Org = nothing
end if

errLogin=false
errEmail=false
USER_ID = trim(Request("USER_ID"))

if request.form("FIRSTNAME")<>nil then 'after form filling
	 FIRSTNAME = sFix(request.form("FIRSTNAME"))
	 LASTNAME  = sFix(trim(Request.Form("LASTNAME")))
	 LOGINNAME = sFix(request.form("LOGINNAME"))
	 PASSWORD  = sFix(request.form("PASSWORD"))
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
     job_pay = trim(Request.Form("job_id"))
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
     	 
     Order_Price = nFix(Request.Form("Order_Price"))
     
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
				"COMPANIES,TASKS,CASH_FLOW,CHIEF,EDIT_APPEAL,Order_Price) " &_
				" values ('" & FIRSTNAME &"','"& LASTNAME &"','"& LOGINNAME &"','"& PASSWORD &"',1,"& OrgID &","&_
				job_id &",'"& phone &"','"& mobile &"','"& fax &"','"& email &"','" & Month_Min_Order & "','" &_
				WORK_PRICING& "','" &SURVEYS& "','" &EMAILS& "','" &COMPANIES& "','" & TASKS & "','" &_
				CASH_FLOW & "'," & CHIEF & ",'" & EDIT_APPEAL & "','" & Order_Price & "'); SELECT @@IDENTITY AS NewID"
				'Response.Write sqlStr
				'Response.End
				set rs_tmp = con.getRecordSet(sqlStr)
					User_ID = rs_tmp.Fields("NewID").value
				set rs_tmp = Nothing	
				
				'--insert into changes table
				sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
				" SELECT 'חשבון משתמש', IsNULL(U.FIRSTNAME, '') + ' ' + IsNULL(LASTNAME, ''), [USER_ID], 'הוספה', getDate(), " & UserID & _
				" FROM dbo.USERS U WHERE ([USER_ID] = " & User_ID & ")"
				con.executeQuery(sqlstr)								
										  
			 elseif errLogin=true then
				 LOGINNAME = "" %>
			     <script language="javascript" type="text/javascript">
			     <!--
				   alert('שם משתמש זה בשימוש, בחר שם משתמש אחר')
				   history.back();
			      //-->
			     </SCRIPT>			 			 
			 <%Response.End
			   elseif errEmail=true then
				  email=""%>
			      <script language="javascript" type="text/javascript">
			      <!--
				     alert('כתובת דואר אלקטרוני שציינת'+'\n'+'רשומה כבר במערכת') ;
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
		         
               Order_Price = nFix(Request.Form("Order_Price"))
				 
				 if errLogin=false AND errEmail=false  then
					sqlStr = "UPDATE USERS set FIRSTNAME='" & FIRSTNAME &"',LASTNAME='"& LASTNAME &_
					"', LOGINNAME='"& LOGINNAME &"', PASSWORD = '"& PASSWORD &"',job_id="& job_id &_
					", TELEPHONE='"& phone &"',MOBILE='"& mobile &"',FAX='"& fax &"',EMAIL='"& email &_
					"', WORK_PRICING='" & WORK_PRICING & "',SURVEYS='" & SURVEYS & "',EMAILS='" & EMAILS &_
					"', COMPANIES='" & COMPANIES & "', TASKS='" & TASKS & "', CASH_FLOW='" & CASH_FLOW &_
					"', Month_Min_Order='" & Month_Min_Order & "', CHIEF = " & CHIEF & ", EDIT_APPEAL = '" & EDIT_APPEAL &_
					"', Order_Price = '" & Order_Price & "' WHERE USER_ID=" & USER_ID
					'Response.Write sqlStr
					'Response.End
					con.GetRecordSet (sqlStr)	
					
				'--insert into changes table
				sqlstr = "INSERT INTO Changes (Change_Table, Object_Title, Table_ID, Change_Type, Change_Date, [User_ID]) " & _
				" SELECT 'חשבון משתמש', IsNULL(U.FIRSTNAME, '') + ' ' + IsNULL(LASTNAME, ''), [USER_ID], 'עדכון', getDate(), " & UserID & _
				" FROM dbo.USERS U WHERE ([USER_ID] = " & User_ID & ")"
				con.executeQuery(sqlstr)					
									     
				 elseif errLogin=true then
					LOGINNAME = ""%>
				    <script language="javascript" type="text/javascript">
				     <!--
					   alert('שם משתמש זה בשימוש, בחר שם משתמש אחר')
					   history.back();
				     //-->
				   </SCRIPT>			 			 
				 <%Response.End
				 elseif errEmail=true then
				   email=""%>
			       <script language="javascript" type="text/javascript">
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
			
	sqlstr="Select bar_id from bars Order by PARENT_ID, bar_Order"
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
	"' WHERE USER_ID="&USER_ID
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
	
    Response.Redirect "addWorker.asp?User_ID=" & USER_ID       
end if
%>
<%
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 39 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
	arrTitlesD = rstitle.getRows()		
	redim arrTitles(Ubound(arrTitlesD,2)+1)
	For i=0 To Ubound(arrTitlesD,2)		
		arrTitles(arrTitlesD(0,i)) = arrTitlesD(1,i)		
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
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<script language="javascript" type="text/javascript">
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
			<%
			If trim(lang_id) = "1" Then
				str_alert = "'*' אנא מלאו כל השדות הצוינות בסימן"
			Else
				str_alert = "Please insert all requested fields !!"
			End If   
			%>			
			window.alert("<%=str_alert%>");						
			return false;				
		}
	else if (!checkEmail(document.all("email").value) && document.all("email").value != "")
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
	else
		{
		document.frmMain.submit();
		return true;
		}		
}

	function GetNumbers ()
	{
		var ch=event.keyCode;
		event.returnValue =(ch >= 48 && ch <= 57) || ch == 46 ;
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
//-->
</script>  
<body >
<!--#include file="../../logo_top.asp"-->
<%numOftab = 4%>
<%numOfLink = 5%>
<%If trim(wizard_id) = "" Then%>
<!--#include file="../../top_in.asp"-->
<%End If%>
<table border="0"  bgcolor="#E6E6E6" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title">&nbsp;&nbsp;</td></tr>  		       	
   </table></td></tr>            
<%
If Len(USER_ID) > 0 Then
	sqlStr = "SELECT FIRSTNAME,LASTNAME, LOGINNAME,PASSWORD,OFFICE,TELEPHONE,MOBILE,FAX,EMAIL,Month_Min_Order,chief, "&_
	" job_id,WORK_PRICING,SURVEYS,EMAILS,COMPANIES,TASKS,CASH_FLOW,CHIEF,ISNULL(EDIT_APPEAL,0),"&_
	" isNULL(Voice_Operator,0),IsNUll(Order_Price,0)  FROM dbo.USERS WHERE USER_ID=" & USER_ID  
	''Response.Write sqlStr
	Set rs_USERS = con.GetRecordSet(sqlStr)
	If Not rs_USERS.eof Then
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
		Order_Price = rs_USERS(21) 			
	end if
	set rs_USERS = nothing	
	
	Groups=""		
	sqlstr="Select Group_Name From Users_to_Groups_view Where User_id = " & USER_ID & " Order By Group_Name"
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
	
	user_groups=""		
	sqlstr="Select Group_ID From Users_to_Groups_view Where User_id = " & USER_ID & " Order By Group_id"
	set rssub = con.getRecordSet(sqlstr)		   
	If not rssub.eof Then
		user_groups = rssub.getString(,,",",",")		
	Else	
		user_groups = ""
	End If		    
	set rssub=Nothing
	If Len(user_groups) > 0 Then
		user_groups = Left(user_groups,(Len(user_groups)-1))
	End If	
	
	ResInGroups=""		
	sqlstr="Select Group_Name From Responsibles_to_Groups_view Where Responsible_id = " & USER_ID & " Order By Group_Name"
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
	
	ResInGroupsID=""		
	sqlstr="Select Group_ID From Responsibles_to_Groups_view Where Responsible_id = " & USER_ID & " Order By Group_ID"
	set rssub = con.getRecordSet(sqlstr)		   
	If not rssub.eof Then
		ResInGroupsID = rssub.getString(,,",",",")	
	Else
		ResInGroupsID=""
	End If		    
	set rssub=Nothing
	If Len(ResInGroupsID) > 0 Then
		ResInGroupsID = Left(ResInGroupsID,(Len(ResInGroupsID)-1))
	End If
	End If
	
	If trim(Request.QueryString("JobId")) <> "" Then	
		job_id = trim(Request.QueryString("JobId"))
	End If %>
<tr><td width="100%"> 
<FORM name="frmMain" ACTION="addWorker.asp?USER_ID=<%=USER_ID%>" METHOD="post" onSubmit="return CheckFields()" ID="Form1">
<table align=center border="0" width="600" cellpadding="2" cellspacing="0">
<tr><td height=5 nowrap colspan=2></td></tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap>
	<select dir="<%=dir_obj_var%>" name="JOB_ID" class="norm" style="width:200px" ID="JOB_ID">
	<option value="0" ><%=arrTitles(15)%></option>
<%	sqlstr = "SELECT job_id,job_name FROM jobs WHERE (ORGANIZATION_ID = " & OrgID & ") ORDER BY job_id"
		set rs_jobs = con.getRecordSet(sqlstr)
		while not rs_jobs.eof	%>
	<option value="<%=trim(rs_jobs(0))%>" <%If trim(job_id) = trim(rs_jobs(0)) Then%> selected <%End If%>><%=rs_jobs(1)%></option>
<%	rs_jobs.moveNext
		wend
		set rs_jobs = nothing %>
	</select>
	</td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--עיסוק--><%=arrTitles(7)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="<%=dir_obj_var%>" name="FIRSTNAME" value="<%=vFix(FIRSTNAME)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--שם פרטי--><%=arrTitles(3)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="<%=dir_obj_var%>" name="LASTNAME" value="<%=vFix(LASTNAME)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--שם משפחה--><%=arrTitles(4)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="LOGINNAME" value="<%=vFix(LOGINNAME)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--שם משתמש--><%=arrTitles(5)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="PASSWORD" value="<%=vFix(PASSWORD_U)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--סיסמה--><%=arrTitles(6)%>&nbsp;<span style="color:#FF0000">*</span></td>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="phone" value="<%=vfix(phone)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--טלפון--><%=arrTitles(8)%>&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="mobile" value="<%=vfix(mobile)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" align="<%=align_var%>" class="title_show_form" nowrap width="150"><!--נייד--><%=arrTitles(9)%>&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="fax" value="<%=vfix(fax)%>"  style="width:200" maxlength=20 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--פקס--><%=arrTitles(10)%>&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="email" value="<%=vfix(email)%>"  style="width:200" maxlength=100 style="font-family:arial"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;<!--דואר אלקטרוני--><%=arrTitles(11)%>&nbsp;<span style="color:#FF0000">*</span></th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type=text dir="ltr" name="Month_Min_Order" value="<%=vfix(Month_Min_Order)%>" size="3" maxlength="3" style="font-family:arial" onkeypress="GetNumbers();"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;יעד הזמנות מינימלי&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="text" size="3" maxlength="3" dir="ltr" style="font-family:arial" name="Order_Price" value="<%=nFix(Order_Price)%>" ID="Order_Price"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;שווי כספי של הזמנה&nbsp;&nbsp;</th>
</tr>
<tr>
	<td align="<%=align_var%>" width=450 nowrap><input type="checkbox" dir="ltr" name="chief" <%If trim(chief) = "1" Then%> checked <%End If%> ID="chief"></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="150">&nbsp;מורשה מחיקה&nbsp;&nbsp;</th>
</tr>
<% If trim(is_groups) = "1" Then %>
<tr><td height=5 nowrap colspan=3></td></tr>
<tr><td height=1 nowrap colspan=3 bgcolor=#808080></td></tr>
<tr><td height=5 nowrap colspan=3></td></tr>
<tr><td colspan=3 align=center><table cellpadding=2 cellspacing=0 width=600 border=0>
<tr><td class="title" align="<%=align_var%>" colspan=3><!--הרשאות--><%=arrTitles(21)%></td></tr>
<tr>
    <td width=300 nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top>
    <!--העובד יוכל להזין את כל הטפסים המשויכים לקבוצות אלו.--><%=arrTitles(23)%></td>
	<td align="<%=align_var%>" valign=top>	
    <INPUT type=image src="../../images/icon_find.gif" style="vertical-align:top" name=groups_list id="groups_list" onclick='window.open("groups_list.asp?USER_ID=<%=USER_ID%>&user_groups=" + GROUP_ID.value,"TypesList","left=300,top=100,width=250,height=350,scrollbars=1"); return false;'>&nbsp;
    <span class="Form_R" dir="<%=dir_obj_var%>" style="width:190;line-height:16px" name="groups_names" id="groups_names"><%=Groups%></span>
    <input type=hidden name=GROUP_ID id=GROUP_ID value="<%=user_groups%>">
	</td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="100" valign=top><!--קבוצה--><%=arrTitles(14)%></td>
</tr> 
<tr>
    <td width=300 nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top>
   <!-- העובד יוכל לצפות או לעדכן את כל הטפסים שהוזנו ע"י העובדים בקבוצות אלו.--><%=arrTitles(24)%>
    </td>
	<td align="<%=align_var%>" valign=top>
    <INPUT type=image src="../../images/icon_find.gif" style="vertical-align:top" onclick='window.open("res_groups_list.asp?USER_ID=<%=USER_ID%>&res_groups_id=" + R_GROUP_ID.value,"TypesList","left=300,top=100,width=250,height=350,scrollbars=1"); return false;'>&nbsp;
    <span class="Form_R" dir="<%=dir_obj_var%>" style="width:190;line-height:16px" name="ResInGroups" id="ResInGroups"><%=ResInGroups%></span>
    <input type=hidden name=R_GROUP_ID id="R_GROUP_ID" value="<%=ResInGroupsID%>">
	</td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="100" valign=top><!--אחראי קבוצות--><%=arrTitles(17)%></td>
</tr>
<tbody name="edit_appeal_tbody" id="edit_appeal_tbody" <%If trim(ResInGroupsID) = "" Then%> style="display:none" <%End If%>>
<tr>
	<th colspan=3 align="<%=align_var%>" dir="<%=dir_var%>" class="title_show_form">&nbsp;<!--אחראי מורשה צפיה בלבד--><%=arrTitles(19)%><input type=radio name="EDIT_APPEAL" value="0" <%If trim(EDIT_APPEAL) = "0" Then%> checked <%End If%> ID="Radio1" style="position:relative;top:2px"></th>
</tr>
<tr>
	<th colspan=3 align="<%=align_var%>" dir="<%=dir_var%>" class="title_show_form">&nbsp;<!--אחראי מורשה עדכון--><%=arrTitles(20)%><input type=radio name="EDIT_APPEAL" value="1" <%If trim(EDIT_APPEAL) = "1" Then%> checked <%End If%> ID="Radio2" style="position:relative;top:2px"></th>
</tr>
</tbody>
<tr><td height=5 nowrap colspan=3></td></tr>
<tr><td height=1 nowrap colspan=3 bgcolor=#808080></td></tr>
</table></td></tr>
<%End If%>
<tr><td height=5 nowrap colspan=3></td></tr>
<tr><td class="title" align="<%=align_var%>" colspan=3><!--הרשאות--><%=arrTitles(13)%>&nbsp;&nbsp;<span style="font-weight:500"><%=arrTitles(22)%></span></td></tr>
<tr><td colspan=3><table cellpadding=0 cellspacing=0 width="100%">
<% 
sql_obj="Select bar_id,bar_title,parent_id,parent_title from dbo.bar_organizations_table('"&OrgID&"') "&_
" WHERE PARENT_ID IS NOT NULL AND IS_VISIBLE = '1' Order by Parent_Order, bar_Order"
set rs_obj=con.getRecordSet(sql_obj)
While not rs_obj.eof
	barID = trim(rs_obj(0))
	barTitle = trim(rs_obj(1))
	barParent = trim(rs_obj(2))
	parentTitle = trim(rs_obj(3))
	If trim(barID) = "43" Then
		barTitle = barTitle & " <font color=red>(הרשאות טפסים)</font>"
	ElseIf trim(barID) = "10" Then
		barTitle = barTitle & " <font color=red></font>"	
	End If
	
	JobId = trim(Request.QueryString("JobId")) 

	If Len(USER_ID) > 0 Or (JobId <> "") then	   	
	
		If Len(USER_ID) > 0 Then
			sql_obj="Select is_visible From dbo.bar_users_table('" & OrgID & "','" & USER_ID & "') WHERE bar_id = " & barID
		Else
			sql_obj="SELECT is_visible FROM dbo.bar_jobs WHERE (organization_id = '" & OrgID & "') AND (job_id = '" & JobId & "') " & _
			" AND (bar_id = " & barID & ")"
		End If
		'Response.Write sql_obj
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

	If Len(USER_ID) > 0 Or (JobId <> "")  then
		If Len(USER_ID) > 0 Then
			sql_obj="SELECT is_visible FROM dbo.bar_users_table('" & OrgID & "','" & USER_ID & "') WHERE bar_id = " & barParent
		Else
		sql_obj="SELECT is_visible FROM dbo.bar_jobs WHERE (organization_id = '" & OrgID & "') AND (job_id = '" & job_id & "') " & _
		" AND (bar_id = " & barParent & ")"			
		End If	
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
<tr>
	<th class="form_title" align="<%=align_var%>" colspan=2><input type=checkbox dir="ltr" name="is_visible<%=barParent%>" <%If trim(parentVisible) = "1" Then%> checked <%End If%> ID="is_visible<%=barParent%>" onclick="return check_all_bars(this,'<%=barParent%>')"></th>
	<th align="<%=align_var%>" class="form_title" nowrap width="250" dir="<%=dir_obj_var%>"><%=parentTitle%>&nbsp;</th>
</tr>
<%
	End If	
	End If
%>
<tr>
	<td align="<%=align_var%>" colspan=2><input type=checkbox dir="ltr" name="is_visible<%=barID%>" id="<%=barID%>!<%=barParent%>"  <%If trim(barVisible) = "1" Then%> checked <%End If%>></td>
	<th align="<%=align_var%>" class="title_show_form" nowrap width="250" dir="<%=dir_obj_var%>"><%=barTitle%>&nbsp;</th>
</tr>
<% 
   old_parent = parentTitle
   rs_obj.moveNext
   Wend
   set  rs_obj = Nothing
%>
</table></td></tr>
<tr><td height=5 nowrap colspan=3></td></tr>
<tr><td height=1 nowrap colspan=3 bgcolor=#808080></td></tr>
<tr><td height=10 nowrap colspan=3></td></tr>
<tr>
<td colspan=3>
<table cellpadding=0 cellspacing=0 width=500 align=center>
<tr>
<td width=50% align="center"><input type=button class="but_menu" style="width:90px" onclick="document.location.href='default.asp';" value="<%=arrButtons(2)%>" id=Button2 name=Button2></td>
<td width=50 nowrap></td>
<td width=50% align=center><input type=button class="but_menu" style="width:90px" onclick="javascript:CheckFields(); return false;" value="<%=arrButtons(1)%>" id=Button1 name=Button1></td></tr>
</table></td></tr>
<tr><td colspan=3 height=10 nowrap></td></tr>
</table></form>
</td></tr></table>
</body>
</html>
<%set con = nothing%>
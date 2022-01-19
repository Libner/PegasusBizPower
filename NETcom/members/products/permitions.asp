<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
  prodID = trim(Request.QueryString("prodID"))
  If trim(prodID) <> "" Then
	 sqlstr = "Select product_name From products WHERE product_id = " & prodID & " And Organization_ID = " & OrgID
	 set rs_prod = con.getRecordSet(sqlstr)
	 If not rs_prod.eof Then
		 prodName = rs_prod(0)
		 found_prod = true
	 Else
		 found_prod = false	
	 End if
  Else	 
	 found_prod = false
  End if
    
  If trim(Request.QueryString("add")) <> nil And trim(prodID) <> "" Then 
      
	users_list = Request.Form("user")
	If Len(users_list) > 0 Then
		users_arr = Split(users_list,",")
		dim Users()
		If IsArray(users_arr) Then
		    redim Users(Ubound(users_arr),2)
			For i=0 To uBound(users_arr)
				user_arr = Split(users_arr(i), "!")
				Users(i,0) = user_arr(0) ' user_id
				Users(i,1) = user_arr(1) ' group_id
			Next
		End If
	End If		
	
	sqlstr = "Delete FROM Users_To_Products WHERE Product_ID = " & prodID & " And ORGANIZATION_ID = " & orgID
	con.executeQuery(sqlstr)
	
    If IsArray(Users) And IsNull(Users)=false Then
		For i=0 To Ubound(Users)
			sqlstr = "Insert Into Users_To_Products values (" & Users(i,0) & "," & Users(i,1) & "," & OrgID & "," & prodID & ")"
			con.executeQuery(sqlstr)
		Next
    End If
        
    Response.Redirect "questions.asp"
  End if
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 66 Order By word_id"				
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
<script language=javascript>
<!--
	function add_remove(chkBoxName,chkBox,id)
	//the function add's and remove's the contacter id number from string
	{
	 var str,string1,beginslice,endSlice,x
     document.all("all_"+chkBoxName+"").checked = false;
	 if(chkBox.checked) 
	 {    
	    if(document.all(""+chkBoxName+"").value != "")
			document.all(""+chkBoxName+"").value=document.all(""+chkBoxName+"").value+','+id; //add dealer to the exclude list
		else
			document.all(""+chkBoxName+"").value=id; //add dealer to the exclude list
	 }
	 else
	 {
	    str = new String(document.all(""+chkBoxName+"").value); // uncheck then check -> remove contact from the exclude list
	    string1 = new String(","+id); 
	    beginSlice = str.indexOf(string1); //index of start of the contact id in the exclude list	   
	    endSlice = beginSlice + string1.length; //index of end of the contact id in the exclude list
	    str = str.slice(0,beginSlice) + str.slice(endSlice+1); // add two peaces together	    
	    document.all(""+chkBoxName+"").value = str; //put into the hidden variable
	 } 
	 //window.alert(document.all(""+chkBoxName+"").value);
	}	
	
function CheckFields()
{  
   return window.confirm("?האם ברצונך להוסיף את <%=trim(Request.Cookies("bizpegasus")("ContactsMulti"))%> לפי קרטריונים שנבחרו לקבוצת הדיוור");
}	

function check_all_users(objChk,groupID)
{
	input_arr = document.getElementsByName("user");	
	for(i=0;i<input_arr.length;i++)
	{
		currGroupId = "";
		objValue = new String(input_arr(i).value);
		value_arr = objValue.split("!");
		currGroupId =  value_arr[1];
		if(currGroupId == groupID)
		{
			//input_arr(i).disabled = objChk.checked;
			input_arr(i).checked = objChk.checked;
		}	
	}
	return true;
}
//-->
</script>
</head>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOfTab = 1%>
<%numOfLink = 1%>
<!--#include file="../../top_in.asp"-->
<%If found_prod Then%>
<table width="100%" cellspacing="0" cellpadding="0" align=center border="0" bgcolor="#ffffff">
<tr><td class="page_title" width=100% dir="<%=dir_obj_var%>">&nbsp;<!--הרשאות--><%=arrTitles(1)%>&nbsp;-&nbsp;<font color="#6F6DA6"><%=prodName%></font></td></tr>
<tr><td height=20 nowrap></td></tr>
<%
	sqlStr = "Select group_Id, group_Name from Users_Groups WHERE ORGANIZATION_ID= "& OrgID &_
	" Order By group_Name"
	'Response.Write sqlStr
	set rs_groups = con.GetRecordSet(sqlStr)
	If not rs_groups.eof Then
%>				
<tr>
<td align=center width=100% valign=top>
<table cellpadding=0 cellspacing=0 width=500 dir="center">
<tr>
	<td align="<%=align_var%>" class=card style="padding:5px">
	<table border=0 width=500 cellpadding=0 cellspacing=0 align=center dir="<%=dir_var%>">		
	<tr><td align="<%=align_var%>" dir="<%=dir_obj_var%>" class=card>אנא סמן את העובדים אשר מורשים להזין טופס זה במערכת.</td></tr>
	<tr><td align="<%=align_var%>" dir="<%=dir_obj_var%>" class=card>כל עובד יראה רק את הטפסים שהוא הזין.</td></tr>
	<tr><td align="<%=align_var%>" dir="<%=dir_obj_var%>" class=card>עובד אחראי על קבוצה יראה את כל הטפסים שהזינו עובדים בקבוצה זו.</td></tr>
	<tr><td align="<%=align_var%>" dir="<%=dir_obj_var%>" class=card>במידה שהטופס הינו טופס אינטרנטי (ההזנה לא מתבצעת ע"י עובד אלא ע"י גולש אינטרנט) כל העובדים המסומנים יוכלו לצפות בנתוני הטופס.</td></tr>
	</table>
	</td>
</tr>
<tr><td height=10 nowrap></td></tr>
<form name="form1" id="form1" target="_self" ACTION="permitions.asp?prodID=<%=prodID%>&add=1" METHOD="post">
<tr>
	<td align="<%=align_var%>" width=100%>
	<input type=hidden name="users_id" id="users_id" value="">
	<table border=0 width=100% cellpadding=0 cellspacing=0 align=center dir="<%=dir_var%>">		
		<tr height=25>																							
			<td align="<%=align_var%>" width=100%>
			<table cellpadding=2 cellspacing=1 width=100% border=0  dir="<%=dir_obj_var%>">			
			<%
				while not rs_groups.eof
					group_Id = rs_groups("group_Id")
					group_Name = rs_groups("group_Name")
					is_checked_gr = false
					If Len(group_Id) > 0 Then
						sqlstr = "Select Top 1 User_ID From Users_To_Products WHERE Group_ID = " &_
						group_Id & " And Product_ID = " & prodID & " And Organization_ID = " & OrgID
						set rs_check = con.getRecordSet(sqlstr)
						If not rs_check.eof Then
							is_checked_gr = true
						Else
							is_checked_gr = false
						End if
						set rs_check = Nothing
					
    					sqlstr = "Select FIRSTNAME + Char(32) + LASTNAME From Responsibles_To_Groups_View WHERE Group_ID = " &_
    					group_Id & " And Organization_ID = " & trim(OrgID)
    					set rs_resp = con.getRecordSet(sqlstr)
    					If not rs_resp.eof Then
    						reponsibles_list = rs_resp.getString(,,",",",")
    						If Len(reponsibles_list) > 1 Then
    							reponsibles_list = Left(reponsibles_list, Len(reponsibles_list)-1)
    						End If
    					End If
    					set rs_resp = nothing
    			  End If					
			
				sqlStr = "Select User_ID, FIRSTNAME + Char(32) + LASTNAME from USERS WHERE ORGANIZATION_ID= "& OrgID &_
				" AND  User_ID IN (Select User_ID FROM Users_To_Groups WHERE ORGANIZATION_ID= "& OrgID &_
				" AND Group_ID = "& group_Id &") Order by 2"
				'Response.Write sqlStr
				set rs_users = con.GetRecordSet(sqlStr)
				If not 	rs_users.Eof Then
			%>
			<tr><td colspan=2 height=1 bgcolor=#808080></td></tr>
			<tr>
			<th width=22 align="<%=align_var%>" nowrap class="form_title"><input type=checkbox name="all_users_id" value="all_users_id" ID="all_users_id" onclick="return check_all_users(this,'<%=group_Id%>')" <%If is_checked_gr Then%> checked <%End If%>></th>
			<th align="<%=align_var%>" width=100% class="form_title">&nbsp;<%=group_Name%>&nbsp;(<%=reponsibles_list%>)</th>
			</tr>	
			<%	
				do while not rs_users.eof
					user_ID = trim(rs_users(0))
					user_Name = trim(rs_users(1))
					is_checked = false
					If Len(user_ID) > 0 Then
						sqlstr = "Select Top 1 User_ID From Users_To_Products WHERE User_ID = " & user_ID &_
						" And Group_ID = " & group_Id & " And Product_ID = " & prodID & " And Organization_ID = " & OrgID
						set rs_check = con.getRecordSet(sqlstr)
						If not rs_check.eof Then
							is_checked = true
						Else
							is_checked = false
						End if
						set rs_check = Nothing
					End if
			%>
			<tr>
			<td width=22 align="<%=align_var%>" nowrap class="card"><input type=checkbox ID="<%=user_ID%>" value="<%=user_ID&"!"&group_Id%>" name="user" onclick="return add_remove('users_id',this,'<%=user_ID%>')" <%If is_checked Then%> checked <%End If%>></td>
			<td align="<%=align_var%>" width=100% class="card">&nbsp;<%=user_Name%>&nbsp;</td>			
			</tr>
			<%
				
				rs_users.movenext
				loop
				set rs_users = nothing
				End If			
				rs_groups.movenext				
				Wend				
			%>
			<tr><td colspan=2 height=1 bgcolor=#808080></td></tr>
			</table>		
			</td>
		</tr>		
	</table>					
	</td>
</tr>
<tr><td colspan=2 height=20 nowrap></td></tr>
<tr><td colspan=2 align="center">
	<table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
	<tr>
		<td width=50% align="center" nowrap>
		<INPUT class="but_menu" type="button" style="width:90px" value="<%=arrButtons(2)%>" id=button2 name=button2 onclick="document.location.href='questions.asp'"></td>
		<td width=50 nowrap></td>
		<td width=50% align="center" nowrap>
		<input style="width:90px" class="but_menu" type="submit" value="<%=arrButtons(1)%>" id=button1 name=button1>
		</td></tr>
	</table></td>
</tr>		
</form>	
</table></td></tr>
<%
Else
%>
<script language=javascript>
<!--
	window.history.back();
//-->
</script>
<%
End If
set rs_groups = nothing
%>
</table>
<%End If%>
</body>		
</html>		
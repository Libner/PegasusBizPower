<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%Response.Buffer = False
  groupId = trim(Request.QueryString("groupId"))
  UserID = trim(trim(Request.Cookies("bizpegasus")("UserID")))
  OrgID = trim(trim(Request.Cookies("bizpegasus")("OrgID")))
  wizard_id = trim(Request.Cookies("bizpegasus")("wizardId"))
  wizard_page_id = 2
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
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

function check_all_messangers(objChk)
{
	input_arr = document.getElementsByName("messanger");	
	for(i=0;i<input_arr.length;i++)
	{
		input_arr(i).disabled = objChk.checked;
		input_arr(i).checked = objChk.checked;
	}
	return true;
}

function check_all_types(objChk)
{
	input_arr = document.getElementsByName("type");	
	for(i=0;i<input_arr.length;i++)
	{
		input_arr(i).disabled = objChk.checked;
		input_arr(i).checked = objChk.checked;
	}
	return true;
}
//-->
</script>

</head>
<body style="margin:0px;background-color:#E5E5E5" onload="window.focus()">
<% If trim(Request.QueryString("add")) <> nil Then%>
<div id="div_save" bgcolor="#e8e8e8" name="div_save" style="position:absolute; left:0px; top:0px; width:100%; height:100%;" >  												
  <table bgcolor="#e8e8e8" height="100%" width="100%" cellspacing="2" cellpadding="2" ID="Table3">  
     <tr><td bgcolor="#e8e8e8" align="center" >
         <table bgcolor="#ebebeb" border="0" height="100" width="400" cellspacing="1" cellpadding="1">
            <tr>  
              <td align="center" bgcolor="#d0d0d0">
              <font style="font-size:14px;color:#FF0000;"><b>מתבצעת טעינת הנתונים</b></font>
              <br>
              <font style="font-size:14px;color:#000000;">...אנא המתן</font>
              </td>
            </tr>
         </table>
         </td>
     </tr>
  </table>
</div>
<%	
    If trim(groupId) <> ""  Then
		If Request.Form("types_id") <> nil Then
			ind = InStr(Request.Form("types_id"),",")
			If ind = 1 Then
				types_id = Right(Request.Form("types_id"),Len(Request.Form("types_id"))-1)
			Else
				types_id = trim(Request.Form("types_id"))
			End if			
		Else
			types_id = ""
		End If
	
		If Request.Form("messangers_id") <> nil Then
			ind = InStr(Request.Form("messangers_id"),",")
			If ind = 1 Then
				messangers_id = Right(Request.Form("messangers_id"),Len(Request.Form("messangers_id"))-1)
			Else
				messangers_id = trim(Request.Form("messangers_id"))
			End if	
		Else
			messangers_id = ""
		End If
	
		sqlcheck="Select Count(PEOPLE_ID) From PEOPLES Inner Join CONTACTS On PEOPLES.Contact_Id = CONTACTS.Contact_Id" & _
		" WHERE CONTACTS.COMPANY_ID IN (Select COMPANY_ID FROM companies WHERE CONTACTS.ORGANIZATION_ID = " & OrgID & ")"
		If Len(messangers_id) > 0 Then
		sqlcheck = sqlcheck & " AND messanger_id IN (" & messangers_id & ")"
		End If
		If Len(types_id) > 0 Then
		sqlcheck = sqlcheck & " AND CONTACTS.CONTACT_ID IN (Select CONTACT_ID FROM contact_to_types WHERE type_id IN (" & types_id & "))"
		End If
		sqlcheck = sqlcheck & " AND (PEOPLE_EMAIL LIKE email) AND (GROUPID = "&groupId&") AND CONTACTS.ORGANIZATION_ID = "&OrgID 
		'Response.Write sqlcheck
		'Response.End
		set rs_check = con.getRecordSet(sqlcheck)
		If not rs_check.eof Then	
			double_mail = trim(rs_check(0))
		End If
		set rs_check = Nothing		

	   sqlstr = "EXECUTE get_contacts '" & types_id & "','" & OrgID & "','" & groupId & "'"
	   'Response.Write sqlstr
	   'Response.End
	   con.executeQuery(sqlstr)
	   If IsNumeric(wizard_id) Then
             strUrl = "../wizard/wizard_" & wizard_id & "_" & wizard_page_id + 1 & ".asp?groupId=" & groupId
       Else
             strUrl = "group.asp?groupId=" & groupId & "&double_mail=" & double_mail
       End If
   %>    
   <script language=javascript>
	<!--
		 window.opener.document.location.href = "<%=strUrl%>";
		 window.close();
	//-->
	</script>
<%
  End if
 Else %>
<table cellpadding=0 cellspacing=0 width=100%>
<tr><td width="100%" class="page_title" dir=rtl>&nbsp;הוספת נמענים ממועדון <%=trim(Request.Cookies("bizpegasus")("CompaniesMulti"))%>&nbsp;<font color="#6E6DA6"><%=groupName%></font></td></tr>         
<tr><td width="100%" colspan=2 height=5></td></tr>
<%If trim(wizard_id) <> "" Then%>
<tr>
<td width=100% align="right" bgcolor="#FFD011" style="padding:5px">
<table border=0 cellpadding=0 cellspacing=0 width=100% ID="Table4">
<tr><td class=explain>
סמן את הסיווגים שברצונך להוסיף לקבוצת הדיוור, ואת בעלי התפקידים אליהם יישלח דף המבצע. 
</td></tr>
<tr><td class=explain>
לדוגמה: סימון הסיווגים "אפנה" ו"פרסום", ותפקיד "שיווק" תיצור קבוצת דיוור של כל אנשי השיווק בחברות המופיעות בסיווגי אפנה ופרסום במועדון ה<%=trim(Request.Cookies("bizpegasus")("CompaniesMulti"))%> שלך.
</td></tr>
<tr><td class=explain>
לאחר שבחרת סיווגים ותפקידים, לחץ על [טען רשימה].
</td></tr>
</table>
</td></tr>
<%End If%>
<tr><td width="100%" colspan=2>
<table border="0" width="95%" align=center cellspacing="1" cellpadding="2" ID="Table2">
<tr><td width=100%>
<table cellpadding=3 cellspacing=0 width=100%>
<tr>
<td width=100% align=right>בחר סיווגים לרשימת הדיוור</td>
<td width=20 nowrap align=center><b>1</b></td></tr>
</table>
</td></tr>
<form name="form1" id="form1" target="_self" ACTION="getContacts.asp?groupId=<%=groupId%>&add=1" METHOD="post">
<tr>
	<td align=right>
	<input type=hidden name="types_id" id="types_id" value="">
	<input type=hidden name="messangers_id" id="messangers_id" value="">
	<table border=0 width=100% cellpadding=3 cellspacing=0 ID="Table1">		
		<tr height=25>																							
			<td align="right" width=100%>
			<table cellpadding=3 cellspacing=0 width=100% dir=rtl>
			<tr>
			<td width=30 align=right><input type=checkbox name="all_types_id" id="all_types_id" value="" onclick="return check_all_types(this)"></td>
			<td align=right width=100% colspan=3><b>כולם</b></td>
			</tr>
			<%
				sqlStr = "Select DISTINCT type_Id, type_Name from contact_type WHERE type_Id IN ( "&_
				" Select distinct type_id from contact_to_types where ORGANIZATION_ID= "& OrgID &_
				" ) order by type_Name"
				'Response.Write sqlStr
				set rs_types = con.GetRecordSet(sqlStr)
				while not rs_types.eof
					type_Id = rs_types("type_Id")
					type_Name = rs_types("type_Name")	
			%>
			<tr>
			<td width=30 align=right nowrap><input type=checkbox name="type" value="<%=type_Id%>" ID="<%=type_Id%>" onclick="return add_remove('types_id',this,'<%=type_Id%>')"></td>
			<td align=right width=100 nowrap><%=type_Name%></td>
			<%
				rs_types.movenext
				If not rs_types.eof Then
					type_Id = rs_types("type_Id")
					type_Name = rs_types("type_Name")	
					
			%>
			<td width=30 align=right nowrap><input type=checkbox name="type" value="<%=type_Id%>" ID="<%=type_Id%>" onclick="return add_remove('types_id',this,'<%=type_Id%>')"></td>
			<td align=right width=100 nowrap><%=type_Name%></td>			
			<%	End If	%>
			</tr>
			<%	
				If not rs_types.eof Then
					rs_types.movenext
				End If
				Wend
				set rs_types = nothing
			%>
			</table>		
			</td>
			<td width=20 nowrap align=center>&nbsp;</td>														
		</tr>		
	</table>					
	</td>
</tr>
<tr><td width=100%>
<table cellpadding=0 cellspacing=0 width=100% ID="Table5">
<tr>
<td width=100% align=right><INPUT type="submit" class="but_menu" value="טען רשימה" id="upload_excel" name="upload_excel" style="width:100" onClick="return CheckFields();"></td>
<td width=150 nowrap align=right>לחץ על לחצן טען רשימה</td>
<td width=20 nowrap align=center><b>2</b></td></tr>
</table>
</td></tr>
</form>	
<tr><td height=10 nowrap></td></tr>	
</table></td></tr></table>
<%End If%>
</body>		
</html>		
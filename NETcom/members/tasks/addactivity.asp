<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%
  companyId=trim(Request("companyId"))
  contactID=trim(Request("contactID"))
  projectID=trim(Request("project_id")) 
  activityId=trim(Request("activityId"))  
  UserID=trim(Request.Cookies("bizpegasus")("UserID"))
  ReciverId=trim(Request("ReciverId"))
  
  If trim(contactID) <> "" And trim(companyId) = "" Then
	sqlstr = "Select company_id from contacts WHERE contact_id = " & contactID
	set rs_compID = con.getRecordSet(sqlstr)
	If not rs_compID.eof Then
		companyId = trim(rs_compID(0))
	End If
	set rs_compID = Nothing
  End If
  
  If companyId <> "" Then
		sqlDealer="SELECT COMPANY_NAME,private from COMPANIES where COMPANY_ID=" & companyId 
		set rs_comp=con.GetRecordSet(sqlDealer)
		if not rs_comp.eof then
			COMPANY_NAME=rs_comp("COMPANY_NAME")
			private_flag=rs_comp("private")		
		end if
		If trim(contactID) = "" And trim(companyId) <> "" And trim(private_flag) = "1" Then
			sqlstr = "Select contact_id from contacts WHERE company_Id = " & companyId
			set rs_compID = con.getRecordSet(sqlstr)
			If not rs_compID.eof Then
				contactID = trim(rs_compID(0))
			End If
			set rs_compID = Nothing
		End If
		where_company = " AND company_Id = " & companyId
  Else
	   private_flag	= trim(Request("private"))
	   where_company = ""
  End If
 
  OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
%>
<html>
<head>
<!--#include file="../../../include/title_meta_inc.asp"-->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
function add_remove(chkBoxName,chkBox,activity_type_id)
	//the function add's and remove's the contacter id number from string
	{
	 var str,string1,beginslice,endSlice,x

	 if(chkBox.checked) 
	 {    
	    if(document.all(""+chkBoxName+"").value != "")
			document.all(""+chkBoxName+"").value=document.all(""+chkBoxName+"").value+','+activity_type_id; //add dealer to the exclude list
		else
			document.all(""+chkBoxName+"").value=activity_type_id; //add dealer to the exclude list
	 }
	 else
	 {
	    str = new String(document.all(""+chkBoxName+"").value); // uncheck then check -> remove contact from the exclude list
	    string1 = new String(","+activity_type_id); 
	    beginSlice = str.indexOf(string1); //index of start of the contact id in the exclude list	   
	    endSlice = beginSlice + string1.length; //index of end of the contact id in the exclude list
	    str = str.slice(0,beginSlice) + str.slice(endSlice+1); // add two peaces together	    
	    document.all(""+chkBoxName+"").value = str; //put into the hidden variable
	 } 
	 //window.alert(document.all(""+chkBoxName+"").value);
	}	
	
function CheckFields(act)
{
   if(window.document.all("task_types").value!='' && window.document.all("ReciverId").value=='')
   {
      window.alert("! �� ����� �� ���� ������");     
      window.document.all("ReciverId").focus();
      return false; 
   }
   /*if(window.document.all("companyId").value=='')
   {
      window.alert("! �� ����� ����");  
      window.document.all("companyId").focus();   
      return false; 
   }
   if(window.document.all("contactID").value=='')
   {
      window.alert("! �� ����� ��� ���"); 
      //window.document.all("contactID").focus();    
      return false; 
   }   
   */
   if(window.document.all("activity_types").value=='')
   {
      window.alert("! �� ����� ��� ������");     
      return false; 
   }
   else if (document.all("activity_content").value=='')
   {
      window.alert("! �� ���� ���� ������");
      document.all("activity_content").focus();
      return false;
   } 
  
   if(window.document.all("task_types").value!='' && document.all("task_content").value=='' )
   {
		window.alert("! �� ���� ���� ������");
		return false;
   }
   
   if(window.document.all("activity_types").value=='' && document.all("activity_content").value!='' )
   {
		window.alert("! �� ����� ��� �����");
		return false;
   }   
    
   if (act=='submit')
   { 
       document.formactivity.submit();             
   }	 
   return true;

}

function DoCal(elTarget){
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
<!--End-->
</script>  
</head>
<%
  
  If Request.QueryString("add") <> nil Then	
    
    companyID = trim(Request.Form("companyID"))
    contactID = trim(Request.Form("contactID"))
    projectID = trim(Request.Form("project_id"))
    
    If Request.Form("activity_types") <> nil Then      
		ind = InStr(Request.Form("activity_types"),",")
		'Response.Write Request.Form("activity_types") & " " & ind
		'Response.End
		If ind = 1 Then
			activity_types = Right(trim(Request.Form("activity_types")),Len(Request.Form("activity_types"))-1)
		Else
			activity_types = trim(Request.Form("activity_types"))
		End If	
    Else
		activity_types = ""
    End If
    add_task = false
    
    taskID="NULL"    
    If Request.Form("task_types") <> nil Then
		ind = InStr(Request.Form("task_types"),",")
		If ind = 1 Then
			task_types = Right(Request.Form("task_types"),Len(Request.Form("task_types"))-1)
		Else
			task_types = trim(Request.Form("task_types"))
		End if	
	Else
		task_types = ""
	End If
	    
	If Request.Form("task_date") <> nil And IsDate(Request.Form("task_date")) Then    
		task_date = Month(trim(Request.Form("task_date"))) & "/" & Day(trim(Request.Form("task_date"))) & "/" & Year(trim(Request.Form("task_date")))
	End If
	
	If Request.Form("activity_date") <> nil And IsDate(Request.Form("activity_date")) Then    
	activity_date = Month(trim(Request.Form("activity_date"))) & "/" & Day(trim(Request.Form("activity_date"))) & "/" & Year(trim(Request.Form("activity_date")))
	End If
		
    if Request.Form("task_types") <> nil And Request.Form("task_content") <> nil Then
		add_task = true		
    End If 
    
    If trim(companyID) = "" Then
		companyID = "NULL"
	End If	
	If trim(contactID) = "" Then
		contactID = "NULL"
	End If	
	If trim(projectID) = "" Then
		projectID = "NULL"
	End If	
    
    If trim(Request.Form("taskId")="") And (add_task = true) And trim(Request.Form("activityId")="") Then ' ������ ����� ����	    
		con.executeQuery("SET DATEFORMAT mdy")
		sqlstr="Insert Into tasks (company_id,contact_id,project_id,User_ID,Organization_ID,task_date,task_content,task_types,task_status,reciver_id) "&_
		" values (" & companyID & "," & contactID & "," & projectID & "," & UserID & "," & OrgID & ",'" &_
		task_date & "','" & sFix(Request.Form("task_content")) & "','" & task_types & "','1'," & reciverID & ")"		
		'Response.Write sqlstr
		'Response.End
		con.executeQuery(sqlstr)
		sqlstr="Select Max(task_id) From tasks"
		set rsmax = con.getRecordSet(sqlstr)
		If not rsmax.eof Then
			taskId = trim(rsmax(0))
		End If
		set rsmax = Nothing	
		
	ElseIf trim(Request.Form("taskId")<>"") Then ' ����� �����
		taskId = trim(Request.Form("taskId"))
		If activityId <> "" Then				 		
			sqlstr = "Update tasks set "				
			sqlstr=sqlstr & " task_content = '" & sFix(Request.Form("task_content")) & "',"			
			sqlstr=sqlstr & " task_types = '" & sFix(task_types) & "',"
			sqlstr=sqlstr & " task_date = '" & sFix(task_date) & "'"						
			sqlstr=sqlstr & " Where task_id = " & taskId
			'Response.Write sqlstr
			'Response.End	
			con.executeQuery(sqlstr)			
		End If
	End If
    
	If trim(Request.Form("activityId")="") Then ' ������ ������ ����    	    
	    If add_task = true Then ' ������ ��� �� ����� ����� ���� �����
			act_status = "1"
	    Else
			act_status = "2" '������ ���� �� ����� ���� �����
			taskId = "NULL"
	    End If 
	    con.executeQuery("SET DATEFORMAT mdy")
		sqlstr="Insert Into activities (company_id,contact_id,project_id,User_ID,Organization_ID,activity_date,"&_
		" activity_content,activity_types,task_id,activity_status) values (" &_
		companyID & "," & contactID & "," & projectID & "," & UserID & "," & OrgID & ",'" & activity_date & "','" &_
		sFix(Request.Form("activity_content")) & "','" & activity_types & "'," & taskId & ",'" & act_status & "')"
		
		'Response.Write sqlstr
		'Response.End
		con.executeQuery(sqlstr)
		sqlstr="Select Max(activity_id) From activities"
		set rsmax = con.getRecordSet(sqlstr)
		If not rsmax.eof Then
			activityId = trim(rsmax(0))
		End If
		set rsmax = Nothing	
		
	Else
		activityId = trim(Request.Form("activityId"))
		If activityId <> "" Then				 		
			sqlstr = "Update activities set "				
			sqlstr=sqlstr & " activity_content = '" & sFix(Request.Form("activity_content")) & "',"			
			sqlstr=sqlstr & " activity_types = '" & sFix(activity_types) & "',"			
			sqlstr=sqlstr & " activity_date = '" & sFix(activity_date) & "'"						
			sqlstr=sqlstr & " Where activity_id = " & activityId
			'Response.Write sqlstr
			'Response.End	
			con.executeQuery(sqlstr)
			
		End If
	End If
	
	%>
	<SCRIPT LANGUAGE=javascript>
	<!--
		opener.focus();
		opener.window.location.reload();
		self.close();  
	//-->
	</SCRIPT>
  <%		
  End If       
%>
<body style="margin:0px;background:#e6e6e6">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
<%  If Len(activityId) > 0 Then
	sqlstr="Select activity_content,activity_date,activity_types,task_id From activities Where activity_id = " & activityId
	set rstel = con.getRecordSet(sqlstr)
	If not rstel.eof Then
		activity_content=trim(rstel("activity_content"))		
		activity_date = trim(rstel("activity_date"))
		activity_types = trim(rstel("activity_types"))		
		taskId = trim(rstel("task_id"))
		If IsNumeric(taskId) Then
			'sqlstr = "Select task_content,task_date,task_types From tasks_view WHERE task_id = " & taskId
			sqlstr = "EXECUTE get_task '" & OrgID & "','" & taskId & "'"
			set rs_task = con.getRecordSet(sqlstr)
			if not rs_task.eof then
				task_content = trim(rs_task("task_content"))
				task_date = trim(rs_task("task_date"))
				task_types = trim(rs_task("task_types"))
			end if
			set rs_task = Nothing
		End if		
	End If
	set rstel=Nothing
	Else
		activity_date = DateValue(Day(date()) & "/" & Month(Date()) & "/" & Year(Date()))
		task_date = DateValue(Day(date()) & "/" & Month(Date()) & "/" & Year(Date()))
	End If
%>
<tr>	     
	<td class="page_title" dir=rtl>&nbsp;<%If trim(activityID)="" Then%>����� ����� ������<%Else%>����� ����� ������<%End If%>&nbsp;</td>  
</tr>
<tr><td valign=top width="100%" nowrap>

<table border="0" cellpadding="1" cellspacing="0" width="100%" align=center>
<FORM name="formactivity" ACTION="addactivity.asp?add=1" METHOD="post" onSubmit="return CheckFields();">
<input type="hidden" name="activityId" value="<%=activityId%>" ID="activityId">
<input type="hidden" name="taskId"     value="<%=taskId%>" ID="taskId">
<tr>
  <td align=right width=50% nowrap bgcolor="#e6e6e6" valign="top" style="border: 1px solid #808080;border-top:none">
   <table border="0" cellpadding="0" align=center cellspacing="5">     
   <tr><td colspan=2 align=right class="title">����� ����&nbsp;</td></tr>   
   <tr><td colspan=2>
   <table cellpadding=0 cellspacing=0 width=100% border=0>  
   <tr>
   <td align=right colspan=4>
   <table cellpadding=2 cellspacing=1 border=0 width=100% ID="Table3">
   <tr>
   <td align=right width=100%><input type="text" style="width:80;" value="<%=task_date%>" id="task_date" class="Form"  onclick="return DoCal(this)" name="task_date" dir="rtl" ReadOnly></td>
   <td align=right width=65 nowrap>����� ���</td>   
   </tr>
   <tr><td align=right width=100%>
    <select name="ReciverId" dir="rtl" class="norm" style="width:100%" ID="ReciverId">
             <option value=""><%=String(20,"-")%> ��� ���� <%=String(20,"-")%></option>
             <%set UserList=con.GetRecordSet("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME FROM Users WHERE ORGANIZATION_ID = " & OrgID & " ORDER BY FIRSTNAME + ' ' + LASTNAME")
             do while not UserList.EOF
                selUserID=UserList(0)
                selUserName=UserList(1)%>
                <option value="<%=selUserID%>" <%if trim(ReciverId)=trim(selUserID) then%> selected <%end if%>><%=selUserName%></option>
                <%
                UserList.MoveNext
                loop
                set UserList=Nothing%>
           </select>
    </td>
    <td align=right width=65 nowrap>... ��</td>
    </tr>    
    </table></td></tr>              
   <%
		  	sqlstr="Select activity_type_name, activity_type_id From activity_types WHERE Organization_Id = " & trim(Request.Cookies("bizpegasus")("OrgID")) & " Order By activity_type_id"
		  	set rssub = con.getRecordSet(sqlstr)		 
		  	If not rssub.eof Then
		  	   arrSub = rssub.getRows()
		  	   recCount = rssub.RecordCount	
		  	   set rssub=Nothing
		  	   i=0			
		  	While i+1<=Ubound(arrSub,2)
		  		selSubgectId=trim(arrSub(1,i+1))
		  		selactivityName=trim(arrSub(0,i+1))	  							
		  %>
		 
		  <tr>
		  <%If i<=Ubound(arrSub,2) Then%>
		  <td align=right nowrap>&nbsp;<%=selactivityName%>&nbsp;</td>
		  <td width=15  align=center nowrap><input type=checkbox id="task_type<%=selSubgectId%>" name="task_type<%=selSubgectId%>" <%If Instr(task_types,selSubgectId) > 0 Then%> checked <%End If%> OnClick="return add_remove('task_types',this,<%=selSubgectId%>)"></td>		 
		  <%Else%>
		  <td>&nbsp;</td><td>&nbsp;</td>
		  <%End If	
		        If i<=Ubound(arrSub,2) Then		  	    
		  	    selSubgectId=trim(arrSub(1,i))
		  		selactivityName=trim(arrSub(0,i))		  		
		  %>		  
		   <td align=right nowrap>&nbsp;<%=selactivityName%>&nbsp;</td>
		   <td width=15  align=center nowrap><input type=checkbox id="task_type<%=selSubgectId%>" name="task_type<%=selSubgectId%>" <%If InStr(task_types,selSubgectId) > 0 Then%> checked <%End If%> OnClick="return add_remove('task_types',this,<%=selSubgectId%>)"></td>									  
		    <%Else%>
		  <td>&nbsp;</td><td>&nbsp;</td>
		  <%End If%>
		   </tr>	
		  <%  i = i+2		
		  	  Wend		  	
			  End If		
          %>         
        <input type=hidden name="task_types" id="task_types" value="<%=task_types%>">   
      </table>        
 </td></tr>
 <tr><td valign="top" align="right">����</td></tr>
 <tr>             
	<td align=right bgcolor="#e6e6e6" valign="top" nowrap>
	<textarea id="task_content" name="task_content" dir=rtl cols=40 rows=5><%=task_content%></textarea>
	</td>	
</tr>         
</table></td>
<td align=right width=50% nowrap bgcolor="#e6e6e6" valign="top" style="border: 1px solid #808080;border-left:none;border-top:none">
   <table border="0" cellpadding="0" align=center cellspacing="5" ID="Table1">                
   <tr><td colspan=2 align=right class="title">����� ������&nbsp;</td></tr>
   <tr><td colspan=2>
       <table cellpadding=0 cellspacing=0 width=100% border=0 ID="Table2"> 
       <tr>
		<td align=right colspan=4>
		<table cellpadding=2 cellspacing=1 border=0 width=100% ID="Table4">
		<tr>
		<td align=right><input type="text" style="width:80;" value="<%=activity_date%>" id="activity_date" name="activity_date" dir="rtl" class="Form"  onclick="return DoCal(this)" readOnly></td>   
		<td align=right width=65 nowrap>�����</td>   
	   </tr>	  
	 <tr><td align=right width=100%>
    <select name="SenderID" dir="rtl" class="norm" style="width:100%" ID="SenderID" disabled>
             <option value=""><%=String(20,"-")%> ��� ���� <%=String(20,"-")%></option>
             <%set UserList=con.GetRecordSet("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME FROM Users WHERE ORGANIZATION_ID = " & OrgID & " ORDER BY FIRSTNAME + ' ' + LASTNAME")
             do while not UserList.EOF
                selUserID=UserList(0)
                selUserName=UserList(1)%>
                <option value="<%=selUserID%>" <%if trim(UserId)=trim(selUserID) then%> selected <%end if%>><%=selUserName%></option>
                <%
                UserList.MoveNext
                loop
                set UserList=Nothing%>
           </select>
    </td>
    <td align=right width=65 nowrap>�����</td>
    </tr>         
    </table></td></tr>                       
          <%
		  	sqlstr="Select activity_type_name, activity_type_id From activity_types WHERE Organization_Id = " & trim(Request.Cookies("bizpegasus")("OrgID")) & " Order By activity_type_id"
		  	set rssub = con.getRecordSet(sqlstr)		 
		  	If not rssub.eof Then
		  	   arrSub = rssub.getRows()
		  	   recCount = rssub.RecordCount	
		  	   set rssub=Nothing
		  	   i=0			
		  	While i+1<=Ubound(arrSub,2)
		  		selSubgectId=trim(arrSub(1,i+1))
		  		selactivityName=trim(arrSub(0,i+1))	  							
		  %>
		 
		  <tr>
		  <%If i<=Ubound(arrSub,2) Then%>
		  <td align=right nowrap>&nbsp;<%=selactivityName%>&nbsp;</td>
		  <td width=15  align=center nowrap><input type=checkbox id="activity_type<%=selSubgectId%>" name="activity_type<%=selSubgectId%>" <%If Instr(activity_types,selSubgectId) > 0 Then%> checked <%End If%> OnClick="return add_remove('activity_types',this,<%=selSubgectId%>)"></td>		 
		  <%Else%>
		  <td>&nbsp;</td><td>&nbsp;</td>
		  <%End If	
		        If i<=Ubound(arrSub,2) Then		  	    
		  	    selSubgectId=trim(arrSub(1,i))
		  		selactivityName=trim(arrSub(0,i))		  		
		  %>		  
		   <td align=right nowrap>&nbsp;<%=selactivityName%>&nbsp;</td>
		   <td width=15  align=center nowrap><input type=checkbox id="activity_type<%=selSubgectId%>" name="activity_type<%=selSubgectId%>" <%If InStr(activity_types,selSubgectId) > 0 Then%> checked <%End If%> OnClick="return add_remove('activity_types',this,<%=selSubgectId%>)"></td>									  
		    <%Else%>
		  <td>&nbsp;</td><td>&nbsp;</td>
		  <%End If%>
		   </tr>	
		  <%  i = i+2		
		  	  Wend		  	
			  End If		
          %>         
        <input type=hidden name="activity_types" id="activity_types" value="<%=activity_types%>">   
        </table>        
 </td></tr>
 <tr><td valign="top" align="right">����</td></tr>
 <tr>             
	<td align=right bgcolor="#e6e6e6" valign="top" nowrap>
	<textarea id="activity_content" name="activity_content" dir=rtl cols=40 rows=5><%=activity_content%></textarea>
	</td>	
</tr>         
</table></td></tr>
<tr>
   <td align=right colspan=4 bgcolor="#C9C9C9"  style="border: 1px solid #808080;border-top:none" style="padding:3px">
   <table cellpadding=2 cellspacing=1 border=0 width=360 align=center>
    <tr><td align=right width=100%>
    <select name="companyId" dir="rtl" class="norm" style="width:100%" ID="companyId" onchange="window.location.href='addactivity.asp?companyId='+this.value+'&contactID='+contactID.value+'&ReciverId='+ReciverId.value">
             <option value=""><%=String(22,"-")%> ��� ���� <%=String(22,"-")%></option>
             <%set CompList=con.GetRecordSet("SELECT company_ID,company_Name FROM companies WHERE ORGANIZATION_ID = " & OrgID & " ORDER BY private_flag DESC,company_Name")
             do while not CompList.EOF
                selCompID=CompList(0)
                selCompName=CompList(1)%>
                <option value="<%=selCompID%>" <%if trim(companyId)=trim(selCompID) then%> selected <%end if%>><%=selCompName%></option>
                <%
                CompList.MoveNext
                loop
                set CompList=Nothing%>
           </select>
       </td>
       <td align=right width=90 nowrap>����� �����</td>
       </tr> 
      <%If trim(private_flag) = "0" Then%>
      <tr><td align=right width=100%>
      <select name="contactID" dir="rtl" class="norm" style="width:100%" ID="contactID">
             <option value=""><%=String(20,"-")%> ��� ��� ��� <%=String(20,"-")%></option>
             <%set ContList=con.GetRecordSet("SELECT contact_ID,contact_Name FROM contacts WHERE company_id IN (Select distinct company_id from companies WHERE ORGANIZATION_ID = " & OrgID & where_company & ") ORDER BY contact_Name")
             do while not ContList.EOF
                selContID=ContList(0)
                selContName=ContList(1)%>
                <option value="<%=selContID%>" <%if trim(contactID)=trim(selContID) then%> selected <%end if%>><%=selContName%></option>
                <%
                ContList.MoveNext
                loop
                set ContList=Nothing%>
           </select>
       </td>
       <td align=right width=90 nowrap>����� ���� ���</td>
       </tr>   
    <%Else%>
    <input type=hidden name="contactID" id="contactID" value="<%=contactID%>">
    <%End If%> 
    <%If trim(companyID) <> "" Then%>
    <TR>
		<td align=right width=100%>
		<select name="project_id" id="project_id" class="norm" dir=rtl style="width:100%">
		<option value=""><%=String(20,"-")%> ��� ������ <%=String(20,"-")%></option>		
		<%  
			con.executeQuery("SET DATEFORMAT dmy")
			sqlstr = "Select PROJECT_ID, PROJECT_NAME FROM PROJECTS WHERE ORGANIZATION_ID = " & OrgID & where_company & " AND active = 1 Order BY PROJECT_NAME"
			'Response.Write sqlstr
			set rs_comp = con.getRecordSet(sqlstr)
			While not rs_comp.eof
		%>
			<option value="<%=rs_comp(0)%>" <%if trim(rs_comp(0))=trim(projectID) then%> selected <%end if%>><%=rs_comp(1)%></option>		
		<%
			rs_comp.moveNext
			Wend
			set rs_comp = Nothing
		%>		
		</select>
		</TD>
		 <td align=right width=90 nowrap>����� �������</td>
	</TR>
	<%End If%>
   </table>
   </td>
  </tr>          
</form>
</table></td></tr>
<tr><td align=center colspan="2" height=10 nowrap></td></tr>
<tr><td align=center colspan="2">
<input type=button value="�����" class="but_menu" style="width:100" onclick="window.close();">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input type=button value="�����" class="but_menu" style="width:100" onClick="return CheckFields('submit')"></td></tr>
</td></tr>
<tr><td align=center colspan="2" height=10 nowrap></td></tr>
</table>
</div>
</body>
<%set con=Nothing%>
</html>

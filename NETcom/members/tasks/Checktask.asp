<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
appid=request("appid")
arr_StatusT = Array("","חדש","בטיפול","סגור")	
 OrgID=264
   ContactId = trim(Request("ContactId"))
set listContact=con.GetRecordSet("EXEC dbo.contacts_contact_details @ContactId=" & ContactId & ", @OrgID=" & OrgID)
   if not listContact.EOF then 
      ContactId = cLng(listContact("contact_ID"))
      companyID = cLng(listContact("company_ID"))
       contacter = trim(listContact("CONTACT_NAME"))
      end if
      set  listContact=nothing
  	sqlstr = "SET DATEFORMAT mdy; select (USERS_1.FIRSTNAME + CHAR(32) + USERS_1.LASTNAME) AS reciver_name, " &_
   " (dbo.USERS.FIRSTNAME + CHAR(32) + dbo.USERS.LASTNAME) AS sender_name, " &_
   "  dbo.tasks.task_date,dbo.tasks.task_content,task_open_date,  " &_  
   "  dbo.tasks.USER_ID, dbo.tasks.ORGANIZATION_ID,dbo.tasks.company_id, dbo.tasks.contact_id, " &_
   "  dbo.tasks.task_id, dbo.tasks.task_status FROM  dbo.tasks INNER JOIN" &_
   "  dbo.USERS USERS_1 ON dbo.tasks.reciver_id = USERS_1.USER_ID INNER JOIN " &_
   "  dbo.USERS ON dbo.tasks.USER_ID = dbo.USERS.USER_ID  " &_
   "  WHERE dbo.tasks.Contact_id = " & ContactId &" and task_status<>3"
	'Response.Write sqlstr
	'	Response.End   
		set tasksList = con.getRecordSet(sqlstr)
	 
       %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<!-- #include file="../../title_meta_inc.asp" -->
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script>
function SendTask(ContactId,companyID,taskID)
{
	h = parseInt(530);
		w = parseInt(470);
		//alert("fff")
		//alert(window.opener.location.href)
	window.document.location.href="../tasks/addtask.asp?ContactId=" + ContactId + "&companyId=" + companyID + "&taskID=" + taskID +"&appid=<%=appid%>"
window.resizeTo(530,600)
//window.open("../tasks/addtask.asp?ContactId=" + ContactId + "&companyId=" + companyID + "&taskID=" + taskID +"&appid=<%=appid%>", "_self" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
//self.close()
}

		</script>
	</head>
	<body style="margin:0px;background:#e6e6e6" onload="self.focus()">
		<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" ID="Table1">
			<tr>
				<td height="30"></td>
			</tr>
			<tr>
				<td width="100%" align="center" dir="rtl">&nbsp;<span style="font-size:16pt"> שים לב, 
						ללקוח
						<%=contacter%>
						קיימות משימות בסטטוסים חדש/בטיפול &nbsp; </span>
				</td>
			</tr>
			<tr>
				<td height="30"></td>
			</tr>
			<tr>
				<td align=center>
					<table width="98%" border=0 cellpadding=2 cellspacing=2  dir="<%=dir_var%>" ID="Table3">
						<tr  height=25>
							<td align="center" class="title_sort" width="29" nowrap>&nbsp;</td>
							<td align="right" align="right" dir="rtl" class="title_sort" width="265" nowrap><span id="word19" name="word19">תוכן</span>&nbsp;</td>
							<td width="90" align="center" dir="rtl" nowrap class="title_sort"><span id="word21">אל</span></td>
							<td width="90" align="center" dir="rtl" nowrap class="title_sort"><span id="word22" name="word22">מאת</span>&nbsp;</td>
							<td align="center" dir="rtl" width="75" nowrap class="title_sort"><span id="word23" name="word23">מתאריך</span></td>
							<td align="center" dir="rtl" width="45" nowrap class="title_sort">&nbsp;<span id="word24" name="word24">סט'</span>&nbsp;</td>
						</tr>
						<% while not tasksList.EOF       
		taskId = trim(tasksList(1))				
		task_date = trim(tasksList("task_open_date"))
		task_status = trim(tasksList("task_status"))
		sender_name = trim(tasksList("sender_name"))
		reciver_name = trim(tasksList("reciver_name"))      
		task_content = trim(tasksList("task_content"))          
	%>
						<tr height=35>
							<td align="center" class="title_sort" width="29" nowrap>&nbsp;</td>
							<td align="right" align="right" dir="rtl" class="title_sort" width="265" nowrap><span id="Span1" name="word19"><%=task_content%></span>&nbsp;</td>
							<td width="90" align="center" dir="rtl" nowrap class="title_sort"><span id="Span2"><%=reciver_name%></span></td>
							<td width="90" align="center" dir="rtl" nowrap class="title_sort"><span id="Span3" name="word22"><%=sender_name%></span>&nbsp;</td>
							<td align="right" dir="ltr" width="75" nowrap class="title_sort"><span id="Span4" name="word23"><%=task_date%></span></td>
							<td align="center" dir="rtl" width=45 nowrap class="task_status_num<%=task_status%>">&nbsp;<span id="Span5" name="word24"><%=arr_StatusT(task_status)%></span>&nbsp;</td>
						</tr>
						<%   tasksList.MoveNext
    Wend%>
					
					</table>
					<br clear=all>
					<table border=0 cellpadding=0 cellspacing=0>
						<tr>
							<td colspan="6" height="50" ></td>
						</tr>
						<tr>
							<td align="center" colspan="6" style="background-color:#e6e6e6"><input style="width:100px" type="submit" class="but_menu" value="הוסף המשימה" id="Button4" name="Button4" onclick="return SendTask('<%=ContactId%>','<%=companyID%>','')">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input style="width:140px" type="submit" class="but_menu" value="בטל הוספת המשימה" id="" name="Button4"
									onclick="javascript:self.close();">
							</td>
						</tr>
					</table>
	</body>
</html>

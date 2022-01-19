<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
appealsId=request("appealsId")
arrAppeals = Split(appealsId, ",")   		
    
arr_StatusT = Array("","חדש","בטיפול","סגור")	
OrgID=264

	
       %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<!-- #include file="../../title_meta_inc.asp" -->
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.8.3/jquery.min.js"></script>

		<script>
function SendTask(appealsId)
{
var st=0;
$('#form1 select').each(
    function(index){  
        var input = $(this);
       // alert('Name: ' + input.attr('name') + 'Value: ' + input.val());
        if (input.val()=='0')
        {
        st=1
        }
    }
);
if (st==1) 
{alert("אנא ציין סטטוס המשך שליחת המשימה לכל הלקוחות")
}
else
{
var delIds=0
$('#form1 select').each(
    function(index){  
        var input = $(this);
          if (input.val()=='2')
        {
        if (delIds==0)
        {
            delIds= input.attr('name').substring(4,input.attr('name').length)
        }
      //  input.attr('name').substring(4,len(input.attr('name')));
      else
      {
        delIds=delIds+','+  input.attr('name').substring(4,input.attr('name').length)
      }  
        }
    }
);
//alert ("delIds="+delIds)
//alert(document.getElementById("appIds").value)
if (document.getElementById("appIds").value==delIds)
{
alert("שים לב, לא נמצאו טפסים להוספת המשימה")
return false;
}
	h = parseInt(530);
		w = parseInt(470);
		window.document.location.href="../tasks/addtasks.asp?appealsId=" + appealsId+"&delIds="+delIds
window.resizeTo(530,600)

	//window.open("../tasks/addtasks.asp?appealsId=" + appealsId+"&delIds="+delIds, "T_WindN" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
//self.close()
}
}

</script>
	</head>
	<body style="margin:0px;background:#e6e6e6" onload="self.focus()">
			<FORM method=get id="form1" name="form1" target="_self">
			<input type=hidden id="appIds" name="appIds" value="<%=appealsId%>">
		<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" ID="Table1">
			<tr>
				<td height="10"></td>
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
<%				For aa=0 To Ubound(arrAppeals)
				
			task_appeal_id	= arrAppeals(aa)
	sqlCont="select Appeals.Contact_Id,Appeals.company_ID,Contacts.Contact_Name from Appeals  left join Contacts on Appeals.Contact_Id=Contacts.Contact_Id where Appeal_id=" & task_appeal_id  &" and dbo.Appeals.CONTACT_ID is not null"
	'response.Write "1=" &sqlCont &"<BR>"
	set contactL=con.getRecordSet(sqlCont)
	if not contactL.eof then
	 ContactId = contactL("contact_ID")
     companyID = contactL("company_ID")
     contacter = trim(contactL("CONTACT_NAME"))
			
	sqlstr = "SET DATEFORMAT mdy; select (USERS_1.FIRSTNAME + CHAR(32) + USERS_1.LASTNAME) AS reciver_name, " &_
   " (dbo.USERS.FIRSTNAME + CHAR(32) + dbo.USERS.LASTNAME) AS sender_name, " &_
   "  dbo.tasks.task_date,dbo.tasks.task_content,task_open_date,  " &_  
   "  dbo.tasks.USER_ID, dbo.tasks.ORGANIZATION_ID,dbo.tasks.company_id, dbo.tasks.contact_id, " &_
   "  dbo.tasks.task_id, dbo.tasks.task_status FROM  dbo.tasks INNER JOIN" &_
   "  dbo.USERS USERS_1 ON dbo.tasks.reciver_id = USERS_1.USER_ID INNER JOIN " &_
   "  dbo.USERS ON dbo.tasks.USER_ID = dbo.USERS.USER_ID  " &_
   "  WHERE dbo.tasks.Contact_id = " & ContactId &" and task_status<>3 and dbo.tasks.CONTACT_ID is not null and datediff(mm,GetDate(), dbo.tasks.task_date)>=-2"
	'response.write "<BR>" & sqlstr
	
		set tasksList = con.getRecordSet(sqlstr)
		if not   tasksList.EOF     then
		%>
			<tr><td colspan=6 height=20></td></tr>

		<tr>
				<td width="100%" align="center" dir="rtl" colspan=6>&nbsp;<span style="font-size:14pt"> שים לב, 
						ללקוח
						<%=contacter%>
						קיימות משימות בסטטוסים חדש/בטיפול &nbsp; </span>
				</td>
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
    <tr><td colspan=6 align=center>
    <table border=0 cellpadding=0 cellspacing=0>
    <tr><td><select id="app_<%=task_appeal_id%>" name="app_<%=task_appeal_id%>">
    <option value="0">בחר</option>
      <option value="1">כן</option>
        <option value="2">לא</option>
    </select></td><td dir=rtl  style="font-size:14pt">&nbsp;האם להמשיך להוסיף משימה?&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>
    </table>
    </td></tr>
				<%set tasksList=Nothing%>
				<%end if%>
							<%set contactL=Nothing%>
				<%end if%>
				
					<%Next%>	
					</table>
					<br clear=all>
					<table border=0 cellpadding=0 cellspacing=0 ID="Table2">
						<tr>
							<td colspan="6" height="50" ></td>
						</tr>
						<tr>
							<td align="center" colspan="3" style="background-color:#736BA6"><a style="width:100px;LINE-HEIGHT:23px;LETTER-SPACING: 1px;color:ffffff;text-decoration:none;FONT-WEIGHT: bolder"  href="javascript:SendTask('<%=appealsId%>')">&nbsp;&nbsp;&nbsp;הוסף המשימה&nbsp;</a>&nbsp;&nbsp;&nbsp;&nbsp;</td>
							<td align="center" colspan="3">&nbsp;&nbsp;&nbsp;&nbsp;<input style="width:140px" type="submit" class="but_menu" value="בטל הוספת המשימה" id="" name="Button4"
									onclick="javascript:self.close();">
							</td>
						</tr>
					</table>
					</form>
	</body>
</html>

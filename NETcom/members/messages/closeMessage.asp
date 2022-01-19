<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<script language="Javascript">

var timeout;

function timeout_trigger() {
    document.getElementById("button_read").disabled = false;
}


function timeout_init() {
    timeout = setTimeout('timeout_trigger()', 3000);
}
function updateStatus(MessageId,st)
{
	document.location.href = "CloseMessage.asp?messageId=" + MessageId+'&updSt='+st;
			return false;
}

</script>


<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%
	  messageId=trim(Request("messageId"))  
	   if Request.QueryString("updSt")<>"" then
	 	sqlstr = "Update messages SET message_status = '"& Request.QueryString("updSt") &"' WHERE message_id = " & messageId
	'response.Write 	sqlstr
	'response.end
		con.executeQuery(sqlstr)
	  %>
	  	  <script>
	  	  opener.window.location.href = opener.window.location.href;
	  	  self.close();
	  	  </script>
	<%response.end
end if
	  
	  UserID=trim(trim(Request.Cookies("bizpegasus")("UserID")))
	  OrgID=trim(trim(Request.Cookies("bizpegasus")("OrgID")))
	  SURVEYS = trim(Request.Cookies("bizpegasus")("SURVEYS"))
	  COMPANIES = trim(Request.Cookies("bizpegasus")("COMPANIES"))
	  lang_id = trim(Request.Cookies("bizpegasus")("LANGID"))
	  if lang_id = "1" then
		arr_Status = Array("","חדש","נקרא","תגובה")	
	 else
	    arr_Status = Array("","new","read","replay")	
	 end if
	 If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	 End If
	 If lang_id = 2 Then
		dir_var = "rtl" : align_var = "left" : dir_obj_var = "ltr" : self_name = "Self"
	 Else
		dir_var = "ltr" : align_var = "right" : dir_obj_var = "rtl" : self_name = "עצמי"
	 End If		%>
<!-- #include file="../../title_meta_inc.asp" -->
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<script language="javascript" type="text/javascript">
function CheckFields(flag)
{  
  /* if(document.all("message_close_date").value=='' )
   {
	<%
		If trim(lang_id) = "1" Then
			str_alert = "! נא למלא תאריך תגובה"
		Else
			str_alert = "Please insert the date!"
		End If   
	%>   
		window.alert("<%=str_alert%>");
		return false;
   }  */ 
   
   if(document.all("message_replay").value=='' )
   {
   	<%
		If trim(lang_id) = "1" Then
			str_alert = "! נא למלא תוכן תגובה"
		Else
			str_alert = "Please insert the content!"
		End If   
	%>   
		window.alert("<%=str_alert%>");	
		return false;
   }
 
	if(flag == "2")
	{
		document.form_closing.action = "closemessage.asp?parentID=<%=messageId%>";
	}   
    
   document.form_closing.submit();               
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
function closeWin()
{
	opener.focus();
	opener.window.location.href = opener.window.location.href;
	self.close();
}
//-->
</script>  
</head>
<%  
  set upl=Server.CreateObject("SoftArtisans.FileUp")
  If upl.Form("close") <> nil Then 	    
		
	messageId = trim(upl.Form("messageId"))
		
	'	If upl.Form("message_close_date") <> nil And IsDate(upl.Form("message_close_date")) Then    
	'		message_close_date = Day(trim(upl.Form("message_close_date"))) & "/" & Month(trim(upl.Form("message_close_date"))) & "/" & Year(trim(upl.Form("message_close_date")))			
	'	else
	'	message_close_date
	'	End If  
	message_close_date=FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4)
		
		If messageId <> "" And trim(upl.Form("close")) = "1" Then
		    con.executeQuery("SET DATEFORMAT DMY")				 		
			sqlstr = "Update messages set message_status = '3', " &_
			" message_replay = '" & sFix(upl.Form("message_replay")) & "'," &_
			" message_close_date = '" & message_close_date & "'" &_					
			" Where message_id = " & messageId
			'Response.Write sqlstr
			'Response.End	
			con.executeQuery(sqlstr)
			
				
		    If trim(messageId) <> "" Then
			sqlstr = "EXECUTE get_message '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & messageId & "'"
			'response.Write sqlstr
			'response.end
			set rs_message = con.getRecordSet(sqlstr)
			if not rs_message.eof then
				message_content = trim(rs_message("message_content"))
				message_date = trim(rs_message("message_date"))
				message_open_date = trim(rs_message("message_open_date"))
				message_types = trim(rs_message("message_types"))
				message_status = trim(rs_message("message_status"))	
				activityId = trim(rs_message("parent_id"))
				message_sender_name = trim(rs_message("sender_name"))
				message_reciver_name = trim(rs_message("reciver_name"))
				message_replay = trim(rs_message("message_replay"))
				message_close_date = trim(rs_message("message_close_date"))
				message_sender_id = trim(rs_message("User_ID"))
				message_reciver_id = trim(rs_message("reciver_id"))
					parentID = trim(rs_message("parent_ID"))
				
				If IsDate(message_date) Then
					message_date = Day(message_date) & "/" & Month(message_date) & "/" & Year(message_date)
				End If	
				
				If IsDate(message_open_date) Then
					message_open_date = FormatDateTime(message_open_date,2) & " " & FormatDateTime(message_open_date,4)
				End If				
				
	
				sqlstr = "Exec dbo.get_message_types '"&messageId&"','"&OrgID&"'"
				set rs_message_types = con.getRecordSet(sqlstr)
				If not rs_message_types.eof Then
					message_types_names = rs_message_types.getString(,,",",",")
				Else
					message_types_names = ""
				End If		
			
				If Len(message_types_names) > 0 Then
					message_types_names = Left(message_types_names,(Len(message_types_names)-1))
				End If								
				
				sqlstr = "Select EMAIL From USERS Where USER_ID = " & message_reciver_id
				set rswrk = con.getRecordSet(sqlstr)
				If not rswrk.eof Then
					fromEmail = trim(rswrk("EMAIL"))
				End If
				set rswrk = Nothing						
				
			end if
			set rs_message = Nothing	 	
			
		
		

    	  
	 'שליחת סגירת המשימה לשולחי משימות הקודמות 

		
        

	End If	
End If	
    If Request.QueryString("parentID") = nil Then	
	%>
	<SCRIPT LANGUAGE=javascript>
	<!--
		if(opener)
		{
			opener.focus();
			opener.window.location.href = opener.window.location.href;
		}
		self.close();

	//-->
	</SCRIPT>
	<%Else%>
  <script language=javascript>
	<!--
		document.location.href = "addmessage.asp?parentID=<%=messageId%>";
	//-->
  </script>	
  <%End If		
  End If  
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 27 Order By word_id"				
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
<body style="margin:0px;" onload="window.focus();timeout_init();">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" bgcolor=#e6e6e6 dir="<%=dir_var%>" ID="Table1">
<%
  If trim(messageId) <> "" Then
	Set rs_message = Server.CreateObject("ADODB.RECORDSET") 
    sqlstr = "EXECUTE get_message '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & messageId & "'"
    set rs_message = con.getRecordSet(sqlstr)
	if not rs_message.eof then
		message_content = trim(rs_message("message_content"))
		message_date = trim(rs_message("message_date"))
		message_open_date = trim(rs_message("message_open_date"))
		message_types = trim(rs_message("message_types"))
		message_status = trim(rs_message("message_status"))	
		activityId = trim(rs_message("parent_id"))
		message_sender_name = trim(rs_message("sender_name"))		
		message_reciver_name = trim(rs_message("reciver_name"))
		message_replay = trim(rs_message("message_replay"))
		message_close_date = trim(rs_message("message_close_date"))
		message_sender_id = trim(rs_message("User_ID"))
		message_reciver_id = trim(rs_message("reciver_id"))
			parentID = trim(rs_message("parent_ID"))
		childID = trim(rs_message("childID"))
		
		If IsDate(message_date) Then
			message_date = Day(message_date) & "/" & Month(message_date) & "/" & Year(message_date)
		End If	
		
		If IsDate(message_open_date) Then
			message_open_date = FormatDateTime(message_open_date,2) & " " & FormatDateTime(message_open_date,4)
		End If		
		
	
	
		sqlstr = "Exec dbo.get_message_types '"&messageId&"','"&OrgID&"'"
		set rs_message_types = con.getRecordSet(sqlstr)
		If not rs_message_types.eof Then
			message_types_names = rs_message_types.getString(,,",",",")
		Else
			message_types_names = ""
		End If		
		
		If Len(message_types_names) > 0 Then
			message_types_names = Left(message_types_names,(Len(message_types_names)-1))
		End If
		
		If trim(message_status) = "1" And message_reciver_id = trim(UserID) Then 'משימה נפתחה פעם ראשונה  ---- 13.11.13 Faina
		
			'sqlstr = "Update messages SET message_status = '2' WHERE message_id = " & messageId
			'con.executeQuery(sqlstr)
			
			xmlFilePath = "../../../download/xml_messages/bizpower_messages_in.xml"
			'------ start deleting the new message from XML file ------
			set objDOM = Server.CreateObject("Microsoft.XMLDOM")
			objDom.async = false			
			if objDOM.load(server.MapPath(xmlFilePath)) then
				set objNodes = objDOM.getElementsByTagName("message")
				for j=0 to objNodes.length-1
					set objmessage = objNodes.item(j)
					node_message_id = objmessage.attributes.getNamedItem("message_ID").text					
					node_user_id = objmessage.attributes.getNamedItem("Reciver_ID").text
					if trim(messageId) = trim(node_message_id) And trim(node_user_id) = trim(UserID)  Then					
						objDOM.documentElement.removeChild(objmessage)
						exit for
					else
						set objmessage = nothing
					end if
				next
				Set objNodes = nothing
				set objmessage = nothing
				objDom.save server.MapPath(xmlFilePath)
			end if
			set objDOM = nothing
		' ------ end  deleting the new message from XML file ------
		End If	
	
	If trim(message_status) = "3" And message_sender_id = trim(UserID) Then 'משימ
	
			xmlFilePath = "../../../download/xml_messages/bizpower_messages_out.xml"
			'------ start deleting the new message from XML file ------
			set objDOM = Server.CreateObject("Microsoft.XMLDOM")
			objDom.async = false			
			if objDOM.load(server.MapPath(xmlFilePath)) then
				set objNodes = objDOM.getElementsByTagName("message")
				for j=0 to objNodes.length-1
					set objmessage = objNodes.item(j)
					node_message_id = objmessage.attributes.getNamedItem("message_ID").text					
					node_user_id = objmessage.attributes.getNamedItem("Sender_ID").text
					if trim(messageId) = trim(node_message_id) And trim(node_user_id) = trim(UserID) Then					
						objDOM.documentElement.removeChild(objmessage)
						exit for
					else
						set objmessage = nothing
					end if
				next
				Set objNodes = nothing
				set objmessage = nothing
				objDom.save server.MapPath(xmlFilePath)
			end if
			set objDOM = nothing
		' ------ end  deleting the new message from XML file ------
		End If
	end if
	set rs_message = Nothing	 	
  End If 
%>
<tr>	     
	<td class="page_title" dir="<%=dir_obj_var%>">&nbsp;<%=trim(Request.Cookies("bizpegasus")("messagesOne"))%> BIZPOWER&nbsp;</td>  
</tr>
<tr><td valign=top width="100%" nowrap>

<table border="0" cellpadding="1" cellspacing="0" width=100% align=center ID="Table2">
<tr>
   <td align="<%=align_var%>" width=100% nowrap bgcolor="#e6e6e6" valign="top" style="border: 1px solid #808080;border-top:none">
    <table width=100% border="0" cellpadding="0" align=center cellspacing="1" ID="Table3"> 
	<!--tr><td colspan=3 align="<%=align_var%>" class="title">פרטי המשימה&nbsp;</td></tr-->
	<tr><td colspan=3 height="5" nowrap></td></tr>
	<tr>
	<td align="<%=align_var%>" class="Form_R" colspan=3 nowrap dir=rtl>
	<table border=0 cellpadding=0 cellspacing=0>
	<tr>
	<td>מאת:&nbsp;</td>
	<td dir=rtl><%=message_sender_name%></td>
	<td>&nbsp;<%=message_date%></td>

	</tr>
	
	</table></td>
		</tr>
	<%if false then%> <tr>
		<td>
        <%If trim(messageId) <> "" And (trim(childID) <> "" Or trim(parentID) <> "") Then%>        
        <A class="but_menu" href="#" style="width:150" onclick='window.open("message_history.asp?messageId=<%=messageId%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=800,height=480,align=center,resizable=0");'>&nbsp;<!--היסטוריה--><%=arrTitles(21)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("messagesOne"))%></a>
        <%End If%>
        </td>
        <td align="<%=align_var%>"><span style="width:75;text-align:center" class="task_status_num<%=message_status%>">
         <%=arr_Status(message_status)%></span>
         </td>
         <td align="<%=align_var%>" width=90 style="padding-right:10px;padding-left:10px" nowrap><b><!--סטטוס--><%=arrTitles(2)%></b></td>
     </tr>
		<tr>
	<td align="<%=align_var%>" colspan=2 nowrap dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:150;"><%=message_sender_name%></span></td>
	<td align="<%=align_var%>" style="padding-right:10px;padding-left:10px" nowrap><b><!--מאת--><%=arrTitles(4)%></b></td>
	</tr><%end if%>
	<%if false then%>
	<tr>
	<td align="<%=align_var%>" colspan=2 nowrap dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:150;"><%=message_reciver_name%></span></td>
	<td align="<%=align_var%>" style="padding-right:10px;padding-left:10px" nowrap><b><!--אל--><%=arrTitles(5)%></b></td>
	</tr>
	<tr>
	<td align="<%=align_var%>" colspan=2 dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:280;line-height:120%;"><%=message_types_names%></span></td>
	<td align="<%=align_var%>" style="padding-right:10px;padding-left:10px" nowrap><b><!--סוג--><%=arrTitles(6)%></b></td>
	</tr><%end if%>
	<tr>
	<td align="<%=align_var%>" colspan=3 valign=top dir="<%=dir_obj_var%>"><span  class="Form_R" style="width:100%;height:50;overflow : auto;""><%=breaks(message_content)%></span></td>
	<tr><td height=5 colspan=3 nowrap></td></tr>

	</table></td></tr>
	
	<%If ((trim(message_status) = "1" Or trim(message_status) = "2") And trim(message_reciver_id) = trim(UserID)) Or trim(message_status) = "3" Then%>
	<FORM name="form_closing" ACTION="closemessage.asp" METHOD="post" enctype="multipart/form-data" ID="Form1">
	<input type="hidden" name="messageId"  value="<%=messageId%>" ID="messageId">
	<input type="hidden" name="deleteFile" value="0" ID="deleteFile">
	<input type="hidden" name="close" value="1" ID="close">
	<tr>
	<td align="<%=align_var%>" width=100% nowrap bgcolor="#C9C9C9" valign="top" style="border: 1px solid #808080;border-top:none">
	<table border="0" cellpadding="0" align=center cellspacing="1" ID="Table4">     
	<!--tr><td colspan=2 align="<%=align_var%>" class="title">פרטי סגירה</td></tr-->      
	<tr><td height=5 colspan=2 nowrap></td></tr>
	<%If IsNull(message_close_date) Then
		message_close_date = Date()
	  End If 
	%> 
	<%if false then%>
	<tr>
	<td align="<%=align_var%>" width=100%><input type="text" style="width:80;" value="<%=message_close_date%>" id="message_close_date"  class="Form_R" readonly name="message_close_date" dir="rtl" ReadOnly></td>
	<td align="<%=align_var%>" width=90 style="padding-right:10px;padding-left:10px" nowrap><b>תאריך תגובה</b></td>   
	</tr> <%end if%>  	
	<tr>             
		<td align="<%=align_var%>" valign="top" nowrap>
		<textarea id="message_replay" name="message_replay" dir="<%=dir_obj_var%>" style="width:340;height:50;overflow : auto;" <%If trim(message_reciver_id) = trim(UserID)  Then%> class="Form" <%Else%> class="Form_R" readonly <%End If%>><%=message_replay%></textarea>
		</td>	
		<td valign="top" style="padding-right:10px;padding-left:10px" align="<%=align_var%>"><b>תוכן תגובה<!--תוכן סגירה--><%'=arrTitles(10)%></b></td>
	</tr>  

	<tr><td height=5 colspan=2 nowrap></td></tr>       
	</table></td></tr>
	</form>
	<% End If%>
	</table></td></tr>	
	<tr>
	<td align="<%=align_var%>" bgcolor="#C9C9C9"  style="border: 1px solid #808080;border-top:none">
	<table cellpadding=1 cellspacing=0 border=0 width=100% align=center ID="Table5">	
	<tr><td height=5 colspan=2 nowrap></td></tr>   
 

  	
	<tr><td height=5 colspan=2 nowrap></td></tr>
	</table>
	</td></tr>
<%If trim(message_reciver_id) = trim(UserID) Then%>
<tr><td align=center colspan="2" height=5 nowrap></td></tr>
<tr><td align=center colspan="2">
<table cellpadding=0 cellspacing=0 width=80% border=0 ID="Table6">
<tr>
<td align=center width=140 nowrap>
<input type=button id="buttonUnread" name=buttonUnread value="לא קראתי" class="but_menu_red" onclick="javascript:updateStatus('<%=MessageId%>','1')"></td>

<td align=center width=140 nowrap>
<input type=button id=button_read name=button_read disabled value="קראתי" class="but_menu_Yellow" onclick="javascript:updateStatus('<%=MessageId%>','2')"></td>
<td align=center width=120 nowrap>
<A class="but_menu" style="width:100" onClick="return CheckFields(1)"><!--סגור-->שלח תגובה</a>
</td>
</tr>
<tr><td align=center colspan="2" height=5 nowrap></td></tr>
<%End If%>
</table>
</div>
</body>
</html>
<%set con=Nothing%>
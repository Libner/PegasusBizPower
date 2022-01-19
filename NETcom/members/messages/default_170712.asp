<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
 T = trim(Request.QueryString("T"))
 reciver_id = trim(Request("reciver_id"))
 sender_id = trim(Request("sender_id"))
 Message_status = trim(Request.QueryString("Message_status")) 
   delMessID = trim(Request("delMessID"))
  If delMessID<>nil And delMessID<>"" Then 		
  con.executeQuery "DELETE From messages WHERE message_id = " & delMessID	
	
    Response.Redirect "default.asp?T=" & T 
  end if

 If trim(T) = "OUT" Then 'לפתוח משימות יוצאות סגורות	
    If trim(sender_id) = "" then
		sender_id = UserID	
	End If
 Else ' לפתוח משימות נכנסות פתוחות
    T = "IN"	
    If trim(reciver_id) = "" Then	
		reciver_id = UserID		
	End If		
	If trim(Message_status) = "" Then
		'Message_status = "1,2"
		Message_status = "1"
		
	End If	
 End If
 
 where_sender = " AND user_id = " & sender_id
 where_reciver = " AND reciver_id = " & reciver_id  


	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 29 Order By word_id"				
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
	  set rstitle = Nothing	  	%> 
<!-- #include file="../../title_meta_inc.asp" -->
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript" src="../../tooltip.js"></script>
<script language="javascript" type="text/javascript">
<!--
	var oPopup = window.createPopup();
	function StatusDropDown(obj)
	{
		oPopup.document.body.innerHTML = Status_Popup.innerHTML; 
		oPopup.document.charset = "windows-1255";
		oPopup.show(-20, 17,65, 82, obj);    
	}
  function ViewMessage(messageID)
	{
			h = parseInt(520);
			w = parseInt(460);
			window.open("ViewMessage.asp?messageId=" + messageID, "T_Wind" ,"scrollbars=1,toolbar=0,top=50,left=100,width="+w+",height="+h+",align=center,resizable=0");
	}
	function closeMessage(messageID)
	{
			h = parseInt(520);
			w = parseInt(460);
			window.open("closeMessage.asp?messageId=" + messageID, "T_Wind" ,"scrollbars=1,toolbar=0,top=50,left=100,width="+w+",height="+h+",align=center,resizable=0");
	}
	function CheckDelMessage(T,MessID)
	{
     <%
		If trim(lang_id) = "1" Then
			str_confirm = "?האם ברצונך למחוק  את ההודעה"
		Else
			str_confirm = "Are you sure want to delete the message?"
		End If	
	 %>	
		if (confirm("<%=str_confirm%>")){     
			document.location.href = "default.asp?T=" + T + "&delMessID=" + MessID;
			return false;
		}else{
			return false;
		}
	}

	function Message_typeDropDown(obj)
	{
	    oPopup.document.body.innerHTML = Message_type_Popup.innerHTML;
	    oPopup.document.charset = "windows-1255"; 
	    oPopup.show(0, 17, obj.offsetWidth+50, 182, obj);    
	}

  function SendMessage()
	{
		h = parseInt(530);
		w = parseInt(460);
		window.open("addMessage.asp", "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}
	function DoCal(elTarget)
	{
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
    
    function addMessage(contactID,companyID,MessageID)
	{
		h = parseInt(530);
		w = parseInt(460);
		window.open("addMessage.asp?contactID=" + contactID + "&companyId=" + companyID + "&MessageID=" + MessageID, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}


//-->
</script>
</head>
<%
 if lang_id = "1" then
    arr_Status = Array("","חדש","נקרא","תגובה")	
    self_name = "עצמי"
 else
    arr_Status = Array("","new","read","replay")	
    self_name = "Self"
 end if
 
 where_status = " And (Message_status in (" & sFix(Message_status) & "))"
 status = Message_status

 If trim(T) = "OUT"  Then
	class_ = "4"
 Else
	class_ = "7"
 End if	  
 
 If Request("start_date") <> nil Then	
	start_date = Request("start_date")
	start_date_ = Month(start_date) & "/" & Day(start_date) & "/" & Year(start_date)   
 Else
	start_date_ = ""  
 End If

 If Request("end_date") <> nil Then
	end_date = Request("end_date")
    end_date_ = Month(end_date) & "/" & Day(end_date) & "/" & Year(end_date)
 Else
    end_date_ = ""
 End If
 
 if trim(Request.QueryString("page"))<>"" then
    page=Request.QueryString("page")
 else
    Page=1
 end if  
 
 if trim(Request.QueryString("numOfRow"))<>"" then
    numOfRow=Request.QueryString("numOfRow")
 else    
    numOfRow = 1
 end if  

 sort = Request.QueryString("sort")	
 if trim(sort)="" then  sort=0 end if  


  delMessageID = trim(Request("delMessageID"))
  If delMessageID<>nil And delMessageID<>"" Then 		
		sqlstr = "EXECUTE get_Message '" & UserID & "','" & OrgID& "','" & lang_ID & "','" & delMessageID & "'"
		set rs_Message = con.getRecordSet(sqlstr)
		if not rs_Message.eof then
				Message_content = trim(rs_Message("Message_content"))
				Message_date = trim(rs_Message("Message_date"))
				Message_types = trim(rs_Message("Message_types"))
				Message_status = trim(rs_Message("Message_status"))	
				activityId = trim(rs_Message("parent_id"))
				Message_sender_name = trim(rs_Message("sender_name"))
				Message_reciver_name = trim(rs_Message("reciver_name"))
				Message_replay = trim(rs_Message("Message_replay"))
				Message_close_date = trim(rs_Message("Message_close_date"))
				Message_sender_id = trim(rs_Message("User_ID"))
				Message_reciver_id = trim(rs_Message("reciver_id"))
				parentID = trim(rs_Message("parent_ID"))
				
				mail_recivers = ""
				sqlstr = "Select FirstName + Char(32) + LastName From Message_to_users Inner Join Users On Users.User_ID = Message_to_users.User_ID " &_
				"Where Message_ID = " & delMessageID
				set rs_names = con.getRecordSet(sqlstr)
				if not rs_names.eof then
					mail_recivers = rs_names.getString(,,",",",")
				end if
				set rs_names = nothing
				If Len(mail_recivers) > 1 Then
					mail_recivers = Left(mail_recivers,Len(mail_recivers)-1)
				End If							
			
				sqlstr = "Exec dbo.get_Message_types '"&delMessageID&"','"&OrgID&"'"
				set rs_Message_types = con.getRecordSet(sqlstr)
				If not rs_Message_types.eof Then
					Message_types_names = rs_Message_types.getString(,,",",",")
				Else
					Message_types_names = ""
				End If		
				
				If Len(Message_types_names) > 0 Then
					Message_types_names = Left(Message_types_names,(Len(Message_types_names)-1))
				End If						
				
			
						
			
			end if
			set rs_Message = Nothing	 				
		   
         
				
End If %>
<body>
<div id="ToolTip"></div>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_obj_var%>">
<tr><td width=100% align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width=100% align="<%=align_var%>">
  <%numOftab = 49%>
  <%If  trim(T) = "IN" Then%>
  <%numOfLink = 0%>
  <%ElseIf  trim(T) = "OUT" Then%>
  <%numOfLink = 1%>
  <%End If%> 
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title">&nbsp;</td></tr>		   
<tr><td width=100%>
<%  
 arr_bars = session("arr_bars")
 If IsArray(arr_bars) Then
		        For j=0 To Ubound(arr_bars)-1
				arr_bar = Split(arr_bars(j),",")
				barID = trim(arr_bar(0))
				barTitle = trim(arr_bar(1))
				barVisible = trim(arr_bar(2))
			'	response.Write barID &":"& barTitle &":"&barVisible&"<BR>"
				if barID="52" and barVisible="1" then
				 flagSend=1
				 end if
	next
			
				end if
				
If trim(Request("MessageTypeID")) <> nil Then
	MessageTypeID = trim(Request("MessageTypeID"))
Else 
	MessageTypeID = ""	
End If	  
		
dim sortby(12)	
If trim(T) = "OUT" Then 
sortby(0) = "Message_status, Message_date, Message_id DESC"
Else
sortby(0) = "Message_status, Message_date, Message_id DESC"
End If
sortby(1) = "company_name, Message_date DESC, Message_id DESC"
sortby(2) = "company_name DESC, Message_date DESC, Message_id DESC"
sortby(3) = "Message_date, Message_id DESC"
sortby(4) = "Message_date DESC, Message_id DESC"
sortby(5) = "contact_name, Message_date DESC, Message_id DESC"
sortby(6) = "contact_name DESC, Message_date DESC, Message_id DESC"
sortby(7) = "U.FIRSTNAME, U.LASTNAME, Message_date DESC,Message_id DESC"
sortby(8) = "U.FIRSTNAME DESC, U.LASTNAME DESC,Message_date DESC,Message_id DESC"
sortby(9) = "U1.FIRSTNAME,  U1.LASTNAME,Message_date DESC,Message_id DESC"
sortby(10) = "U1.FIRSTNAME DESC,  U1.LASTNAME DESC,Message_date DESC,Message_id DESC"
sortby(11) = "project_name,Message_date DESC,Message_id DESC"
sortby(12) = "project_name DESC,Message_date DESC,Message_id DESC"

urlSort="default.asp?search_company="& Server.URLEncode(search_company) &"&amp;search_contact="& Server.URLEncode(search_contact)&"&amp;search_project="&Server.URLEncode(search_project)&"&amp;start_date="&start_date&"&amp;end_date="&end_date & "&amp;Message_status="&Message_status & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id & "&amp;MessagetypeID=" & MessagetypeID & "&amp;T=" & T
UrlStatus="default.asp?search_company="& Server.URLEncode(search_company) &"&amp;search_contact="& Server.URLEncode(search_contact)&"&amp;search_project="&Server.URLEncode(search_project)&"&amp;start_date="&start_date&"&amp;end_date="&end_date & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id & "&amp;MessagetypeID=" & MessagetypeID & "&amp;T=" & T
urlType="default.asp?search_company="& Server.URLEncode(search_company) &"&amp;search_contact="& Server.URLEncode(search_contact) &"&amp;search_project="&Server.URLEncode(search_project)&"&amp;start_date="&start_date&"&amp;end_date="&end_date&"&amp;Message_status="&Message_status & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id  & "&amp;T=" & T
%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellpadding=0 cellspacing=0 dir="<%=dir_var%>">
   <tr>    
    <td width="100%" valign="top" align="center">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
    <tr>
    <td bgcolor=#FFFFFF align="left" width="100%" valign=top>
    <table width="100%" cellspacing="1" cellpadding="0" border=0>      
     <tr style="line-height:18px"> 	  
          <%If trim(sender_id) <> "" Then%>   
          <td align="center" class="title_sort" nowrap><!--מחק--><%=arrTitles(3)%></td> 
          <%End If%>
                <td width="145" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;תוכן תגובה&nbsp;</td>	                         
  
  	       <td width="145" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;תוכן הודעה&nbsp;</td>	          
	      <td id="td_type" width="150" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort"><!--סוג--><%=arrTitles(5)%>&nbsp;<IMG style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(24)%>" align=absmiddle onclick="return false" onmousedown="Message_typeDropDown(td_type)"></td>
	      <!--td width="80" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="11" OR trim(sort)="12" then%>_act<%end if%>"><%if trim(sort)="11" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="12" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=11"  title="<%=arrTitles(26)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="11" then%>bot<%elseif trim(sort)="12" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td-->
		          <td width="80" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="9" OR trim(sort)="10" then%>_act<%end if%>"><%if trim(sort)="9" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="10" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=9"  title="<%=arrTitles(26)%>"><%end if%>&nbsp;<!--אל--><%=arrTitles(6)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="9" then%>bot<%elseif trim(sort)="10" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
          <td width="80" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="7" OR trim(sort)="8" then%>_act<%end if%>"><%if trim(sort)="7" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="8" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=7"  title="<%=arrTitles(26)%>"><%end if%>&nbsp;<!--מאת--><%=arrTitles(7)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="7" then%>bot<%elseif trim(sort)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	      <td width="65" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="<%=arrTitles(25)%>"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="<%=arrTitles(26)%>"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=3"  title="<%=arrTitles(26)%>"><%end if%><!--תאריך יעד-->תאריך<img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	      <td width="40" id="td_status" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort"><!--'סט--><%=arrTitles(9)%>&nbsp;<IMG style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(24)%>" align=absmiddle onmousedown="StatusDropDown(td_status)"></td>
	     
    </tr>   
<%  
   PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
   If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
     	PageSize = 10
   End If	
   
   Set MessagesList = Server.CreateObject("ADODB.RECORDSET")
   sqlstr = "exec dbo.get_Messages_paging " & Page & "," & PageSize & ",'" & sFix(search_company) & "','" & sFix(search_contact) & "','" & sFix(search_project) & "','" & status & "','" & UserID & "','" & OrgID & "','" & lang_id & "','" & MessageTypeID & "','" & reciver_id & "','" & sender_id & "','" & sortby(sort) & "','" & start_date_ & "','" & end_date_ & "'"
  ' Response.Write sqlstr
   'Response.End
   set  MessagesList = con.getRecordSet(sqlstr)
   current_date = Day(date()) & "/" & Month(date()) & "/" & Year(date())
   dim	IS_DESTINATION 
   ids = ""
   If not MessagesList.EOF then		
        recCount = MessagesList("CountRecords")		
		       
   while not MessagesList.EOF       
       MessageId = trim(MessagesList(1))   :  ids = ids & MessageId 		
           Message_date = trim(MessagesList(6))
       status = trim(MessagesList(8))       
	   sender_name = trim(MessagesList(9))
	   reciver_name = trim(MessagesList(10))      
       Message_content = trim(MessagesList(11))          
       parentID = trim(MessagesList(12))  
       childID = trim(MessagesList(15))  
       Message_replay = trim(MessagesList(16))  
Replay_Flag=trim(MessagesList(19))  
       If Len(Message_content) > 25 Then
			Message_content_short = Left(Message_content,22) & "..."
	   Else
		   Message_content_short = Message_content 		
       End If
       
       If Len(Message_replay) > 25 Then
			Message_replay_short = Left(Message_replay,22) & "..."
	   Else
		   Message_replay_short = Message_replay 		
       End If       
     
       If IsDate(trim(Message_date)) Then
		'  Message_date = Day(trim(Message_date)) & "/" & Month(trim(Message_date)) & "/" & Right(Year(trim(Message_date)),2)
		  if DateDiff("d",Message_date,current_date) >= 0 then
			IS_DESTINATION = true
		  else
			IS_DESTINATION = false
		  end if
	   Else
		   Message_date=""
		   IS_DESTINATION = false
	   End If     				
	   Message_types_n = ""
	   sqlstr = "exec dbo.get_Message_types '"&MessageID&"','"&OrgID&"'"
	   set rs_Message_types = con.getRecordSet(sqlstr)
	   If not rs_Message_types.eof Then
		    Message_types_n = rs_Message_types.getString(,,"<br>","<br>")
	   Else
			Message_types_n = ""
	   End If
       
       If (trim(status) = "1" OR trim(status) = "2") And trim(UserID) = trim(sender_id) Then
			href = "href=""javascript:addMessage('" & contactID & "','" & companyID & "','" & MessageID & "')"""   
       Elseif trim(status) = "3" then
           	href = "href=""javascript:ViewMessage('" & MessageID & "')""" 
       elseif Replay_Flag="1" then
			href = "href=""javascript:closeMessage('" & MessageID & "')"""     
			else
				href = "href=""javascript:ViewMessage('" & MessageID & "')""" 
       End If      %>
        <tr>
        <%If trim(sender_id) <> "" Then%>  
          <td align="center" class="card<%=class_%>" valign="top"><input type="image" src="../../images/delete_icon.gif" onclick="return CheckDelMessage('<%=T%>','<%=MessageID%>')"></td>         
        <%End If%>  
       <%if message_replay<>"" then%>
            <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>" onMouseover="EnterContent('ToolTip','<%=arrTitles(4)%>','<%=Escape(breaks(message_replay))%>','<%=dir_obj_var%>'); Activate();" onMouseout="deActivate()">&nbsp;<%=breaks(trim(message_replay_short))%>&nbsp;</td>	      	      	      
	 <%else%>
	 <td class="card<%=class_%>"></td>
	 <%end if%>
 	       <td class="card<%=class_%>" valign="top" width=100% align="<%=align_var%>" dir="<%=dir_obj_var%>" onMouseover="EnterContent('ToolTip','<%=arrTitles(4)%>','<%=Escape(breaks(Message_content))%>','<%=dir_obj_var%>'); Activate();" onMouseout="deActivate()">&nbsp;<%=breaks(trim(Message_content_short))%>&nbsp;</td>	      	      	      
          <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>><%=Message_types_n%></a></td>	
		  <!--td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=projectName%>&nbsp;</a></td-->
           <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=reciver_name%>&nbsp;</a></td>
          <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%>>&nbsp;<%=sender_name%>&nbsp;</a></td>
          <td class="card<%=class_%>" valign="top" align="center" nowrap dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%> style="font-size: 8pt;" ><%If isDate(Message_date) Then%>&nbsp;<%=FormatDateTime(Message_date,2)%>&nbsp;<%=FormatDateTime(Message_date,4)%>&nbsp;&nbsp;<%End If%></a></td>         
	      <td class="card<%=class_%>" valign="top" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num<%=status%>" <%=href%>><%=arr_Status(status)%></a></td>	  
			
      </tr> 
	  <%	MessagesList.MoveNext
		if not MessagesList.eof then
		ids = ids & ","
		end if	  
	  Wend
	  
	  NumberOfPages = Fix((recCount / PageSize)+0.9)
	  if NumberOfPages > 1 then
	  urlSort = urlSort & "&amp;sort=" & sort 	  %>
	  <tr>
		<td width="100%" align="center" colspan="12" nowrap class="card<%=class_%>">
			<table border="0" cellspacing="0" cellpadding="2" dir="ltr" >
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center" align="<%=align_var%>"><A class="pageCounter" name=word28 title="<%=arrTitles(28)%>" href="<%=urlSort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  align="<%=align_var%>"><A class="pageCounter" href="<%=urlSort%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(29)%>" href="<%=urlSort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>
	<!--tr><td  class="title_form" align="center" colspan=9><%=recCount%> :ואצמנ תומושר כ"הס</td></tr-->										 
	<%End If%> 
	<tr>
	   <td colspan="12" height=18 class="card<%=class_%>" align="center" dir="ltr" style="color:#6E6DA6;font-weight:600"><!--נמצאו--><%=arrTitles(10)%>&nbsp;<%=recCount%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("MessagesMulti"))%> &nbsp;</td>
	</tr>
	<%Else%>
	<tr><td colspan="12" class="card<%=class_%>" align="center">&nbsp;&nbsp;</td></tr>
<% End If
   set MessagesList = Nothing%>
</table><input type="hidden" id="ids" value="<%=ids%>">
<input type="hidden" name="trapp" value="" ID="trapp">
</td>
<td width=80 nowrap valign="top" class="td_menu" style="border: 1px solid #808080; border-top: 0px">
<FORM action="default.asp?sort=<%=sort%>&amp;T=<%=T%>" method=POST id=form_search name=form_search target="_self">   
<table cellpadding="1" cellspacing="0" width="100%" >
<tr><td align="<%=align_var%>" colspan="2" height=20 class="title_search"><!--חיפוש--><%=arrTitles(11)%></td></tr>
<tr><td align="<%=align_var%>" colspan="2"><b><!--מתאריך--><%=arrTitles(12)%></b>&nbsp;</td></tr>
<tr>
	<td align="center" colspan=2 nowrap>
	<input dir="ltr" class="texts" type="text" id="start_date" name="start_date" value="<%=start_date%>" style="width:70" onclick="return DoCal(this);" readonly></td>					
</tr>
<tr><td width=100% align="<%=align_var%>" colspan=2>&nbsp;<b><!--עד תאריך--><%=arrTitles(13)%></b>&nbsp;</td></tr>
<TR>
	<td align="center" colspan=2 nowrap>
	<input dir="ltr" class="texts" type="text" id="end_date" name="end_date" value="<%=end_date%>" style="width:70" onclick="return DoCal(this);" readonly></td>		
</TR>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:80;" href="#" onclick="document.form_search.submit();">&nbsp;<!--עדכן תאריך--><%=arrTitles(14)%>&nbsp;</a></td></tr>

<%If trim(T) = "OUT" Then%>
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><b><!--אל--><%=arrTitles(15)%></b></td></tr>
<tr>
<td align="<%=align_var%>" colspan=2>
<select name="reciver_id" dir="<%=dir_obj_var%>" class="norm" style="width:100%" ID="reciver_id" onChange="form_search.submit();">
    <option value="" id=word19 name=word19><!-- כולם --><%=arrTitles(19)%></option>
    <%set UserList=con.GetRecordSet("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME FROM Users WHERE (ORGANIZATION_ID = " & OrgID & ") AND (ACTIVE = 1) ORDER BY FIRSTNAME + ' ' + LASTNAME")
    do while not UserList.EOF
    selUserID=UserList(0)
    selUserName=UserList(1)%>
    <option value="<%=selUserID%>" <%if trim(reciver_id)=trim(selUserID) then%> selected <%end if%>><%=selUserName%></option>
    <%
    UserList.MoveNext
    loop
    set UserList=Nothing%>
</select>
</td>
</tr>
<%End If%>
<%If trim(T) = "IN" Then%>
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><b><!--מאת--><%=arrTitles(16)%></b></td></tr>
<tr>
<td align="<%=align_var%>" colspan=2>
<select name="sender_id" dir="<%=dir_obj_var%>" class="norm" style="width:100%" ID="sender_id" onChange="form_search.submit();">
    <option value="" id=word17><!-- כולם --><%=arrTitles(17)%></option>
    <%set UserList=con.GetRecordSet("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME FROM Users WHERE (ORGANIZATION_ID = " & OrgID & ") AND (ACTIVE = 1) ORDER BY FIRSTNAME + ' ' + LASTNAME")
    do while not UserList.EOF
    selUserID=UserList(0)
    selUserName=UserList(1)%>
    <option value="<%=selUserID%>" <%if trim(sender_id)=trim(selUserID) then%> selected <%end if%>><%=selUserName%></option>
    <%
    UserList.MoveNext
    loop
    set UserList=Nothing%>
</select>
</td>
</tr>
<%End If%>
<tr><td colspan=2 height=7 nowrap></td></tr>
<tr><td colspan=2 height=1 bgcolor=#808080></td></tr>
<tr><td colspan=2 height=7 nowrap></td></tr>
<%If flagSend=1 Then%>
<tr><td align="center" nowrap colspan=2><a class="button_edit_1" style="width:78px;" href='javascript:void(0)' onclick="return SendMessage();">שלח הודעה</a></td></tr>
<%End If%>
<tr><td colspan=2 height=10 nowrap></td></tr>
</table>
</form>
</td></tr></table>
</td></tr></table>
</td></tr></table>
<DIV ID="Status_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; width:65; height:82; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;" >
<%For i=1 To uBound(arr_Status)	%>
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand; border-bottom:1px solid black"
	ONCLICK="parent.location.href='<%=UrlStatus%>&amp;Message_status=<%=i%>'">
    &nbsp;<%=arr_Status(i)%>&nbsp;
    </DIV>
<%Next%>    
    <DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand;"
	ONCLICK="parent.location.href='<%=UrlStatus%>&amp;Message_status=1,2,3'">&nbsp;<!--כל הרשימה--><%=arrTitles(20)%>&nbsp;</DIV>
</div>
</DIV>
<DIV ID="Message_type_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="overflow: scroll; position:absolute; top:0; left:0; width:100%; height:182; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#D3D3D3;SCROLLBAR-HIGHLIGHT-COLOR: #D3D3D3; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%  sqlstr = "Select messages_type_id, messages_type_name from messages_types WHERE ORGANIZATION_ID = " & OrgID & " Order By messages_type_id"
	set rsMessage_type = con.getRecordSet(sqlstr)
	while not rsMessage_type.eof %>
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; border-bottom:1px solid black; cursor:hand;"
	ONCLICK="parent.location.href='<%=urlType%>&amp;MessagetypeID=<%=rsMessage_type(0)%>&amp;Message_status=1,2,3'">&nbsp;<%=rsMessage_type(1)%>&nbsp;</DIV>
	<%
		rsMessage_type.moveNext
		Wend
		set rsMessage_type=Nothing
	%>
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand;"
	ONCLICK="parent.location.href='<%=urlType%>'">
    &nbsp;<!--כל הרשימה--><%=arrTitles(21)%>&nbsp;</DIV>
</div>
</DIV>
<script language="javascript" type="text/javascript">
<!--
document.body.onmousemove = overhere;
//-->
</script>
</body>
</html>
<%set con=Nothing%>
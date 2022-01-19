<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
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
	function Message_typeDropDown(obj)
	{
	    oPopup.document.body.innerHTML = Message_type_Popup.innerHTML;
	    oPopup.document.charset = "windows-1255"; 
	    oPopup.show(0, 17, obj.offsetWidth+50, 182, obj);    
	}
	
//-->
</script>
</head>
<%
	OrgID	 = trim(Request.Cookies("bizpegasus")("OrgID"))
  	lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))%>

<%

	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"
		align_var = "left"
		dir_obj_var = "ltr"
	Else
		dir_var = "ltr"
		align_var = "right"
		dir_obj_var = "rtl"
	End If		
	
	class_ = "7"
 
	if lang_id = "1" then
   arr_Status = Array("","חדש","נקרא","תגובה")	
	else
	   arr_Status = Array("","new","read","replay")	
	end if  
If trim(Request("MessageTypeID")) <> nil Then
	MessageTypeID = trim(Request("MessageTypeID"))
Else 
	MessageTypeID = ""	
End If
	   sort = Request.QueryString("sort")	
 if trim(sort)="" then  sort=0 end if  

 Message_status = trim(Request("Message_status")) 
 UserID = trim(Request.Form("User_ID"))  
 T = trim(Request("T"))
 'response.Write "T="& T
 If trim(T) = "OUT" Then 'לפתוח משימות יוצאות סגורות	
    sender_id = UserID		
    pageTitle="דו''ח הודעות שנשלחו"

	where_sender = " AND user_id = " & sender_id
 ElseIf trim(T) = "IN" Then ' לפתוח משימות נכנסות פתוחות	
     pageTitle="דו''ח הודעות שהתקבלו"
    reciver_id = UserID			
    where_reciver = " AND reciver_id = " & reciver_id
 End If
'response.Write  where_reciver &"<BR>"
 where_status = " And (Message_status in (" & sFix(Message_status) & "))"
'response.Write  where_status &"<BR>"
 If Request("date_start") <> nil Then	
	date_start = Request("date_start")
	start_date_ = Month(date_start) & "/" & Day(date_start) & "/" & Year(date_start)   
 Else
	start_date_ = ""  
 End If

 If Request("date_end") <> nil Then
	date_end = Request("date_end")
    end_date_ = Month(date_end) & "/" & Day(date_end) & "/" & Year(date_end)
 Else
    end_date_ = ""
 End If 
 

 If trim(UserID) <> "" Then
	sqlstr = "Select FirstName, LastName From Users WHERE User_ID = " & UserID
	set rs_user = con.getRecordSet(sqlstr)
	if not rs_user.eof then
		userName = "<br>&nbsp;עובד:&nbsp;<font color=""#666699"">" & trim(rs_user(0)) & " " & trim(rs_user(1))& "</font>"
	end if
	set rs_user = nothing
End If%>

<body>
<div id="ToolTip"></div>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_obj_var%>">
<tr><td width=100% align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width=100% align="<%=align_var%>">
  <%numOftab = 49%>
  <%numOfLink = 3
    UrlStatus="report_Messages.asp?date_start="&date_start&"&amp;date_end="&date_end & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id & "&amp;MessagetypeID=" & MessagetypeID & "&amp;T=" & T
    urlSort="report_Messages.asp?date_start="&date_start&"&amp;date_end="&date_end & "&amp;Message_status="&Message_status & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id & "&amp;MessagetypeID=" & MessagetypeID & "&amp;T=" & T
    urlType="report_Messages.asp?date_start="&date_start&"&amp;date_end="&date_end&"&amp;Message_status="&Message_status & "&amp;reciver_id=" & reciver_id & "&amp;sender_id=" & sender_id  & "&amp;T=" & T

  %>
 
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title">&nbsp;</td></tr>		   
<tr><td width=100%>
<FORM action="report_Messages.asp?sort=<%=sort%>&amp;T=<%=T%>" method=POST id=form_search name=form_search target="_self">   

<table border="0" width="100%" bgcolor="#FFFFFF" cellpadding=0 cellspacing=0 dir="<%=dir_var%>" ID="Table1">
   <tr>    
    <td width="100%" valign="top" align="center">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>" ID="Table2">
    <tr>
    <td bgcolor=#FFFFFF align="left" width="100%" valign=top>
    <table width="100%" cellspacing="1" cellpadding="0" border=0 ID="Table3">  
    	<tr><td colspan=8  align=center> <%=pageTitle%> לתקופה <%=date_start%>  - <%=date_end%>
    
     <tr style="line-height:18px"> 	  
       <td width="145" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;תוכן תגובה&nbsp;</td>	                         
  <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">תאריך תגובה</td>
  	       <td width="145" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;תוכן הודעה&nbsp;</td>	          
	      <td id="td_type" width="150" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort">סוג&nbsp;<IMG style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="סוג" align=absmiddle onclick="return false" onmousedown="Message_typeDropDown(td_type)"></td>
	      <td width="80" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="9" OR trim(sort)="10" then%>_act<%end if%>"><%if trim(sort)="9" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="למיון בסדר יורד"><%elseif trim(sort)="10" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=9"  title="למיון בסדר עולה"><%end if%>&nbsp;<!--אל-->אל&nbsp;<img src="../../images/arrow_<%if trim(sort)="9" then%>bot<%elseif trim(sort)="10" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
          <td width="80" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="7" OR trim(sort)="8" then%>_act<%end if%>"><%if trim(sort)="7" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="למיון בסדר יורד"><%elseif trim(sort)="8" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=7"  title="למיון בסדר עולה"><%end if%>&nbsp;<!--מאת-->מאת&nbsp;<img src="../../images/arrow_<%if trim(sort)="7" then%>bot<%elseif trim(sort)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	      <td width="65" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="למיון בסדר יורד"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=3"  title="למיון בסדר עולה"><%end if%><!--תאריך יעד-->תאריך<img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	      <td width="40" id="td_status" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort">סט&nbsp;<IMG style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="סט" align=absmiddle onmousedown="StatusDropDown(td_status)"></td>
	     
    </tr>   
	<%'response.end%>
	<%

	dim sortby(12)	
	If trim(T) = "OUT" Then 
	sortby(0) = "Message_date DESC,Message_status,Message_id DESC"
	Else
	sortby(0) = "Message_status,Message_date DESC,Message_id DESC"
	End If
sortby(1) = "Message_date DESC, Message_id DESC"
sortby(2) = "Message_date DESC, Message_id DESC"
sortby(3) = "Message_date, Message_id DESC"
sortby(4) = "Message_date DESC, Message_id DESC"
sortby(5) = "contact_name, Message_date DESC, Message_id DESC"
sortby(6) = "contact_name DESC, Message_date DESC, Message_id DESC"
sortby(7) = "USERS.FIRSTNAME, USERS.LASTNAME, Message_date DESC,Message_id DESC"
sortby(8) = "USERS.FIRSTNAME DESC, USERS.LASTNAME DESC,Message_date DESC,Message_id DESC"
sortby(9) = "USERS_1.FIRSTNAME,  USERS_1.LASTNAME,Message_date DESC,Message_id DESC"
sortby(10) = "USERS_1.FIRSTNAME DESC,  USERS_1.LASTNAME DESC,Message_date DESC,Message_id DESC"
sortby(11) = "Message_date DESC,Message_id DESC"
sortby(12) = "Message_date DESC,Message_id DESC"

'response.Write ("sortby(sort)="&sortby(7))
   Set MessagesList = Server.CreateObject("ADODB.RECORDSET")  
   sqlstr = "EXECUTE get_Messages '','','','" & Message_status & "','" & OrgID & "','" &MessageTypeID & "','" & reciver_id & "','" & sender_id & "','" & sortby(sort) & "','" & start_date_ & "','" & end_date_ & "','" & companyID & "'"

  ' Response.Write sqlstr
   'Response.End
   set  MessagesList = con.getRecordSet(sqlstr)
   current_date = Day(date()) & "/" & Month(date()) & "/" & Year(date())
   dim IS_DESTINATION 
   If not MessagesList.EOF then		
		arrMessages = MessagesList.getRows()
		MessagesList.Close()
		set MessagesList = Nothing	
i=0
	do while i <= uBound(arrMessages,2)            
      ' companyName =  trim(arrMessages(0,i))
       'Response.Write companyName
       'Response.End
      ' contactName =  trim(arrMessages(1,i))
         reciver_name = trim(arrMessages(3,i))
	   sender_name = trim(arrMessages(4,i))
       Message_types = trim(arrMessages(6,i))
       MessageId = trim(arrMessages(13,i))
       status = trim(arrMessages(17,i))
       Message_replay = trim(arrMessages(18,i))
       Message_close_date = trim(arrMessages(19,i))
       Message_content = trim(arrMessages(8,i))   
       Message_sender = trim(arrMessages(9,i))
       parentID = trim(arrMessages(15,i))         
       If IsDate(trim(arrMessages(21,i))) Then
			Message_open_date = Day(trim(arrMessages(21,i))) & "/" & Month(trim(arrMessages(21,i))) & "/" & Right(Year(trim(arrMessages(21,i))),2)
       Else
		    Message_open_date = ""
       End If
       If IsDate(trim(arrMessages(7,i))) Then
		'  Message_date=Day(trim(arrMessages(7,i))) & "/" & Month(trim(arrMessages(7,i))) & "/" & Right(Year(trim(arrMessages(7,i))),2)
		Message_date=trim(arrMessages(7,i))
		  if DateDiff("d",Message_date,current_date) >= 0 then
			IS_DESTINATION = true
		  else
			IS_DESTINATION = false
		  end if
	   Else
		   Message_date=""
		   IS_DESTINATION = false
	   End If     				
	  
	    If Len(Message_replay) > 25 Then
			Message_replay_short = Left(Message_replay,22) & "..."
	   Else
		   Message_replay_short = Message_replay 		
       End If       
		sqlstr = "Exec dbo.get_Message_types '"&MessageID&"','"&OrgID&"'"
		set rs_Message_types = con.getRecordSet(sqlstr)
		If not rs_Message_types.eof Then
			Message_types_names = rs_Message_types.getString(,,",",",")
		Else
			Message_types_names = ""
		End If		
		
		If Len(Message_types_names) > 0 Then
			Message_types_names = Left(Message_types_names,(Len(Message_types_names)-1))
		End If%>

     <tr>
    <%if message_replay<>"" then%>
            <td class="card<%=class_%>" valign="top" align="<%=align_var%>" dir="<%=dir_obj_var%>" onMouseover="EnterContent('ToolTip','תוכן','<%=Escape(breaks(message_replay))%>','<%=dir_obj_var%>'); Activate();" onMouseout="deActivate()">&nbsp;<%=breaks(trim(Message_replay_short))%>&nbsp;----</td>	      	      	      
	 <%else%>
	 <td class="card<%=class_%>"></td>
	 <%end if%> 
	  <td class="card<%=class_%>" align="<%=align_var%>" dir="<%=dir_obj_var%>"><%=Message_close_date%></td>
     <td class="card<%=class_%>" align="<%=align_var%>" dir="<%=dir_obj_var%>"><%=Message_content%></td>
     <td class="card<%=class_%>" align="<%=align_var%>" dir="<%=dir_obj_var%>"><%=Message_types_names%></td>
     <td class="card<%=class_%>" align="<%=align_var%>" dir="<%=dir_obj_var%>"><%=reciver_name%></td>
     <td class="card<%=class_%>" align="<%=align_var%>" dir="<%=dir_obj_var%>"><%= sender_name%></td>

     <td class="card<%=class_%>" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" <%=href%> style="font-size: 8pt;" ><%If isDate(Message_date) Then%>&nbsp;<%=FormatDateTime(Message_date,2)%>&nbsp;<%=FormatDateTime(Message_date,4)%>&nbsp;&nbsp;<%End If%></a></td>
  
     <td class="card<%=class_%>" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num<%=status%>" <%=href%>><%=arr_Status(status)%></a></td>
     </tr>
	<% ' response.Write "i="& i &"<BR>"
	  i=i+1	
	  loop
 End If%>
</form>
</TABLE>
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
	ONCLICK="parent.location.href='<%=UrlStatus%>&amp;Message_status=1,2,3'">&nbsp;כל הרשימה&nbsp;</DIV>
</div>
</DIV>
<DIV ID="Message_type_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="overflow: scroll; position:absolute; top:0; left:0; width:100%; height:182; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#D3D3D3;SCROLLBAR-HIGHLIGHT-COLOR: #D3D3D3; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%  
  sqlstr = "Select messages_type_id, messages_type_name from messages_types WHERE ORGANIZATION_ID = " & OrgID & " Order By messages_type_id"
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
    &nbsp;כל הרשימה&nbsp;</DIV>
</div>
</DIV>

<%set con=Nothing%>

</body>
</html>


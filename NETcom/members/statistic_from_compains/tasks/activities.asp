<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--

	var oPopup = window.createPopup();
	function StatusDropDown(obj)
	{
		oPopup.document.body.innerHTML = Status_Popup.innerHTML; 
		oPopup.document.charset = "windows-1255";
		oPopup.show(-65, 17, 80, 62, obj);    
	}
  
	function openactivity(contactID,companyID,activityID)
	{
			h = parseInt(570);
			w = parseInt(740);
			window.open("editactivity.asp?contactID=" + contactID + "&companyId=" + companyID + "&oldactivityId=" + activityID, "T_Wind" ,"scrollbars=1,toolbar=0,top=5,left=20,width="+w+",height="+h+",align=center,resizable=0");
	}
	
	function activity_typeDropDown(obj)
	{
	    oPopup.document.body.innerHTML = activity_type_Popup.innerHTML;
	    oPopup.document.charset = "windows-1255"; 
	    oPopup.show(-115, 17, 130, 182, obj);    
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
    
    function addactivity(contactID,companyID,activityID)
	{
		h = parseInt(400);
		w = parseInt(400);
		window.open("editactivity.asp?contactID=" + contactID + "&companyId=" + companyID + "&activityID=" + activityID, "T_Wind" ,"scrollbars=1,toolbar=0,top=100,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}
	
	function CheckDelActivity(activityID)
	{
		if (confirm("?האם ברצונך למחוק לצמיתות את דיווח הפעילות")){     
			document.location.href = "activities.asp?delActivityID=" + activityID;
			return false;
		}else{
			return false;
		}
    }

//-->
</SCRIPT>
</head>
<% 
  
 if trim(reciver_id) <> "" then
	where_reciver = " AND reciver_id = " & reciver_id
	where_sender = ""
	title = "נכנסות"
 elseif trim(sender_id) <> "" then
	where_sender = " AND user_id = " & sender_id
	where_reciver = ""	
	title = "יוצאות"		
 end if
 
 if trim(Request.QueryString("activity_status"))<>"" then
    activity_status=Request.QueryString("activity_status")    
 end if  
 
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

 if trim(Request.Form("search_company")) <> "" Or trim(Request.QueryString("search_company")) <> "" then
	 search_company = trim(Request.Form("search_company"))
	 if trim(Request.QueryString("search_company")) <> "" then
		search_company = trim(Request.QueryString("search_company"))
	 end if					
	 where_company = " And company_Name LIKE '%"& sFix(search_company) &"%'"
	 activity_status = "all"			
 else
	 where_company = ""		
 end if


if trim(Request.Form("search_contact")) <> "" Or trim(Request.QueryString("search_contact")) <> "" then
	 search_contact = trim(Request.Form("search_contact"))
	 if trim(Request.QueryString("search_contact")) <> "" then
		search_contact = trim(Request.QueryString("search_contact"))
	 end if					
	 where_contact = " And CONTACT_NAME LIKE '%"& sFix(search_contact) &"%'"
	 activity_status = "all"					
else
	 where_contact = ""		
end if

if trim(Request("search_project")) <> "" then		
	search_project = trim(Request("search_project"))	
	where_project = " And project_Name LIKE '%"& sFix(search_project) &"%'"			
	status = "all"	
else
	where_project = ""	
	search_project = ""		
end if

  delActivityID = trim(Request("delActivityID"))
  If delActivityID<>nil And delActivityID<>"" Then 			
	'con.executeQuery "delete From tasks WHERE task_id IN (Select task_id FROM activities Where activity_id=" &  delActivityID & ")"' delete contact also from activities	
	con.executeQuery "delete From activities Where activity_id=" &  delActivityID ' delete contact also from activities	
    Response.Redirect "activities.asp?sender_id=" & sender_id
  End If 

%>
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table1">
<tr><td width=100% align="right">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width=100% align="right">
<%numOftab = 0%> 
<%numOfLink = 6%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title">רשימת דיווחי פעילויות&nbsp;<%=title%></td></tr>		   
<tr><td width=100%>
<%
 If trim(activity_status) <>""  Then		
	if trim(activity_status) = "all" Then
		status = ""
		where_status = ""
	else	
		where_status = " And (activity_status in (" & sFix(activity_status) & "))"
		status = activity_status
	end If	
 Else where_status = " And (activity_status = '1') " 
 	   status = "1"	 	   
 End If
 
If trim(Request("activityTypeID")) <> nil Then
	 activityTypeID = trim(Request("activityTypeID"))
Else activityTypeID = ""	
End If	  
		
Select Case(status)
 Case "1" : no_search = "פתוחות"
 Case "2" : no_search = "סגורות" 
 Case "all" : no_search = ""
End Select

dim arr_Status(2)
arr_Status(1)="פתוחה"
arr_Status(2)="סגורה"

dim sortby(12)	
sortby(0) = "activity_status, activity_date"
sortby(1) = "rtrim(ltrim(company_name))"
sortby(2) = "rtrim(ltrim(company_name)) DESC"
sortby(3) = "activity_date"
sortby(4) = "activity_date DESC"
sortby(5) = "contact_name"
sortby(6) = "contact_name DESC"
sortby(7) = "sender_name"
sortby(8) = "sender_name DESC"
sortby(9) = "reciver_name"
sortby(10) = "reciver_name DESC"
sortby(11) = "project_name"
sortby(12) = "project_name DESC"

urlSort="activities.asp?search_company="& search_company &"&search_contact="& search_contact&"&search_project="&search_project & "&activity_status="&activity_status & "&reciver_id=" & reciver_id & "&sender_id=" & sender_id
UrlStatus="activities.asp?search_company="& search_company &"&search_contact="& search_contact&"&search_project="&search_project & "&reciver_id=" & reciver_id & "&sender_id=" & sender_id
urlType="activities.asp?search_company="& search_company &"&search_contact="& search_contact &"&search_project="&search_project& "&activity_status="&activity_status & "&reciver_id=" & reciver_id & "&sender_id=" & sender_id
%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellpadding=0 cellspacing=0 ID="Table2">
   <tr>    
    <td width="100%" valign="top" align="center">
    <table border="0" cellpadding="0" cellspacing="1" width="100%" ID="Table3">
    <tr>
    <td bgcolor=#FFFFFF align="left" width="100%" valign=top>
    <table width="100%" cellspacing="1" cellpadding="0" border=0 ID="Table4">      
     <tr style="line-height:18px"> 	           
          <td align=center class="title_sort" width=30 nowrap>מחק</td>            
  	      <td width=100% align="right" class="title_sort">תוכן&nbsp;</td>	          
	      <td width="90" align="right" nowrap class="title_sort">סוג דיווח&nbsp;<IMG id="find_stat" style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 ALT="בחר מרשימה" align=absmiddle onclick="return false" onmousedown="activity_typeDropDown(this)"></a></td>
	      <td width="80" align="right" nowrap class="title_sort<%if trim(sort)="11" OR trim(sort)="12" then%>_act<%end if%>"><%if trim(sort)="11" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="12" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=11" title="למיון בסדר עולה"><%end if%>פרויקט&nbsp;<img src="../../images/arrow_<%if trim(sort)="11" then%>bot<%elseif trim(sort)="12" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	      <td width="85" align="right" nowrap class="title_sort<%if trim(sort)="5" OR trim(sort)="6" then%>_act<%end if%>"><%if trim(sort)="5" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="6" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=5" title="למיון בסדר עולה"><%end if%>איש קשר&nbsp;<img src="../../images/arrow_<%if trim(sort)="5" then%>bot<%elseif trim(sort)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	      
          <td width="150" nowrap align="right" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=1" title="למיון בסדר עולה"><%end if%>לקוח&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	                  
          <td width="80" align="right" nowrap class="title_sort<%if trim(sort)="7" OR trim(sort)="8" then%>_act<%end if%>"><%if trim(sort)="7" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="8" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=7" title="למיון בסדר עולה"><%end if%>מדווח&nbsp;<img src="../../images/arrow_<%if trim(sort)="7" then%>bot<%elseif trim(sort)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>         
	      <td width="48" align="right" nowrap class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=3" title="למיון בסדר עולה"><%end if%>תאריך<img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	      <td width="43" align="right" nowrap class="title_sort">'סט&nbsp;<IMG id="Img1" style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 ALT="בחר מרשימה" align=absmiddle onmousedown="StatusDropDown(this)"></td>
	    
    </tr>   
<%  
   PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
   If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
     	PageSize = 10
   End If	
   
   Set activitysList = Server.CreateObject("ADODB.RECORDSET")  
   sqlstr = "EXECUTE get_activities '" & search_company & "','" & search_contact & "','" & search_project & "','" & status & "','" & OrgID & "','" & activityTypeID & "','" & reciver_id & "','" & sender_id & "','" & sortby(sort) & "'"
   'Response.Write sqlstr
   'Response.End
   set  activitysList =  con.getRecordSet(sqlstr)
   If not activitysList.EOF then		
		activitysList.PageSize=PageSize
		activitysList.AbsolutePage=Page
		recCount=activitysList.RecordCount 
		NumberOfPages=activitysList.PageCount
		i=0	
		arr_activities = activitysList.getRows(activitysList.PageSize)
		activitysList.Close()
		set activitysList = Nothing	
       
  do while i <= uBound(arr_activities,2)            
       companyName =  trim(arr_activities(0,i))
       'Response.Write companyName
       'Response.End
       contactName =  trim(arr_activities(1,i))
       projectName = trim(arr_activities(2,i))
       activity_sender_id = trim(arr_activities(8,i))       
       sender_name = trim(arr_activities(3,i))	  
       activity_types = trim(arr_activities(5,i))
       companyID =  trim(arr_activities(10,i))
       contactID = trim(arr_activities(11,i))
       activityId = trim(arr_activities(12,i))
       activityStatus = trim(arr_activities(15,i))
       activity_content = trim(arr_activities(7,i))   
       
       If IsDate(trim(arr_activities(6,i))) Then
		  activity_date=Day(trim(arr_activities(6,i))) & "/" & Month(trim(arr_activities(6,i))) & "/" & Right(Year(trim(arr_activities(6,i))),2)
	   Else
		   activity_date=""
	   End If     				
	  
	   activity_types_n = ""          
	   short_activity_types = ""
	   If Len(activity_types) > 0 Then				
			sqlstr="Select activity_type_name From activity_types Where activity_type_id IN (" & activity_types & ") Order By activity_type_id"
			'Response.Write sqlstr
			'Response.End
			set rssub = con.getRecordSet(sqlstr)				
			While not rssub.eof
				activity_types_n = activity_types_n & rssub(0) & "<BR>"
				rssub.moveNext
			Wend		    
			set rssub=Nothing
			If Len(activity_types_n) > 0 Then
				activity_types_n = Left(activity_types_n,(Len(activity_types_n)-4))
			End If				
		End If
       
       'If trim(sender_id) <> userID Then
       'href = "href=""javascript:openactivity('" & contactID & "','" & companyID & "','" & activityID & "')"""     
       'Else
       If trim(activityStatus) = "1" And  trim(activity_sender_id) = trim(userID) Then
       href = "href=""javascript:addactivity('" & contactID & "','" & companyID & "','" & activityID & "')"""   
       Else
       href = ""
       End If
       If trim(activity_sender_id) = trim(userID) Then
			myactivity = "4"
	   Else
		    myactivity = ""
	   End if	
      %>
        <tr height=18>             
          <td align=center class="card<%=myactivity%>" valign=top>
          <%If trim(activity_sender_id) = trim(userID) Then%> 
          <input type=image src="../../images/delete_icon.gif" border=0 hspace=0 vspace=0 onclick="return CheckDelActivity('<%=activityId%>')">
          <%End If%>
          </td>                  
	      <td class="card<%=myactivity%>" dir=rtl valign=top align="right"><a class="link_categ<%=myactivity%>" <%=href%>><%=trim(activity_content)%></a></td>	      	      	      
          <td class="card<%=myactivity%>" valign=top align="right"><a class="link_categ<%=myactivity%>" <%=href%>><%=activity_types_n%></a></td>	
		  <td class="card<%=myactivity%>" valign=top align="right"><a class="link_categ<%=myactivity%>" <%=href%>><%=projectName%></a></td>
          <td class="card<%=myactivity%>" valign=top align="right"><a class="link_categ<%=myactivity%>" <%=href%>><%=contactName%></a></td>
          <td class="card<%=myactivity%>" valign=top align="right"><a class="link_categ<%=myactivity%>" <%=href%>><%=companyName%></a></td>
          <td class="card<%=myactivity%>" valign=top align="right"><a class="link_categ<%=myactivity%>" <%=href%>><%=sender_name%></a></td>         
          <td class="card<%=myactivity%>" valign=top align="right"><a class="link_categ<%=myactivity%>" <%=href%>><%If isDate(activity_date) Then%><%=activity_date%><%End If%></a></td>         
	      <td class="card<%=myactivity%>" valign=top align="center" style="padding-top:2px"><a class="status_num<%=activityStatus%>" <%=href%>><%=arr_Status(activityStatus)%></a></td>	  
      </tr> 
	  <%
	  i=i+1	
	  loop
	  if NumberOfPages > 1 then
	  urlSort = urlSort & "&sort=" & sort
	  %>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap bgcolor="#E6E6E6">
			<table border="0" cellspacing="0" cellpadding="2" ID="Table5">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center" align="right"><A class=pageCounter title="לדפים הקודמים" href="<%=urlSort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle" align="right"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  align="right"><A class=pageCounter href="<%=urlSort%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td align="right"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center" align="right"><A class=pageCounter title="לדפים הבאים" href="<%=urlSort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>	
	<%End If%> 
	<tr>
	   <td colspan="10" height=18 bgcolor="#e6e6e6" align=center class="card" dir="ltr" style="color:#6E6DA6;font-weight:600">נמצאו <%=recCount%> דיווחי פעילויות <%=no_search%></td>
	</tr>
	<%Else%>
	<tr><td colspan=10 class="title_sort1" align=center>&nbsp;לא נמצאו דיווחי פעילויות&nbsp;<%=no_search%></td></tr>
<% End If%>
</table></td>
<td width=80 nowrap valign=top class="td_menu">
<table cellpadding=1 cellspacing=0 width=100% ID="Table6">
<tr><td align="right" colspan=2 class="title_search">חיפוש</td></tr>
<FORM action="activities.asp?sort=<%=sort%>&reciver_id=<%=reciver_id%>&sender_id=<%=sender_id%>" method=POST id=form_search name=form_search target="_self">   
<tr><td colspan=2 align="right" class="right_bar">שם חברה</td></tr>
<tr>
<td align=right><input type="image" onclick="form_search.submit();" src="../../images/search.gif" ID="Image1" NAME="Image1"></td>
<td align=right><input type="text" class="search" dir="rtl" style="width:70;" value="<%=search_company%>" name="search_company" ID="search_company"></td></tr>
<tr><td colspan=2 align="right" class="right_bar">שם איש קשר</td></tr>
<tr><td align=right><input type="image" onclick="form_search.submit();" src="../../images/search.gif" ID="Image2" NAME="Image2"></td>
<td align=right><input type="text" class="search" dir="rtl" style="width:70;" value="<%=search_contact%>" name="search_contact" ID="search_contact"></td></tr>
<tr><td colspan=2 align="right" class="right_bar">שם פרויקט</td></tr>
<tr><td align=right><input type="image" onclick="form_search.submit();" src="../../images/search.gif" ID="Image3" NAME="Image1"></td>
<td align=right><input type="text" class="search" dir="rtl" style="width:70;" value="<%=vFix(search_project)%>" name="search_project" ID="search_project"></td></tr>
</FORM>
<tr><td colspan=2 height=10 nowrap></td></tr>
<tr><td align="center" nowrap colspan=2><a class="button_edit_1" style="width:78;" href="#" onclick='window.open("addtask.asp", "T_Wind" ,"scrollbars=1,toolbar=0,top=100,left=120,width=400,height=400,align=center,resizable=0");'>הוסף משימה</a></td></tr>
<tr><td align="center" nowrap colspan=2><a class="button_edit_1" style="width:78;" href="#" onclick='window.open("addactivity.asp", "T_Wind" ,"scrollbars=1,toolbar=0,top=100,left=20,width=740,height=430,align=center,resizable=0");'>הוסף דיווח</a></td></tr>
<tr><td colspan=2 height=10 nowrap></td></tr>
</table></td></tr></table>
</td></tr></table>
</td></tr></table>
<DIV ID="Status_Popup" STYLE="display:none;">
<div dir="rtl" style="position:absolute; top:0; left:0; width:80; height:62; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;" >
<%For i=1 To uBound(arr_Status)	%>
	<DIV onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand; border-bottom:1px solid black"
	ONCLICK="parent.location.href='<%=UrlStatus%>&activity_status=<%=i%>'">
    &nbsp;<%=arr_Status(i)%>&nbsp;
    </DIV>
<%Next%>    
    <DIV onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand;"
	ONCLICK="parent.location.href='<%=UrlStatus%>&activity_status=all'">
    &nbsp;כל הסטטוסים&nbsp;
    </DIV>
</div>
</DIV>

<DIV ID="activity_type_Popup" STYLE="display:none;">
<div dir=rtl style="overflow: scroll; position:absolute; top:0; left:0; width:130; height:182; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#D3D3D3;SCROLLBAR-HIGHLIGHT-COLOR: #D3D3D3; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%  sqlstr = "Select activity_type_id, activity_type_name from activity_types WHERE ORGANIZATION_ID = " & OrgID & " Order By activity_type_id"
	set rsactivity_type = con.getRecordSet(sqlstr)
	while not rsactivity_type.eof %>
	<DIV dir=ltr onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; text-align:right; height:20px; background:#D3D3D3; border-bottom:1px solid black; cursor:hand;"
	ONCLICK="parent.location.href='<%=urlType%>&activitytypeID=<%=rsactivity_type(0)%>'">
    &nbsp;<%=rsactivity_type(1)%>&nbsp;
    </DIV>
	<%
		rsactivity_type.moveNext
		Wend
		set rsactivity_type=Nothing
	%>
	<DIV dir="rtl"  onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand;"
	ONCLICK="parent.location.href='<%=urlType%>'">
    &nbsp;כל הרשימה&nbsp;
    </DIV>
</div>
</DIV>
</body>
<%set con=Nothing%>
</html>

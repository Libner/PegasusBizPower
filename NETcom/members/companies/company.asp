<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<% 
  companyID = trim(Request("companyID"))
  if trim(lang_id) = "1" then
	   arr_status_comp = Array("","עתידי","פעיל","סגור","פונה")
  else
	   arr_status_comp = Array("","new","active","close","appeal")
  end if	   
  found_company = false
  If trim(companyID)<>"" then   
  set pr=con.GetRecordSet("EXEC dbo.companies_company_details @CompanyID="& companyID & ", @OrgID=" & OrgID)
  if not pr.EOF then	
	company_name   =pr("company_name")
  	company_name_E =pr("company_name_E")
	address	       =pr("address")
	address2	   =pr("address2")	
	street_number  =pr("street_number")
	apartment	   =pr("apartment")
	post_box 	   =pr("post_box")
	cityName	   =pr("city_Name")
	zip_code	   =pr("zip_code")	
	prefix_phone1  =pr("prefix_phone1")
	prefix_phone2  =pr("prefix_phone2")
	prefix_fax1	   =pr("prefix_fax1")
	prefix_fax2	   =pr("prefix_fax2")
	phone1	       =pr("phone1")
	phone2	       =pr("phone2")
	fax1           =pr("fax1")
	fax2	       =pr("fax2")
	company_email  =pr("email")
	url	           =pr("url")	
	date_update	   =pr("date_update")	
	date_start = trim(pr("date_start"))		   
    status_company =pr("status")
    private_flag   =pr("private_flag")
  	company_desc   = trim(pr("company_desc"))
    update_user_name  = trim(pr("FIRSTNAME") & " " & pr("LASTNAME"))

  	If Len(company_desc) > 43 Then
  		company_desc_short = Left(company_desc,40) & ".."
  	Else
  		company_desc_short = company_desc
  	End If      
    
    found_company = true 	
	Else
		found_company = false   
	End if  
 Else
	found_company = false   
 End if  
 
  delContactID=trim(Request("delContactID"))  
  If delContactID<>nil And delContactID<>"" Then 
  	'-------------------- deleting from tasks xml ---------------------------------------------
	sqlstr = "Select DISTINCT task_id From tasks WHERE contact_id = " & delContactID
	set rs_tasks = con.getRecordset(sqlstr)
	while not rs_tasks.eof 
	    taskId = rs_tasks(0)
		xmlFilePath = "../../../download/xml_tasks/bizpower_tasks_in.xml"			
		set objDOM = Server.CreateObject("Microsoft.XMLDOM")
		objDom.async = false			
		if objDOM.load(server.MapPath(xmlFilePath)) then
			set objNodes = objDOM.getElementsByTagName("TASK")
			for j=0 to objNodes.length-1
				set objTask = objNodes.item(j)
				node_task_id = objTask.attributes.getNamedItem("TASK_ID").text					
				if trim(taskId) = trim(node_task_id) Then					
					objDOM.documentElement.removeChild(objTask)
					exit for
				else
					set objTask = nothing
				end if
			next
			Set objNodes = nothing
			set objTask = nothing
			objDom.save server.MapPath(xmlFilePath)
		end if
		set objDOM = nothing	
	rs_tasks.moveNext
	Wend
	set rs_tasks = Nothing	

	sqlstr = "Select DISTINCT task_id From tasks WHERE contact_id = " & delContactID
	set rs_tasks = con.getRecordset(sqlstr)
	while not rs_tasks.eof 
	    taskId = rs_tasks(0)	
		xmlFilePath = "../../../download/xml_tasks/bizpower_tasks_out.xml"
		set objDOM = Server.CreateObject("Microsoft.XMLDOM")
		objDom.async = false			
		if objDOM.load(server.MapPath(xmlFilePath)) then
			set objNodes = objDOM.getElementsByTagName("TASK")
			for j=0 to objNodes.length-1
				set objTask = objNodes.item(j)
				node_task_id = objTask.attributes.getNamedItem("TASK_ID").text					
				if trim(taskId) = trim(node_task_id) Then					
					objDOM.documentElement.removeChild(objTask)
					exit for
				else
					set objTask = nothing
				end if
			next
			Set objNodes = nothing
			set objTask = nothing
			objDom.save server.MapPath(xmlFilePath)
		end if
		set objDOM = nothing
		rs_tasks.moveNext
	Wend
	set rs_tasks = Nothing	
	'-------------------- deleting from tasks xml ---------------------------------------------
	
	'-------------------- deleting from appeals xml -------------------------------------------
	sqlstr = "Select DISTINCT appeal_id From appeals WHERE contact_id = " & delContactID
	set rs_appeals = con.getRecordset(sqlstr)
	while not rs_appeals.eof 
	    appealId = rs_appeals(0)		
		xmlFilePath = "../../../download/xml_forms/bizpower_forms.xml"
		set objDOM = Server.CreateObject("Microsoft.XMLDOM")
		objDom.async = false			
		if objDOM.load(server.MapPath(xmlFilePath)) then
			set objNodes = objDOM.getElementsByTagName("FORM")
			for j=0 to objNodes.length-1
				set objappeal = objNodes.item(j)
				node_app_id = objappeal.attributes.getNamedItem("APPEAL_ID").text										
				if trim(appealId) = trim(node_app_id) Then					
					objDOM.documentElement.removeChild(objappeal)
					exit for
				else
					set objappeal = nothing
				end if
				
			next
			Set objNodes = nothing
			set objappeal = nothing
			objDom.save server.MapPath(xmlFilePath)
		end if
		set objDOM = nothing
		rs_appeals.moveNext
	Wend
	set rs_appeals = Nothing		
	'-------------------- deleting from appeals xml -------------------------------------------
   	con.ExecuteQuery "DELETE FROM Form_Value WHERE APPEAL_ID IN (Select appeal_id FROM appeals WHERE Contact_Id=" & delContactID & ")"
	con.ExecuteQuery "Delete From Appeals WHERE Contact_Id=" & delContactID	
	con.executeQuery "Delete From Tasks WHERE Contact_Id=" & delContactID
	con.executeQuery "Update meetings set Contact_Id = NULL WHERE Contact_Id=" & delContactID
	con.ExecuteQuery "Delete From Contacts WHERE Contact_Id=" & delContactID	
    Response.Redirect "company.asp?companyId=" & companyId
  End If      
  
  delprojectId=trim(Request("delprojectId"))  
  If delprojectId<>nil And delprojectId<>"" Then 
	con.ExecuteQuery "DELETE FROM FORM_PROJECT_VALUE WHERE PROJECT_ID IN (Select PROJECT_ID FROM PROJECTS WHERE PROJECT_ID=" & delprojectId & ")"
	con.ExecuteQuery "delete from HOURS where PROJECT_ID=" & delprojectId
	con.ExecuteQuery "delete from PROJECTS where PROJECT_ID=" & delprojectId
	con.ExecuteQuery "Delete From pricing_to_jobs where pricing_id=" & delprojectId	
	con.ExecuteQuery "Delete From pricing_to_projects where project_id=" & delprojectId
	Response.Redirect "company.asp?companyId=" & companyId               
  End If 
  
  delDocumentID=trim(Request("delDocumentID"))
  if delDocumentID <> nil then   
		set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
'-----deleting files----
		sqlstr = "select * from DOCUMENTS where document_id = "& delDocumentID
		'Response.Write sqlstr
		'Response.End
		set files = con.getRecordSet(sqlstr)
		do while not files.eof
			file_path="../../../download/documents/" & trim(files("document_file"))
			if fs.FileExists(server.mappath(file_path)) then
				set a = fs.GetFile(server.mappath(file_path))
				a.delete '------------------------------
			else
				Response.Write(file_path)	
			end if	
		files.movenext
		loop
		set files =nothing
		set fs = nothing
		
		con.executeQuery("delete from DOCUMENTS where document_id = " & delDocumentID)
		con.executeQuery("delete from company_documents where document_id = " & delDocumentID)
		Response.Redirect "company.asp?companyId=" & companyId
  End if 
  
	'עדכן סטטוס פגישות שעברו
	sqlstr = "exec dbo.UpdateMeetingStatus '" & OrgID & "'"
    'Response.Write sqlstr
    con.ExecuteQuery(sqlstr)    

	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 7 Order By word_id"				
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
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
var DetailsList,ContactList
var changed = false; 

var oPopup = window.createPopup();
function appealDropDown(obj)
{
    oPopup.document.body.innerHTML = appeal_Popup.innerHTML; 
    oPopup.document.charset="windows-1255";    
    oPopup.show(0, 20, obj.offsetWidth, 68, obj);    
}

function taskDropDown(obj)
{
    oPopup.document.body.innerHTML = task_Popup.innerHTML; 
    oPopup.document.charset="windows-1255";
    oPopup.show(0, 17, obj.offsetWidth, 68, obj);    
}

function descDropDown(obj)
{
  document.all["div_comp_desc"].style.zIndex=11; 
  document.all["div_comp_desc"].style.display='inline';
  document.all["div_comp_desc"].style.visibility="visible";
  return false; 
}

function closeDesc()
{
   document.all["div_comp_desc"].style.display='none';
   document.all["div_comp_desc"].style.visibility="hidden"; 
   return false;
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

function CheckDelContact(companyId,contactId) {
  <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את ה" & trim(Request.Cookies("bizpegasus")("ContactsOne")) & " עם כל המשימות והטפסים המשויכים לו"
     Else
		str_confirm = "Are you sure want to delete the " & trim(Request.Cookies("bizpegasus")("ContactsOne")) & " with all his tasks and forms ?"
     End If   
  %>
  if (window.confirm("<%=str_confirm%>"))
  {         
     document.location.href = "company.asp?companyId="+companyId+"&delContactID="+contactId;
    
  }
  return false;  
}

function CheckDelDocument(companyId,documentId) {
  <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את המסמך"
     Else
		str_confirm = "Are you sure want to delete the document?"
     End If   
  %>
  if (confirm("<%=str_confirm%>"))
  {         
     document.location.href = "company.asp?companyId="+companyId+"&delDocumentID="+documentId;
    
  }
  return false;  
}

function CheckDelProject(projectId,companyId) {
  <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את ה" & trim(Request.Cookies("bizpegasus")("Projectone"))
     Else
		str_confirm = "Are you sure want to delete the " & trim(Request.Cookies("bizpegasus")("Projectone")) & " ?"
     End If   
  %>
  if (confirm("<%=str_confirm%>"))
  {         
     document.location.href = "company.asp?delprojectId="+projectId+"&companyId="+companyId;
    
  }
  return false;  
}

function openactivity(contactID,companyID,activityID)
{
	h = parseInt(330);
	w = parseInt(740);
	window.open("../tasks/addactivity.asp?contactID=" + contactID + "&companyId=" + companyID + "&activityID=" + activityID, "T_Wind" ,"scrollbars=1,toolbar=0,top=100,left=20,width="+w+",height="+h+",align=center,resizable=0");		
}

function addtask(contactID,companyID,taskID)
{
	h = parseInt(530);
	w = parseInt(470);
	window.open("../tasks/addtask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskID=" + taskID, "AddTask" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
}

function closeTask(contactID,companyID,taskID)
{
		h = parseInt(480);
		w = parseInt(470);
		window.open("../tasks/closetask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskId=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=50,left=100,width="+w+",height="+h+",align=center,resizable=0");
}

function openTask(contactID,companyID,taskID)
{
	h = parseInt(450);
	w = parseInt(740);
	window.open("../tasks/edittask.asp?contactID=" + contactID + "&companyId=" + companyID + "&oldtaskId=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=70,left=20,width="+w+",height="+h+",align=center,resizable=0");
}
	function addmeeting(meetingID)
	{
		h = parseInt(500);
		w = parseInt(465);
		window.open("../meetings/addmeeting.asp?meetingID=" + meetingID + "&meeting_date=<%=Date()%>&participant_id=<%=UserID%>&companyId=<%=companyID%>", "AddMeeting" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}
	function closemeeting(meetingID)
	{
		h = parseInt(490);
		w = parseInt(465);
		window.open("../meetings/addmeeting.asp?meetingID=" + meetingID + "", "CloseMeeting" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}
function openPreview(pageId)
{
	result = window.open("../Pages/result.asp?pageId="+pageId,"Result","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=780, height=540, left=2, top=2");
	return false;
}			
//-->
</script>  
</head>
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF dir="<%=dir_var%>">
<tr><td width="100%" colspan=2>
<!-- #include file="../../logo_top.asp" -->
</td></tr>
<%numOftab = 0%>
<%numOfLink = 1%>
<tr><td width="100%" colspan=2>
<!--#include file="../../top_in.asp"-->
</td></tr>
<%If found_company Then%>
<tr>
<td width="100%" class="page_title" dir="<%=dir_obj_var%>"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<font color="#6F6DA6"><%=company_name%></font>&nbsp;</td></tr>         
<tr><td width="100%" valign=top>
<table cellpadding=0 cellspacing=1 width=100% bgcolor=white>
<tr><td width="100%" valign=top>
<table cellpadding=0 bgcolor="#808080"  dir="ltr" cellspacing=1 width=100% border=0 style="border-collpase:collapse;">
<FORM name="form_company" ACTION="addcompany.asp?companyID=<%=companyID%>" METHOD="post" ID="form_company">   
<tr>          
  <td align="<%=align_var%>" width="50%" bgcolor=#E6E6E6 nowrap valign="top">  
  <input type="hidden" name="companyID" id="companyID" value="<%=companyID%>">  
  <table cellpadding=0 cellspacing=1 width=100% align=center dir="<%=dir_var%>">
     <tr><td colspan=2 class="title_form" align="<%=align_var%>"><!--פרטים--><%=arrTitles(3)%>&nbsp;</td></tr>
     <tr><td colspan=2 height=5></td></tr>
     <tr>
         <td width="100%" align="<%=align_var%>"><span dir="<%=dir_obj_var%>" style="width:250" class="Form_R"><%=company_name%></span></td>
         <td width="100" nowrap align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--שם--><%=arrTitles(4)%>&nbsp;</td> 
      </tr>            
      <tr>
        <td align="<%=align_var%>"><span dir="<%=dir_obj_var%>" style="width:250" class="Form_R"><%=address%></span></td>
        <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--כתובת--><%=arrTitles(5)%>&nbsp;</td>
     </tr>             
     <tr>
         <td align="<%=align_var%>" nowrap><span dir="<%=dir_obj_var%>" class="Form_R" style="width:130"><%=cityName%></span></td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--עיר--><%=arrTitles(6)%>&nbsp;</td>
      </tr>                        
      <tr>
         <td align="<%=align_var%>" nowrap><span dir="ltr" class="Form_R" style="width:130"><%=zip_code%></span></td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--מיקוד--><%=arrTitles(59)%>&nbsp;</td>
      </tr>
      <tr>
          <td align="<%=align_var%>" dir=ltr><span dir="ltr" class="Form_R" style="width:130"><%=phone1%></span></td>
          <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--1 -טלפון--><%=arrTitles(7)%>&nbsp;</td>
      </tr>                                
      <tr>  
          <td align="<%=align_var%>" dir=ltr><span dir="ltr" class="Form_R" style="width:130"><%=phone2%></span></td>
          <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--2 טלפון--><%=arrTitles(8)%>&nbsp;</td>
      </tr>  
      <tr>  
          <td align="<%=align_var%>" dir=ltr><span dir="ltr" class="Form_R" style="width:130"><%=fax1%></span></td>
          <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<!--פקס--><%=arrTitles(9)%>&nbsp;</td>
      </tr>                          
      <tr>
         <td align="<%=align_var%>">
         <%If Len(company_email) > 0 Then%>
         <a href="mailto:<%=company_email%>" dir="<%=dir_obj_var%>" style="width:250;direction:ltr;text-align:left" class="Form_R hand">
         <%=trim(company_email)%></a>
		 <%Else%>
         <span style="width:250" class="Form_R"></span>         
         <%End If%>
         </td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;Email&nbsp;</td>
      </tr>       
       <tr>
         <td align="<%=align_var%>">
         <%If Len(url) > 0 Then%>
         <a href="<%=url%>" target=_blank dir="<%=dir_obj_var%>" style="width:250;direction:ltr;text-align:left"  class="Form_R hand">         
         <%=trim(url)%></a>
         <%Else%>
         <span style="width:250" class="Form_R"></span>
         <%End If%>
         </td>
         <td align="<%=align_var%>" style="padding-right:15px">&nbsp;<span id="word10" name=word10><!--אתר--><%=arrTitles(10)%></span>&nbsp;</td>
      </tr>       
      <tr>
      <td align="<%=align_var%>" valign=bottom>
      <%If Len(trim(company_desc)) > 0 Then%><input type=image src="../../images/popup.gif" onclick="return descDropDown(this)" name=word57 title="<%=arrTitles(57)%>" border=0 hspace=0 vspace=0><%End if%>
      <span class="Form_R" dir="<%=dir_obj_var%>" style="width:250;line-height:16px"><%=trim(company_desc_short)%></span>      
      </td>
      <td align="<%=align_var%>" style="padding-right:15px" valign=top>&nbsp;<span id="word12" name=word12><!--פרטים נוספים--><%=arrTitles(12)%></span>&nbsp;</td>
      </tr>     
       <tr>
           <td align="<%=align_var%>">
           <span class="Form_R" dir="<%=dir_var%>" style="width:250;"><%=date_start%></span>
           </td>
           <td align="<%=align_var%>" style="padding-right:15px" valign=top>&nbsp;תאריך הצטרפות&nbsp;</td>
      </tr>       
       <tr>
           <td align="<%=align_var%>">
           <span class="Form_R" dir="<%=dir_var%>" style="width:250;"><%=date_update%></span>
           </td>
           <td align="<%=align_var%>" style="padding-right:15px" valign=top>&nbsp;תאריך עדכון&nbsp;</td>
      </tr> 
       <tr>
           <td align="<%=align_var%>">
           <span class="Form_R" dir="<%=dir_var%>" style="width:250;"><%=update_user_name%></span>
           </td>
           <td align="<%=align_var%>" style="padding-right:15px" valign=top>&nbsp;עודכן ע"י&nbsp;</td>
      </tr>        
      <tr>
           <td align="<%=align_var%>"><span style="width:60;text-align:center" class="status_num<%=status_company%>">
           <%=arr_status_comp(status_company)%></span>
           </td>
           <td align="<%=align_var%>" style="padding-right:15px" valign=top>&nbsp;<!--סטטוס--><%=arrTitles(13)%>&nbsp;</td>
      </tr>
      <tr><td colspan="2" height="5" nowrap></td></tr>             
    </FORM>                        
</table>
</td>
<td width=50% align=center  bgcolor=#E6E6E6 nowrap valign="top">
<table cellpadding=0 cellspacing=1 width=100% align=center bgcolor=white dir="<%=dir_var%>">
<%if trim(companyID)<>"" then
 urlSort="company.asp?companyID="& companyID 
 dim sortby(7)	
 sortby(1) = "CONTACT_NAME"
 sortby(2) = "CONTACT_NAME DESC"
 sortby(3) = "date_update"
 sortby(4) = "date_update DESC"

 sort = trim(Request.QueryString("sort"))
 If trim(sort) = "" Then
	sort = 1
 End If

'PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
 If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
	PageSize = 4
 End If	

 if trim(Request.QueryString("pageCont"))<>"" then
    pageCont=Request.QueryString("pageCont")
 else
    pageCont=1
 end if  
 
 if trim(Request.QueryString("numOfRowCont"))<>"" then
    numOfRowCont=Request.QueryString("numOfRowCont")
 else
    numOfRowCont = 1
 end if  

  sql="SELECT contact_ID,email,prefix_phone,phone,prefix_cellular,prefix_fax,fax,"&_
  " cellular,CONTACT_NAME,messanger_Name,date_update FROM Contacts "&_
  " WHERE company_ID=" & companyID &_
  " ORDER BY "& sortby(sort)
 set listContact=con.GetRecordSet(sql)
 if not listContact.EOF then
 %>
<tr><td colspan=5 class="title_form" align="<%=align_var%>"><%=trim(Request.Cookies("bizpegasus")("ContactsMulti"))%>&nbsp;</td></tr>
<%
	listContact.PageSize=PageSize
	listContact.AbsolutePage=pageCont
	recCount=listContact.RecordCount 
	NumberOfPages=listContact.PageCount
	i=1	
%>
<tr>     
     <td width="35"   nowrap align="center" class="title_sort"><span id=word14 name=word14><!--מחק--><%=arrTitles(14)%></span></td>        
     <td width="125"  nowrap class="title_sort" align="<%=align_var%>"><span id="word15" name=word15><!--תפקיד--><%=arrTitles(15)%></span>&nbsp;</td>     
     <td width="100%" align="<%=align_var%>" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=1" name=word52 title="<%=arrTitles(52)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
</tr> 
<% do while (not listContact.EOF and i<=listContact.PageSize)
      contact_ID=trim(listContact("contact_ID"))
      CONTACT_NAME=trim(listContact("CONTACT_NAME"))      
      email=trim(listContact("email"))
      prefix_phone=trim(listContact("prefix_phone"))
      phone=trim(listContact("phone"))
      prefix_cellular=trim(listContact("prefix_cellular"))
      cellular=trim(listContact("cellular"))
      prefix_fax=trim(listContact("prefix_fax"))
      fax=trim(listContact("fax"))
      date_update=trim(listContact("date_update"))
      messangerName=trim(listContact("messanger_Name"))           
%>      
       <tr>   
         <td class="card" align="center"><input name=word55 type=image src="../../images/delete_icon.gif" title="<%=arrTitles(55)%>" style="border: none" hspace=0 vspace=0 onclick="return CheckDelContact('<%=companyId%>','<%=contact_ID%>')"></td>         
         <td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a href="contact.asp?companyId=<%=companyId%>&contactID=<%=contact_ID%>" class="link_categ" target=_self>&nbsp;<%=messangerName%>&nbsp;</a></td>         
         <td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a href="contact.asp?companyId=<%=companyId%>&contactID=<%=contact_ID%>" class="link_categ" target=_self>&nbsp;<%=CONTACT_NAME%>&nbsp;</a></td>
      </tr>
<%  listContact.MoveNext
	i=i+1
	loop
	if NumberOfPages > 1 then
	urlSort = urlSort & "&sort=" & sort
%>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2" dir=ltr>               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>
	         
	         <tr>
	         <%if numOfRowCont <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word53 title="<%=arrTitles(53)%>" href="<%=urlSort%>&pageCont=<%=10*(numOfRowCont-1)-9%>&amp;numOfRowCont=<%=numOfRowCont-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowCont-1)) <= NumberOfPages Then
	                  if CInt(pageCont)=CInt(i+10*(numOfRowCont-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRowCont-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&pageCont=<%=i+10*(numOfRowCont-1)%>&amp;numOfRowCont=<%=numOfRowCont%>" ><%=i+10*(numOfRowCont-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRowCont) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word54 title="<%=arrTitles(54)%>" href="<%=urlSort%>&pageCont=<%=10*(numOfRowCont) + 1%>&amp;numOfRowCont=<%=numOfRowCont+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>		
	<%listContact.close 
	set listContact=Nothing
  End if 
%>
<% End If%>
<% End If%>
<%if trim(companyID)<>"" then
 urlSort="company.asp?companyID="& companyID 
 dim sortby_pr(7)	
 sortby_pr(1) = "project_name"
 sortby_pr(2) = "project_name DESC"
 sortby_pr(3) = "project_code"
 sortby_pr(4) = "project_code DESC"

 sort = trim(Request.QueryString("sort"))
 If trim(sort) = "" Then
	sort = 1
 End If

'PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
 If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
	PageSize = 5
 End If	

 if trim(Request.QueryString("pageProj"))<>"" then
    pageProj=Request.QueryString("pageProj")
 else
    pageProj=1
 end if  
 
 if trim(Request.QueryString("numOfRowProj"))<>"" then
    numOfRowProj=Request.QueryString("numOfRowProj")
 else
    numOfRowProj = 1
 end if  

  sql="SELECT projects.project_id, projects.project_code, projects.project_name FROM projects "&_
  " WHERE company_ID = " & companyID & " AND projects.ORGANIZATION_ID = " & OrgID & " ORDER BY "& sortby_pr(sort)
  
 set list_pr=con.GetRecordSet(sql)
 if not list_pr.EOF then
 %>
<tr><td colspan=5 class="title_form" align="<%=align_var%>"><%=trim(Request.Cookies("bizpegasus")("ProjectMulti"))%>&nbsp;</td></tr>
<%
	list_pr.PageSize=PageSize
	list_pr.AbsolutePage=pageProj
	recCount=list_pr.RecordCount 
	NumberOfPages=list_pr.PageCount
	i=1	
%>
<tr>     
     <td width="35"   nowrap align="center" class="title_sort"><span id="word16" name=word16><!--מחק--><%=arrTitles(16)%></span></td>        
     <td width="125"  nowrap class="title_sort" align="<%=align_var%>"><span id="word17" name=word17><!--קוד--><%=arrTitles(17)%></span>&nbsp;</td>     
     <td width="100%" align="<%=align_var%>" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=1" name=word52 title="<%=arrTitles(52)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
</tr> 
<% do while (not list_pr.EOF and i<=list_pr.PageSize)
      PROJECT_ID=trim(list_pr("PROJECT_ID"))
      project_code=trim(list_pr("project_code"))      
      project_name=trim(list_pr("project_name"))                  
%>      
       <tr>   
         <td class="card" align="center"><input type=image name=word55 src="../../images/delete_icon.gif" title="<%=arrTitles(55)%>" style="border: none" hspace=0 vspace=0 onclick="return CheckDelProject('<%=PROJECT_ID%>','<%=companyId%>')" ID="Image2" NAME="Image2"></td>         
         <td class="card" align="<%=align_var%>"><a href="../projects/project.asp?projectID=<%=PROJECT_ID%>&companyID=<%=companyID%>" class="link_categ" target=_self>&nbsp;<%=project_code%>&nbsp;</a></td>         
         <td class="card" align="<%=align_var%>"><a href="../projects/project.asp?projectID=<%=PROJECT_ID%>&companyID=<%=companyID%>" class="link_categ" target=_self>&nbsp;<%=project_name%>&nbsp;</a></td>
      </tr>
<%  list_pr.MoveNext
	i=i+1
	loop
	if NumberOfPages > 1 then
	urlSort = urlSort & "&sort=" & sort
%>
	  <tr>
		<td width="100%" align=middle colspan="8" nowrap  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2" dir=ltr>               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>
	         
	         <tr>
	         <%if numOfRowProj <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word53 title="<%=arrTitles(53)%>" href="<%=urlSort%>&pageProj=<%=10*(numOfRowProj-1)-9%>&amp;numOfRowProj=<%=numOfRowProj-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowProj-1)) <= NumberOfPages Then
	                  if CInt(pageProj)=CInt(i+10*(numOfRowProj-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRowProj-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&pageProj=<%=i+10*(numOfRowProj-1)%>&amp;numOfRowProj=<%=numOfRowProj%>" ><%=i+10*(numOfRowProj-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRowProj) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word54 title="<%=arrTitles(54)%>" href="<%=urlSort%>&pageProj=<%=10*(numOfRowProj) + 1%>&amp;numOfRowProj=<%=numOfRowProj+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>		
	<%list_pr.close 
	set list_pr=Nothing
  End if   
End If%>
<tr><td colspan=4 class="title_form" align="<%=align_var%>"><%=trim(Request.Cookies("bizpegasus")("ActivitiesMulti"))	%>&nbsp;</td></tr>
<%
dim sortby_pr_cl(7)	
 sortby_pr_cl(1) = "project_name"
 sortby_pr_cl(2) = "project_name DESC"
 sortby_pr_cl(3) = "project_code"
 sortby_pr_cl(4) = "project_code DESC"

 sort_pr_cl = trim(Request.QueryString("sort_pr_cl"))
 If trim(sort_pr_cl) = "" Then
	sort_pr_cl = 1
 End If

'PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
 If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
	PageSize = 5
 End If	

 if trim(Request.QueryString("pageProjCl"))<>"" then
    pageProjCl=Request.QueryString("pageProjCl")
 else
    pageProjCl=1
 end if  
 
 if trim(Request.QueryString("numOfRowProjCl"))<>"" then
    numOfRowProjCl=Request.QueryString("numOfRowProjCl")
 else
    numOfRowProjCl = 1
 end if  

  sql="SELECT projects.project_id, projects.project_code, projects.project_name FROM projects "&_
  " WHERE company_ID = 0 AND projects.ORGANIZATION_ID = " & OrgID & " ORDER BY "& sortby_pr_cl(sort_pr_cl)
  
 set list_pr=con.GetRecordSet(sql)
 if not list_pr.EOF then
	list_pr.PageSize=PageSize
	list_pr.AbsolutePage=pageProjCl
	recCount=list_pr.RecordCount 
	NumberOfPages=list_pr.PageCount
	i=1	
%>
<tr>     
     <td colspan=2 nowrap class="title_sort" align="<%=align_var%>"><span id=word18 name=word18><!--קוד--><%=arrTitles(18)%></span>&nbsp;</td>     
     <td width="100%" align="<%=align_var%>" class="title_sort<%if trim(sort_pr_cl)="1" OR trim(sort_pr_cl)="2" then%>_act<%end if%>"><%if trim(sort_pr_cl)="1" then%><a class="title_sort" href="<%=urlSort%>&sort_pr_cl=<%=sort_pr_cl+1%>" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_pr_cl)="2" then%><a class="title_sort" href="<%=urlSort%>&sort_pr_cl=<%=sort_pr_cl-1%>" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_pr_cl=1" name=word52 title="<%=arrTitles(52)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("ActivitiesOne"))	%>&nbsp;<img src="../../images/arrow_<%if trim(sort_pr_cl)="1" then%>bot<%elseif trim(sort_pr_cl)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
</tr> 
<% do while (not list_pr.EOF and i<=list_pr.PageSize)
      PROJECT_ID=trim(list_pr("PROJECT_ID"))
      project_code=trim(list_pr("project_code"))      
      project_name=trim(list_pr("project_name"))                  
%>      
       <tr>            
         <td colspan=2 class="card" align="<%=align_var%>"><a href="../projects/project.asp?projectID=<%=PROJECT_ID%>&companyID=<%=companyID%>" class="link_categ" target=_self>&nbsp;<%=project_code%>&nbsp;</a></td>         
         <td class="card" align="<%=align_var%>"><a href="../projects/project.asp?projectID=<%=PROJECT_ID%>&companyID=<%=companyID%>" class="link_categ" target=_self>&nbsp;<%=project_name%>&nbsp;</a></td>
      </tr>
<%  list_pr.MoveNext
	i=i+1
	loop
	if NumberOfPages > 1 then
	urlSort = urlSort & "&sort=" & sort
%>
	  <tr>
		<td width="100%" align=middle colspan="8" nowrap  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2" dir=ltr>               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>
	         
	         <tr>
	         <%if numOfRowProjCl <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word53 title="<%=arrTitles(53)%>" href="<%=urlSort%>&pageProjCl=<%=10*(numOfRowProjCl-1)-9%>&amp;numOfRowProjCl=<%=numOfRowProjCl-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowProjCl-1)) <= NumberOfPages Then
	                  if CInt(pageProjCl)=CInt(i+10*(numOfRowProjCl-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRowProjCl-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&pageProjCl=<%=i+10*(numOfRowProjCl-1)%>&amp;numOfRowProjCl=<%=numOfRowProjCl%>" ><%=i+10*(numOfRowProjCl-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRowProjCl) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word54 title="<%=arrTitles(54)%>" href="<%=urlSort%>&pageProjCl=<%=10*(numOfRowProjCl) + 1%>&amp;numOfRowProjCl=<%=numOfRowProjCl+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>		
	<%list_pr.close 
	set list_pr=Nothing
  End if  
   End If%>
<% End If%>
</table>   
</td></tr>
<%
   If trim(companyID) <> "" Then
        if lang_id = "1" then
        arr_Status = Array("","חדש","בטיפול","סגור")	
        else
        arr_Status = Array("","new","active","close")	
        end if
        
        dim sortby_task(12)			
		sortby_task(1) = "CONTACT_NAME"
		sortby_task(2) = "CONTACT_NAME DESC "
		sortby_task(3) = "task_date"
		sortby_task(4) = "task_date DESC"
		sortby_task(5) = "task_status,task_date DESC"
		sortby_task(6) = "task_status DESC,task_date DESC"
		sortby_task(7) = "U.FIRSTNAME, U.LASTNAME"
		sortby_task(8) = "U.FIRSTNAME DESC, U.LASTNAME DESC"
		sortby_task(9) = "U1.FIRSTNAME,  U1.LASTNAME"
		sortby_task(10) = "U1.FIRSTNAME DESC,  U1.LASTNAME DESC"
		sortby_task(11) = "project_name"
		sortby_task(12) = "project_name DESC "		
      
        sort_task = trim(Request.QueryString("sort_task"))
        If Len(sort_task) = 0 Then
			sort_task = 5
		Else sort_task = trim(Request.QueryString("sort_task"))	
        End If      
		
		If trim(Request("taskTypeID")) <> nil Then
			  taskTypeID = trim(Request("taskTypeID"))
			  search = true 
		Else  taskTypeID = ""	
		      search = false
		End If	           		
		
		task_status = trim(Request("task_status")) ' הצגת כל סטטוסי המשימות 		
		title_tasks = trim(Request.Cookies("bizpegasus")("TasksMulti"))   
    
		if trim(Request.QueryString("page_task"))<>"" then
			page_task=Request.QueryString("page_task")
		else
			page_task=1
		end if  
	 
		if trim(Request.QueryString("row_task"))<>"" then
			row_task=Request.QueryString("row_task")
		else
			row_task = 1
		end if  
		PageSize = 5

		sqlstr = "EXECUTE get_tasks_paging " & page_task & "," & PageSize & ",'','','','" & task_status & "','" & UserID & "','" & OrgID & "','" & lang_id & "','" & taskTypeID & "','','','" & sortby_task(sort_task) & "','','','" & companyID & "'"
		'Response.Write sqlstr
		'Response.End   
		set tasksList = con.getRecordSet(sqlstr)	
		If not tasksList.eof  Then
			recCount = tasksList("CountRecords")
	    End If		
  	    If not tasksList.eof Or search = true Then    
%>
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6">
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width=100%>
  <tr><td width=100%><A name="table_tasks"></A>  
  <table cellpadding=0 cellspacing=0 width=100%>
  <tr>  
  <td class="title_form" width=100%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%>&nbsp;<font color="#E6E6E6">(<%=company_name%>)</font>&nbsp;</td>
  </tr>
  </table></td></tr>  	  
  <tr>
  <td width="100%" valign=top>
   <table width=100% dir="<%=dir_var%>" border=0 bordercolor=red cellpadding=0 cellspacing=1 bgcolor="#FFFFFF" ID="Table9">   	    
    <tr> 
      <td align=center class="title_sort" width=29 nowrap>&nbsp;</td>     
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort" width=220 nowrap><span id=word19 name=word19><!--תוכן--><%=arrTitles(19)%></span>&nbsp;</td>                  
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=110 nowrap class="title_sort" id=td_types name=td_types>&nbsp;<span id="word20" name=word20><!--סוגים--><%=arrTitles(20)%></span>&nbsp;<IMG name=word56 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(56)%>" align=absmiddle onmousedown="taskDropDown(td_types)"></td>      
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=90 nowrap class="title_sort<%if trim(sort_task)="1" OR trim(sort_task)="2" then%>_act<%end if%>"><%if trim(sort_task)="1" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_task)="2" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=1#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%end if%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="1" then%>bot<%elseif trim(sort_task)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=90 nowrap valign=top class="title_sort<%if trim(sort_task)="9" OR trim(sort_task)="10" then%>_act<%end if%>"><%if trim(sort_task)="9" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_task)="10" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=9#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%end if%>&nbsp;<span id="word21" name=word21><!--אל--><%=arrTitles(21)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="9" then%>bot<%elseif trim(sort_task)="10" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>          
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=90 nowrap valign=top class="title_sort<%if trim(sort_task)="7" OR trim(sort_task)="8" then%>_act<%end if%>"><%if trim(sort_task)="7" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_task)="8" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=7#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%end if%>&nbsp;<span id="word22" name=word22><!--מאת--><%=arrTitles(22)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="7" then%>bot<%elseif trim(sort_task)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>                
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=78 nowrap valign=top class="title_sort<%if trim(sort_task)="3" OR trim(sort_task)="4" then%>_act<%end if%>"><%if trim(sort_task)="3" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_task)="4" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=3#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%end if%><span id="word23" name=word23><!--תאריך יעד--><%=arrTitles(23)%></span><img src="../../images/arrow_<%if trim(sort_task)="3" then%>bot<%elseif trim(sort_task)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>          
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=43 nowrap valign=top class="title_sort<%if trim(sort_task)="5" OR trim(sort_task)="6" then%>_act<%end if%>"><%if trim(sort_task)="5" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_task)="6" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=5#table_tasks" name=word52 title="<%=arrTitles(52)%>"><%end if%><span id="word24" name=word24><!--'סט--><%=arrTitles(24)%></span><img src="../../images/arrow_<%if trim(sort_task)="5" then%>bot<%elseif trim(sort_task)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>
    </tr>   
 <%	current_date = Day(date()) & "/" & Month(date()) & "/" & Year(date())
	dim	 IS_DESTINATION    
    while not tasksList.EOF       
		taskId = trim(tasksList(1))			
		contact_Name = trim(tasksList(5))
		task_date = trim(tasksList(6))
		project_Name = trim(tasksList(7))  
		task_status = trim(tasksList(8))
		sender_name = trim(tasksList(9))
		reciver_name = trim(tasksList(10))      
		task_content = trim(tasksList(11))          
		parentID = trim(tasksList(12))  
		ReciverID = trim(tasksList(13))
		SenderID = trim(tasksList(14))
	    childID = trim(tasksList(15)) 	
		
		task_types_names=""	 
		sqlstr = "Exec dbo.get_task_types '"&taskID&"','"&OrgID&"'"   
		set rs_task_types = con.getRecordSet(sqlstr)
		If not rs_task_types.eof Then
			task_types_names = rs_task_types.getString(,,",",",")
		Else
			task_types_names = ""
		End If		
		
		If Len(task_types_names) > 0 Then
			task_types_names = Left(task_types_names,(Len(task_types_names)-1))
		End If
		
		If cInt(strScreenWidth) > 800 Then
			numOfLetters = 150
		Else
			numOfLetters = 55
		End If
		tel_text = trim(tasksList("task_content"))
		If Len(tel_text) > numOfLetters Then
			tel_text_short = Left(tel_text , numOfLetters-2) & ".."
		Else tel_text_short = tel_text	
		End If
		task_date = trim(tasksList("task_date"))
		If isDate(task_date) Then
			d_s = Day(task_date) & "/" & Month(task_date) & "/" & Right(Year(task_date),2)
			if DateDiff("d",d_s,current_date) >= 0 then
				IS_DESTINATION = true
			else
				IS_DESTINATION = false
			end if
		else
			d_s = ""
			IS_DESTINATION = false  
		End If
		If trim(UserID) = trim(SenderID) Then
			class_ = "4"
		ElseIf trim(UserID) = trim(ReciverID) Then
			class_ = "7"
		Else
			class_ = ""	
	    End if	
	    If trim(UserID) = trim(SenderID) AND trim(task_status) = "1" Then
			href = "href=""javascript:addtask('" & contactID & "','" & companyID & "','" & taskID & "')"""   
        Else      
			href = "href=""javascript:closeTask('" & contactID & "','" & companyID & "','" & taskID & "')"""     
        End If
      %>      
      <tr>  
		<td align=center class="card<%=class_%>" valign=middle>
		<%If trim(taskID) <> "" And trim(childID) <> "" Then%>
		<input type=image src="../../images/hets4.gif" border=0 hspace=0 vspace=0 onclick='window.open("../tasks/task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' ID="Image1" NAME="Image1">
		<%End If%>
		<%If trim(parentID) <> "" And trim(taskID) <> "" Then%>
		<input type=image src="../../images/hets4a.gif" border=0 hspace=0 vspace=0 onclick='window.open("../tasks/task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' ID="Image4" NAME="Image4">
		<%End If%>
		</td>              
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>" title="<%=vFix(tel_text)%>">&nbsp;<%=tel_text_short%>&nbsp;</a></td>
      <td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=task_types_names%>&nbsp;</a></td>
      <td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=CONTACT_NAME%>&nbsp;</a></td>
      <td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=reciver_name%>&nbsp;</a></td>                 
      <td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=sender_name%>&nbsp;</a></td>                         
      <td align="<%=align_var%>" class="card<%=class_%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> <%if IS_DESTINATION and task_status <> 3 then%> name=word58 title="<%=arrTitles(58)%>"><span style="width:9px;COLOR: #FFFFFF;BACKGROUND-COLOR: #FF0000;text-align:center"><b>!</b></span><%else%>><%end if%>&nbsp;<%=d_s%>&nbsp;</a></td>            
      <td align=center class="card<%=class_%>" valign=top><a class="task_status_num<%=task_status%>" style="line-height:120%;"><%=arr_Status(task_status)%></A></td>
      </tr>
<%
    tasksList.MoveNext
	Wend
	  
	NumOfPagesTasks = Fix((recCount / PageSize)+0.9)
    urlSort = urlSort & "&sort_task=" & sort_task
	if NumOfPagesTasks > 1 then	
%>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2" dir=ltr>               
	        <% If NumOfPagesTasks > 10 Then 
	              num = 10 : row_task = cInt(NumOfPagesTasks / 10)
	           else num = NumOfPagesTasks : row_task = 1    	                      
	           End If
            %>
	         
	         <tr>
	         <%if row_task <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word53 title="<%=arrTitles(53)%>" href="<%=urlSort%>&page_task=<%=10*(row_task-1)-9%>&amp;row_task=<%=row_task-1%>#table_tasks" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(row_task-1)) <= NumOfPagesTasks Then
	                  if CInt(page_task)=CInt(i+10*(row_task-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(row_task-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&page_task=<%=i+10*(row_task-1)%>&amp;row_task=<%=row_task%>#table_tasks" ><%=i+10*(row_task-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumOfPagesTasks > cint(num * row_task) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word54 title="<%=arrTitles(54)%>" href="<%=urlSort%>&page_task=<%=10*(row_task) + 1%>&amp;row_task=<%=row_task+1%>#table_tasks" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>		
	
	<%
	End if 
%>
 </table></td></tr>
	<%If tasksList.recordCount = 0 Then%>
	<tr><td align=center class=card1>&nbsp;</td></tr>									
	<%End If%>	 
 </table></td></tr>  
<% 
	End if 
    set tasksList = Nothing
    End If

   If trim(companyID) <> "" Then       
        
        dim sortby_doc(2)			
        sortby_doc(0) = "document_id desc"
		sortby_doc(1) = "document_name"
		sortby_doc(2) = "document_name DESC "		
      
        sort_doc = trim(Request.QueryString("sort_doc"))
        If Len(sort_doc) = 0 Then
			sort_doc = 0
		Else sort_doc = trim(Request.QueryString("sort_doc"))	
        End If    
    
		if trim(Request.QueryString("page_doc"))<>"" then
			page_doc=Request.QueryString("page_doc")
		else
			page_doc=1
		end if  
	 
		if trim(Request.QueryString("row_doc"))<>"" then
			row_doc=Request.QueryString("row_doc")
		else
			row_doc = 1
		end if  
		PageSize = 5
		sqlstr = "Select document_id,document_name,document_desc,document_file " &_
		" from company_documents_view Where company_ID = " & companyID		
		sqlstr=sqlstr& " Order By " & sortby_doc(sort_doc)
		'Response.Write sqlstr
		'Response.End      
		set rsdoc = con.getRecordSet(sqlstr)
	
		If not rsdoc.eof  Then
			rsdoc.PageSize=PageSize
			rsdoc.AbsolutePage=page_doc
			records_doc=rsdoc.RecordCount 
			pages_doc=rsdoc.PageCount
			i=1		
	   End If		
  	   If not rsdoc.eof Then    
%>
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6">
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width=100%>
  <tr><td width=100%><A name="table_doc"></A>  
  <table cellpadding=0 cellspacing=0 width=100%>
  <tr>  
  <td class="title_form" width=100%>&nbsp;<span id=word25 name=word25><!--מסמכים מצורפים--><%=arrTitles(25)%></span>&nbsp;<font color="#E6E6E6">(<%=company_name%>)</font>&nbsp;</td>
  </tr>
  </table></td></tr>  	  
  <tr>
  <td width="100%" valign=top>
   <table width=100% dir="<%=dir_var%>" border=0 bordercolor=red cellpadding=0 cellspacing=1 bgcolor="#FFFFFF">   	    
    <tr> 
      <td align=center class="title_sort" width=40 nowrap><span id="word26" name=word26><!--מחיקה--><%=arrTitles(26)%></span></td>     
      <td align="<%=align_var%>" class="title_sort" width=100%>&nbsp;<span id="word27" name=word27><!--תיאור--><%=arrTitles(27)%></span>&nbsp;</td>                  
      <td align="<%=align_var%>" width=193 nowrap class="title_sort<%if trim(sort_doc)="1" OR trim(sort_doc)="2" then%>_act<%end if%>"><%if trim(sort_doc)="1" then%><a class="title_sort" href="<%=urlSort%>&sort_doc=<%=sort_doc+1%>#table_doc" name=word51 title="<%=arrTitles(51)%>"><%elseif trim(sort_doc)="2" then%><a class="title_sort" href="<%=urlSort%>&sort_doc=<%=sort_doc-1%>#table_doc" name=word52 title="<%=arrTitles(52)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_doc=1#table_doc" name=word52 title="<%=arrTitles(52)%>"><%end if%>&nbsp;<span id=word50 name=word50><!--כותרת--><%=arrTitles(50)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_doc)="1" then%>bot<%elseif trim(sort_doc)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>
    </tr>   
 <%
      task_types_name = ""
      do while (not rsdoc.EOF and i<=rsdoc.PageSize)       
		documentID = trim(rsdoc("document_ID"))			
		document_name = trim(rsdoc("document_name"))
	    document_desc = trim(rsdoc("document_desc"))
	    document_file = trim(rsdoc("document_file"))
	    href = " href = ""../../../download/documents/" & document_file & """"
		If cInt(strScreenWidth) > 800 Then
			numOfLetters = 150
		Else
			numOfLetters = 55
		End If
		
		If Len(document_desc) > numOfLetters Then
			document_desc_short = Left(document_desc , numOfLetters-2) & ".."
		Else document_desc_short = document_desc	
		End If
      %>      
      <tr>  
		<td align=center class="card" valign=middle><input type=image name=word55 src="../../images/delete_icon.gif" border=0 vspace=0 hspace=0 title="<%=arrTitles(55)%>" onclick="return CheckDelDocument('<%=companyId%>','<%=documentID%>')"></td>              
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card" valign=top><a class="link_categ" <%=href%> target=_blank dir="<%=dir_obj_var%>" title="<%=vFix(document_desc)%>">&nbsp;<%=document_desc_short%>&nbsp;</a></td>   
        <td align="<%=align_var%>" class="card" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> target=_blank dir="<%=dir_obj_var%>">&nbsp;<%=document_name%>&nbsp;</a></td>                         
     </tr>
      <%
      rsdoc.moveNext
       i=i+1
	loop
    urlSort = urlSort & "&sort_doc=" & sort_doc
	if pages_doc > 1 then	
%>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2">               
	        <% If pages_doc > 10 Then 
	              num = 10 : row_doc = cInt(pages_doc / 10)
	           else num = pages_doc : row_doc = 1    	                      
	           End If
            %>
	         
	         <tr>
	         <%if row_doc <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word53 title="<%=arrTitles(53)%>" href="<%=urlSort%>&page_doc=<%=10*(row_doc-1)-9%>&amp;row_doc=<%=row_doc-1%>#table_doc" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(row_doc-1)) <= pages_doc Then
	                  if CInt(page_doc)=CInt(i+10*(row_doc-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(row_doc-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&page_doc=<%=i+10*(row_doc-1)%>&amp;row_doc=<%=row_doc%>#table_doc" ><%=i+10*(row_doc-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if pages_doc > cint(num * row_doc) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word54 title="<%=arrTitles(54)%>" href="<%=urlSort%>&page_doc=<%=10*(row_doc) + 1%>&amp;row_doc=<%=row_doc+1%>#table_doc" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>		
	
	<%
	End if 
%>
 </table></td></tr>	
 </table></td></tr>  
<% 
	End if 
    set rsdoc = Nothing
    End If    
    If trim(SURVEYS)  = "1" Then
	dim sortby_app(16)	
	sortby_app(1) = "appeal_date"
	sortby_app(2) = "appeal_date DESC"
	sortby_app(3) = "appeal_id"
	sortby_app(4) = "appeal_id DESC"
	sortby_app(5) = "User_Name, appeal_id DESC"
	sortby_app(6) = "User_Name DESC, appeal_id DESC"
	sortby_app(7) = "product_name, appeal_id DESC"
	sortby_app(8) = "product_name DESC, appeal_id DESC"
	sortby_app(9) = "CONTACT_NAME, appeal_id DESC"
	sortby_app(10) = "CONTACT_NAME DESC, appeal_id DESC"
	sortby_app(11) = "appeal_status, appeal_id DESC"
	sortby_app(12) = "appeal_status DESC, appeal_id DESC"
	sortby_app(13) = "Company_NAME, appeal_id DESC"
	sortby_app(14) = "Company_NAME DESC, appeal_id DESC"

	sort_app = Request("sort_app")	
	if sort_app = nil then
		sort_app = 2
	end if
	
	search = false
	If trim(Request("productID")) <> nil Then
		productID = trim(Request("productID"))
		where_product = " AND QUESTIONS_ID = " & productID
		search = true
	Else 
	    productID = ""	
	    where_product = ""
	    search = false
	End If
	
	arr_Status = session("arr_Status")

	sqlstr = "Exec dbo.get_appeals '','','','','" & OrgID & "','" & sortby_app(sort_app) & "','','','" & companyID & "','','','','" & productID & "','" & UserID & "','','','" & is_groups & "'"
    'Response.Write sqlStr
	set app=con.GetRecordSet(sqlStr)
	app_count = app.RecordCount
	if Request("page_app")<>"" then
		page_app=request("page_app")
	else
		page_app=1
	end if
	if not app.eof then
		app.PageSize = 5
		app.AbsolutePage=page_app
		recCount=app.RecordCount 		
		NumberOfPagesApp = app.PageCount
		i=1
		j=0
		ids = "" 'list of appeal_id
	end if
	if not app.eof Or search = true then %>
<input type="hidden" name="trapp" value="" ID="trapp">		
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6">
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width=100%>  
  <tr><td width=100%><A name="table_appeals"></A>  
  <table cellpadding=0 cellspacing=0 width=100% ID="Table6">
  <tr>  
  <td class="title_form" width=100%>&nbsp;<!--טפסים מצורפים--><%=arrTitles(28)%>&nbsp;<font color="#E6E6E6">(<%=company_name%>)</font>&nbsp;</td>
  </tr>
  </table></td></tr>  	  	
  <tr>
	<td width="100%" align="center" valign=top dir="<%=dir_var%>">
	<table width="100%" BORDER="0" CELLSPACING="1" CELLPADDING="1" bgcolor="#FFFFFF" dir="<%=dir_var%>">	
	<tr>	   
	    <td nowrap dir="<%=dir_obj_var%>" class="title_sort" align=center><span id="word29" name=word29><!--הדפס--><%=arrTitles(29)%></span></td>	  
		<td nowrap dir="<%=dir_obj_var%>" class="title_sort" align=center><%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%></td>
		<td nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=13 or sort_app=14 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=14 then%>13<%elseif sort_app=13 then%>14<%else%>14<%end if%>#table_appeals" target="_self">&nbsp;<%=trim(Request.Cookies("bizpegasus")("Contactsone"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="13" then%>bot<%elseif trim(sort_app)="14" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>		
		<%key_table_width=400%>
		<td class="title_sort">&nbsp;</td>
		<td width="100%" dir="<%=dir_obj_var%>" id=td_app_prod name=td_app_prod align="<%=align_var%>" class="title_sort">&nbsp;<!--סוג טופס--><%=arrTitles(31)%>&nbsp;<IMG name=word56 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(56)%>" align="absmiddle" onmousedown="appealDropDown(td_app_prod)"></td>				
		<td dir="<%=dir_obj_var%>" nowrap align=center class="title_sort<%if sort_app=3 or sort_app=4 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=4 then%>3<%elseif sort_app=3 then%>4<%else%>4<%end if%>#table_appeals" target="_self">&nbsp;ID&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="3" then%>bot<%elseif trim(sort_app)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td dir="<%=dir_obj_var%>" nowrap align=center class="title_sort<%if sort_app=1 or sort_app=2 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=2 then%>1<%elseif sort_app=1 then%>2<%else%>2<%end if%>#table_appeals" target="_self">&nbsp;<!--תאריך--><%=arrTitles(32)%>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="1" then%>bot<%elseif trim(sort_app)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>			
		<td dir="<%=dir_obj_var%>" nowrap align=center class="title_sort<%if sort_app=11 or sort_app=12 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=12 then%>11<%elseif sort_app=11 then%>12<%else%>12<%end if%>#table_appeals" target="_self">&nbsp;<!--'סט--><%=arrTitles(33)%>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="11" then%>bot<%elseif trim(sort_app)="12" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	</tr>
<%do while (not app.eof and j<app.PageSize)
		appid=app("appeal_id")			
		COMPANY_NAME = app("COMPANY_NAME")
		If trim(COMPANY_NAME) = "" Or IsNull(COMPANY_NAME) Then
			COMPANY_NAME = ""
		End If	
		CONTACT_NAME = app("CONTACT_NAME")
		If trim(CONTACT_NAME) = "" Or IsNull(CONTACT_NAME) Then
			CONTACT_NAME = ""
		End If	
		PROJECT_NAME = app("PROJECT_NAME")
		If trim(PROJECT_NAME) = "" Or IsNull(PROJECT_NAME) Then
			PROJECT_NAME = ""
		End If
		product_name = app("product_name")
		If trim(product_name) = "" Or IsNull(product_name) Then
			product_name = ""
		End If	
		if len(product_name) > 30 then
			product_name = left(product_name,27) & "..."		
		end if
			
		User_Name = app("User_Name")		
		prod_id = app("product_id")
		quest_id = trim(app("questions_id"))
		
		sqlstr = "EXECUTE get_appeal_tasks '" & OrgID & "','" & appid & "'"	
	    set rsmess = con.getRecordSet(sqlstr)
		if not rsmess.eof then
			mes_new = rsmess("mes_new") : 	mes_work = rsmess("mes_work") : mes_close = rsmess("mes_close") 
		else
			mes_new = 0 : mes_work = 0 : mes_close = 0
		end if
		If trim(SURVEYS)  = "1" Then
		    href_ = " HREF=""../appeals/appeal_card.asp?quest_id=" & quest_id & "&appid=" & appid & """"
		Else
			href_ = " nohref"
		End If 
		appeal_status = trim(app("appeal_status")) 	%>
		<tr>		   
		    <td align="center" class="card" nowrap>&nbsp;<a  href="#" onclick="javascript:window.open('../appeals/view_appeal.asp?quest_id=<%=quest_id%>&appid=<%=appid%>','','top=100,left=100,width=500,height=500,scrollbars=1,resizable=1,menubar=1')"><IMG SRC="../../images/print_icon.gif" BORDER=0 hspace=0 vspace=0></a>&nbsp;</td>
		   	<td align="center" class="card" nowrap>						
			&nbsp;
			<a HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&task_status=3" target="_self" style="WIDTH:10pt;" class="task_status_num3"><%=mes_close%></a>
			<a HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&task_status=2" target="_self" style="WIDTH:10pt;" class="task_status_num2"><%=mes_work%></a>
			<a HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&task_status=1" target="_self" style="WIDTH:10pt;" class="task_status_num1"><%=mes_new%></a>
			&nbsp;		
			</td>		
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=CONTACT_NAME%>&nbsp;</a></td>						
			<td class="card" align="<%=align_var%>">
		    <!--#include file="../appeals/key_fields_t.asp"-->
			</td>			
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=product_name%>&nbsp;</a></td>			
			<td class="card" nowrap align=center><a class="link_categ" <%=href_%> target="_self"><%=appid%></a></td>
			<td class="card" nowrap align=center><a class="link_categ" <%=href_%> target="_self">&nbsp;<%=day(app("appeal_date"))%>/<%=month(app("appeal_date"))%>/<%=mid(year(app("appeal_date")),3,2)%>&nbsp;</a></td>			
			<td class="card" nowrap align=center><a class=status_num<%=appeal_status%> <%=href_%> target="_self"><%=arr_status_app(appeal_status)%></a></td>
		</tr>
<%		app.movenext
		j=j+1
		if not app.eof and j <> app.PageSize then
		ids = ids & ","
		end if
		loop 
		%>
		</table>		
		<input type="hidden" name="ids" value="<%=ids%>" ID="ids">
		</td></tr>
		</form>				
		<% if NumberOfPagesApp > 1 then%>
		<tr class="card">
		<td width="100%" align=center nowrap class="card" dir="<%=dir_var%>">
			<table border="0" cellspacing="0" cellpadding="2" dir="<%=dir_var%>">               
	        <% If NumberOfPagesApp > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPagesApp / 10)
	           else num = NumberOfPagesApp : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRowApp") <> nil Then
	               numOfRowApp = Request.QueryString("numOfRowApp")
	           Else numOfRowApp = 1
	           End If
	         %>
	         
	         <tr>
	         <%if numOfRowApp <> 1 then%> 
			 <td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=10*(numOfRowApp-1)-9%>&numOfRowApp=<%=numOfRowApp-1%>#table_appeals" name=word53 title="ל<%=arrTitles(53)%>"><b><<</b></a></td>			                
			  <%end if%>
	             <td><font size="2" color="#060165">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowApp-1)) <= NumberOfPagesApp Then
	                  if CInt(page_app)=CInt(i+10*(numOfRowApp-1)) then %>
		                 <td align="center"><span class="pageCounter"><%=i+10*(numOfRowApp-1)%></span></td>
	                  <%else%>
	                     <td align="center"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=i+10*(numOfRowApp-1)%>&numOfRowApp=<%=numOfRowApp%>#table_appeals"><%=i+10*(numOfRowApp-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td><font size="2" color="#060165">]</font></td>
				<%if NumberOfPagesApp > cint(num * numOfRowApp) then%>  
					<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=10*(numOfRowApp) + 1%>&numOfRowApp=<%=numOfRowApp+1%>#table_appeals" name=word54 title="<%=arrTitles(54)%>"><b>>></b></a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>
	<%End If%>										
	<%If app.recordCount = 0 Then%>
	<tr><td align=center class=card1>&nbsp;</td></tr>									
	<%End If%>			 			 
	</table></td></tr>
	<%End If%>	
<%set app = nothing	%>		

<%
	dim sortby_f(16)	
	sortby_f(1) = "appeal_date"
	sortby_f(2) = "appeal_date DESC"
	sortby_f(3) = "appeal_id"
	sortby_f(4) = "appeal_id DESC"
	sortby_f(5) = "PEOPLE_EMAIL"
	sortby_f(6) = "PEOPLE_EMAIL DESC"
	sortby_f(7) = "product_name,appeal_date"
	sortby_f(8) = "product_name DESC,appeal_date"
	sortby_f(9) = "PEOPLE_NAME"
	sortby_f(10) = "PEOPLE_NAME DESC"
	sortby_f(11) = "PEOPLE_COMPANY,appeal_date"
	sortby_f(12) = "PEOPLE_COMPANY DESC,appeal_date"
	sortby_f(13) = "appeal_status,appeal_date"
	sortby_f(14) = "appeal_status DESC,appeal_date"

	sort_f = Request("sort_f")	
	if sort_f = nil then
		sort_f = 2
	end if
	
	if lang_id = 1 then
		arr_status_f = Array("","חדש","בטיפול","סגור")
	else
		arr_status_f = Array("","new","active","close")
	end if		

	sqlstr = "EXECUTE get_feedbacks '','','','','" & OrgID & "','" & sortby_f(sort_f) & "','','','" & companyID & "','','','',''"
    'Response.Write sqlStr
	set app=con.GetRecordSet(sqlStr)
	feedbaks_count = app.RecordCount
	if Request("page_f")<>"" then
		page_f=request("page_f")
	else
		page_f=1
	end if
	if not app.eof then %>
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6">
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width="100%">  
  <tr><td width=100%><A name="table_feedbacks"></A>  
  <table cellpadding=0 cellspacing=0 width="100%" >
  <tr>  
  <td class="title_form" width=100%>&nbsp;<!--משובים מדיוור--><%=arrTitles(34)%>&nbsp;<font color="#E6E6E6">(<%=company_name%>)</font>&nbsp;</td>
  </tr>
  </table></td></tr>  	  	
  <tr>
<tr>
	<td width="100%" align="center" valign=top>

	<table BORDER="0" CELLSPACING="1" CELLPADDING="1" bgcolor="#FFFFFF" dir="<%=dir_var%>">	
	<tr>	    
		<td width="160"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort_f=11 or sort_f=12 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&sort_f=<%if sort_f=12 then%>11<%elseif sort_f=11 then%>12<%else%>12<%end if%>#table_appeals" target="_self">&nbsp;Email&nbsp;<img src="../../images/arrow_<%if trim(sort_f)="11" then%>bot<%elseif trim(sort_f)="12" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
		<td width="150"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort_f=5 or sort_f=6 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&sort_f=<%if sort_f=6 then%>5<%elseif sort_f=5 then%>6<%else%>6<%end if%>#table_appeals" target="_self">&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort_f)="5" then%>bot<%elseif trim(sort_f)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>			
		<td width="100%" id=td_prod name=td_prod align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort_f=7 or sort_f=8 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&sort_f=<%if sort_f=7 then%>8<%elseif sort_f=8 then%>7<%else%>7<%end if%>#table_appeals" target="_self">&nbsp;<span id="word35" name=word35><!--סוג טופס--><%=arrTitles(35)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_f)="7" then%>bot<%elseif trim(sort_f)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="60" nowrap  align=center dir="<%=dir_obj_var%>" class="title_sort<%if sort_f=1 or sort_f=2 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort_f=<%if sort_f=2 then%>1<%elseif sort_f=1 then%>2<%else%>2<%end if%>#table_appeals" target="_self">&nbsp;<span id="word36" name=word36><!--תאריך--><%=arrTitles(36)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_f)="1" then%>bot<%elseif trim(sort_f)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
		<td width="40" nowrap align=center dir="<%=dir_obj_var%>" class="title_sort<%if sort_f=3 or sort_f=4 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort_f=<%if sort_f=4 then%>3<%elseif sort_f=3 then%>4<%else%>4<%end if%>#table_feedbacks" target="_self">&nbsp;ID&nbsp;<img src="../../images/arrow_<%if trim(sort_f)="3" then%>bot<%elseif trim(sort_f)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="43" nowrap align=center dir="<%=dir_obj_var%>" class="title_sort<%if sort_f=13 or sort_f=14 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort_f=<%if sort_f=14 then%>13<%elseif sort_f=13 then%>14<%else%>14<%end if%>#table_feedbacks" target="_self"><span id="word37" name=word37><!--'סט--><%=arrTitles(37)%></span><img src="../../images/arrow_<%if trim(sort_f)="13" then%>bot<%elseif trim(sort_f)="14" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	</tr>
	<%	app.PageSize = 5
		app.AbsolutePage=page_f
		recCount=app.RecordCount 		
		NumberOfPagesF = app.PageCount
		i=1
		j=0
		ids = "" 'list of appeal_id
		do while (not app.eof and j<app.PageSize)
		appid=app("appeal_id")	
		PEOPLE_ID = app("PEOPLE_ID")		
		PEOPLE_EMAIL = app("PEOPLE_EMAIL")
		PEOPLE_COMPANY = app("PEOPLE_COMPANY")		
		If trim(PEOPLE_COMPANY) = "" Or IsNull(PEOPLE_COMPANY) Then
			PEOPLE_COMPANY = ""
		End If	
		PEOPLE_NAME = app("PEOPLE_NAME")
		If trim(PEOPLE_NAME) = "" Or IsNull(PEOPLE_NAME) Then
			PEOPLE_NAME = ""
		End If			
	
		groupID = app("groupID")		
		prod_id = app("product_id")
		quest_id = trim(app("questions_id"))
		If trim(quest_id) <> "" Then
			sqlstr = "Select PRODUCT_NAME FROM PRODUCTS WHERE PRODUCT_ID = " & quest_id
			set rs_name = con.getRecordSet(sqlstr)
			If not rs_name.eof Then
				if len(rs_name("product_name")) > 30 then
					product_name = left(rs_name("product_name"),27) & "..."
				else
					product_name = rs_name("product_name")
				end if	
			End if
			set rs_name = Nothing
		End If	
		If trim(SURVEYS)  = "1" Then
		    href_ = " HREF='../appeals/feedback.asp?quest_id=" & quest_id & "&appid=" & appid & "'"
		Else
			href_ = " nohref"
		End If 
		feedback_status = trim(app("appeal_status"))			
		%>
		<tr>			
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=PEOPLE_EMAIL%>&nbsp;</a></td>
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=PEOPLE_NAME%>&nbsp;</a></td>			
			<td class="card" nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=product_name%>&nbsp;</a></td>
			<td class="card" align=center><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=day(app("appeal_date"))%>/<%=month(app("appeal_date"))%>/<%=mid(year(app("appeal_date")),3,2)%>&nbsp;</a></td>
			<td class="card" nowrap align=center><a class="link_categ" <%=href_%> target="_self" dir="<%=dir_obj_var%>"><%=appid%></a></td>
			<td class="card" nowrap align=center><a class=status_num<%=feedback_status%> <%=href_%> target="_self"><%=arr_status_f(feedback_status)%></a></td>
		</tr>
<%		app.movenext
		j=j+1
		if not app.eof and j <> app.PageSize then
		ids = ids & ","
		end if
		loop 
		%>
		<% if NumberOfPagesF > 1 then%>
		<tr class="card">
		<td width="100%" align=center class="card" colspan="6">
			<table border="0" cellspacing="0" cellpadding="2">
	        <% If NumberOfPagesF > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPagesF / 10)
	           else num = NumberOfPagesF : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRowF") <> nil Then
	               numOfRowF = Request.QueryString("numOfRowF")
	           Else numOfRowF = 1
	           End If
	         %>
	         
	         <tr>
	         <%if numOfRowF <> 1 then%> 
			 <td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&sort_f=<%=sort_f%>&page_f=<%=10*(numOfRowF-1)-9%>&numOfRowF=<%=numOfRowF-1%>" name=word53 title="<%=arrTitles(53)%>"><b><<</b></a></td>			                
			  <%end if%>
	             <td><font size="2" color="#060165">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowF-1)) <= NumberOfPagesF Then
	                  if CInt(page_f)=CInt(i+10*(numOfRowF-1)) then %>
		                 <td align="center"><span class="pageCounter"><%=i+10*(numOfRowF-1)%></span></td>
	                  <%else%>
	                     <td align="center"><a class="pageCounter" href="<%=urlSort%>&sort_f=<%=sort_f%>&page_f=<%=i+10*(numOfRowF-1)%>&numOfRowF=<%=numOfRowF%>"><%=i+10*(numOfRowF-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td><font size="2" color="#060165">]</font></td>
				<%if NumberOfPagesF > cint(num * numOfRowF) then%>  
					<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&sort_f=<%=sort_f%>&page_f=<%=10*(numOfRowF) + 1%>&numOfRowF=<%=numOfRowF+1%>" name=word54 title="<%=arrTitles(54)%>"><b>>></b></a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>				
	<%End If%>	
	<%set app = nothing	%>
		
</table></td></tr>
</table></td></tr>
<%End If%>	
<%	end if ' if not app.eof%>
   <!-----------------------------------------------------הפצות------------------------------------------>
   <%		
	dim sortPbyP(16)	
	sortPbyP(1) = 1'" CAST(DATE_SEND as float), PRODUCT_NAME"
	sortPbyP(2) = 2'" CAST(DATE_SEND as float) DESC, PRODUCT_NAME"
	sortPbyP(3) = 3'" CAST(DATE_ANSWER as float), PRODUCT_NAME"
	sortPbyP(4) = 4'" CAST(DATE_ANSWER as float) DESC, PRODUCT_NAME"
	sortPbyP(5) = 5'" PRODUCT_NAME"
	sortPbyP(6) = 6'" PRODUCT_NAME DESC"
	sortPbyP(7) = 7'" Page_Title"
	sortPbyP(8) = 8'" Page_Title DESC"
	
	sortP = Request("sortP")	
	if sortP = nil then
		sortP = 2
	end if
	
	'sqlStr = "Select Page_Id, QUESTIONS_ID, DATE_SEND, DATE_ANSWER, PRODUCT_NAME, Page_Title,"&_
	'" FILE_ATTACHMENT, IS_OPENED, PRODUCT_ID, PEOPLE_NAME, PEOPLE_ID, EMAIL_SUBJECT,"&_
	'" PRODUCT_TYPE From dbo.tbl_mailing('" & OrgID & "','" & companyID & "','')" &_
	'" ORDER BY " & sortPbyP(sortP) 
	sqlstr = "Exec dbo.get_mailing_list '" & OrgID & "','" & companyID & "','','" & sortPbyP(sortP) & "'" 
	'Response.Write sqlStr
    'Response.End
	set prod = con.GetRecordSet(sqlStr)
    if not prod.eof then
		PageSize = 5
		if Request("pageP")<>"" then
			pageP=request("pageP")
		else
			pageP=1
		end if		
	    prod.PageSize = PageSize
		prod.AbsolutePage=pageP
		recCountP=prod.RecordCount 	
		NumberOfPagesP = prod.PageCount
 %>
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6">
  <table cellpadding=0 cellspacing=0 dir="<%=dir_var%>" width=100%>  
  <tr><td width=100%><A name="table_mailing"></A>  
  <table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
  <tr>  
  <td class="title_form" width=100% align="<%=align_var%>" dir=<%=dir_obj_var%>>&nbsp;<!--הפצות--><%=arrTitles(65)%>&nbsp;<font color="#E6E6E6">(<%=company_name%>)</font>&nbsp;</td>
  </tr>
  </table></td></tr>  	  	
  <tr>
<tr>
	<td width="100%" align="center" valign=top dir=ltr>
	<table BORDER="0" CELLSPACING="1" CELLPADDING="1" bgcolor="#FFFFFF" dir="<%=dir_var%>">	
	<tr>
		<td valign=top nowrap class="title_sort" align=center><!--קובץ מצורף--><%=arrTitles(66)%></td>
		<td valign=top nowrap align=center class="title_sort"><!--נפתח--><%=arrTitles(67)%></td>
		<td valign=top nowrap align=center class="title_sort<%if sortP=3 or sortP=4 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sortP=<%if sortP=4 then%>3<%elseif sortP=3 then%>4<%else%>4<%end if%>#table_mailing" target="_self"><!--תאריך מענה--><%=arrTitles(68)%><img src="../../images/arrow_<%if trim(sortP)="3" then%>down<%elseif trim(sortP)="4" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td valign=top nowrap align=center class="title_sort<%if sortP=1 or sortP=2 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sortP=<%if sortP=2 then%>1<%elseif sortP=1 then%>2<%else%>2<%end if%>#table_mailing" target="_self"><!--תאריך שליחה--><%=arrTitles(69)%><img src="../../images/arrow_<%if trim(sortP)="1" then%>down<%elseif trim(sortP)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>			
		<td nowrap align=right class="title_sort<%if sortP=9 or sortP=10 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sortP=<%if sortP=10 then%>9<%elseif sortP=9 then%>10<%else%>10<%end if%>#table_mailing" target="_self"><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sortP)="9" then%>down<%elseif trim(sortP)="10" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>		
		<td width=150 valign=top nowrap align=right class="title_sort<%if sortP=5 or sortP=6 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sortP=<%if sortP=6 then%>5<%elseif sortP=5 then%>6<%else%>6<%end if%>#table_mailing" target="_self"><!--טופס מופץ--><%=arrTitles(70)%>&nbsp;<img src="../../images/arrow_<%if trim(sortP)="5" then%>down<%elseif trim(sortP)="6" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="100%" align=right valign=top class="title_sort<%if sortP=7 or sortP=8 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sortP=<%if sortP=8 then%>7<%elseif sortP=7 then%>8<%else%>8<%end if%>#table_mailing" target="_self"><!--דף מופץ--><%=arrTitles(71)%>&nbsp;<img src="../../images/arrow_<%if trim(sortP)="7" then%>down<%elseif trim(sortP)="8" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>		
		<td width=150 nowrap valign=top nowrap class="title_sort" align="right">&nbsp;<%=arrTitles(74)%>&nbsp;</td>
		<td valign=top nowrap class="title_sort" align="right">&nbsp;<%=arrTitles(73)%>&nbsp;</td>
	</tr>	
	<%			
		i=1
		j=0
		ids = "" 'list of appeal_id
		do while (not prod.eof and j<prod.PageSize)
			page_id = prod.Fields(0)
			questions_id = prod.Fields(1)
			date_send = prod.Fields(2)
			date_answer = prod.Fields(3)
			product_name = prod.Fields(4)
			Page_Title = prod.Fields(5)
			file_attachment = prod.Fields(6)
    		If Len(file_attachment) > 0 Then
    			attachment_path = "../../../download/products/" & FILE_ATTACHMENT
    		Else
    			attachment_path = ""	
    		End If	
    		is_opened = prod.Fields(7)    		
			prod_Id = prod.Fields(8)
			people_name = prod.Fields(9)			
			people_id = prod.Fields(10)
			email_subject = prod.Fields(11)			
    		product_type = prod.Fields(12)
    		If trim(product_type) = "3" Then
    			product_type = "דואר"
    		Else
    			product_type = "מייל"
    		End If%>
<tr>
	<td align="center" class="card"><%If Len(attachment_path) > 0 Then%><a href="<%=attachment_path%>" target=_blank class=link1><img src="../../images/bull2.gif" border=0 hspace=0 vspace=0></a><%Else%>&nbsp;<%End If%></td>
	<td align="center" class="card"><%If trim(is_opened) = "1" Then%><img src="../../images/vi.gif" border=0 hspace=0 vspace=0><%End If%></td>
	<td align="center" class="card"><%=date_answer%></td>
	<td align="center" class="card"><%=date_send%></td>	
	<td align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="card">&nbsp;<%=people_name%>&nbsp;</td>	
	<td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card"><%If IsNumeric(questions_id) Then%><a href="#" onclick="window.open('../products/check_form.asp?prodId=<%=questions_id%>','Preview','left=20,top=20,tollbar=0,menubar=0,scrollbars=1,resizable=0,width=660');" class="link_categ"><%=product_name%>&nbsp;</a><%Else%><%=product_name%>&nbsp;<%End If%></td>
	<td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card"><a href="#table_mailing" onclick="return openPreview('<%=page_id%>')" class="link_categ"><%=Page_Title%>&nbsp;</a></td>
	<td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card">&nbsp;<%=email_subject%>&nbsp;</td>
	<td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card">&nbsp;<%=product_type%>&nbsp;</td>
</tr>
<%		prod.MoveNext
        j=j+1
		loop
	set prod=nothing%>
<% if NumberOfPagesP > 1 then%>
	<tr class="card">
		<td width="100%" align=center nowrap class="card" colspan=10>
			<table border="0" cellspacing="0" cellpadding="2" ID="Table4">
	        <% If NumberOfPagesP > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPagesP / 10)
	           else num = NumberOfPagesP : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRowP") <> nil Then
	               numOfRowP = Request.QueryString("numOfRowP")
	           Else numOfRowP = 1
	           End If
	         %>
	         <tr>
	         <%if numOfRowP <> 1 then%> 
			 <td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&sortP=<%=sortP%>&pageP=<%=10*(numOfRowP-1)-9%>&numOfRowP=<%=numOfRowP-1%>" name=word56 title="<%=arrTitles(56)%>"><b><<</b></a></td>			                
			  <%end if%>
	             <td><font size="2" color="#060165">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowP-1)) <= NumberOfPagesP Then
	                  if CInt(pageP)=CInt(i+10*(numOfRowP-1)) then %>
		                 <td align="center"><span class="pageCounter"><%=i+10*(numOfRowP-1)%></span></td>
	                  <%else%>
	                     <td align="center"><a class="pageCounter" href="<%=urlSort%>&sortP=<%=sortP%>&pageP=<%=i+10*(numOfRowP-1)%>&numOfRowP=<%=numOfRowP%>"><%=i+10*(numOfRowP-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
					<td><font size="2" color="#060165">]</font></td>
				<%if NumberOfPagesP > cint(num * numOfRowP) then%>  
					<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&sortP=<%=sortP%>&pageP=<%=10*(numOfRowP) + 1%>&numOfRowP=<%=numOfRowP+1%>" name=word57 title="<%=arrTitles(57)%>"><b>>></b></a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>					
	<%End If%>		
	</table></td></tr>		
	</table></td></tr>
  <%end if ' if not app.eof%>
 <!-----------------------------------------------------הפצות------------------------------------------>
<%If trim(is_meetings) = "1" Then%>
<!-------------------------------------------------מפגשים-------------------------------------------------->
<%
	if lang_id = "1" then
		arr_Status_M = Array("","עתידית","הסתיימה","הוכנס סיכום","נדחתה")	
	else
		arr_Status_M = Array("","Future","Done","Summary added","Postponed")	
	end if
	if Request("PageM")<>"" then
		PageM=request("PageM")
	else
		PageM=1
	end if
	PageSizeM = 5
	'sqlStr = "Select DISTINCT meeting_date,meeting_id,start_time,end_time,meeting_status, "&_
	'" company_name, contact_name from meetings_view where ORGANIZATION_ID = "& OrgID &_
	'" AND Company_ID = " & CompanyID & " Order BY meeting_date, start_time, end_time"
	sqlstr = "EXECUTE get_meetings_paging " & PageM & "," & PageSizeM & ",'','','','','','" & OrgID & "',' meeting_date, start_time, end_time','','','" & CompanyID & "'"
	'Response.Write sqlStr
	'Response.End
	set rs_meetings = con.GetRecordSet(sqlStr)
	if not rs_meetings.EOF then
		recCountM = rs_meetings("CountRecords")
%>
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6"><A name="table_meetings"></A>  
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width=100%>  
  <tr>  
  <td class="title_form" width=100%>&nbsp;<!--פגישות--><%=arrTitles(60)%>&nbsp;<font color="#E6E6E6">(<%=company_name%>)</font>&nbsp;</td>
  </tr>
<tr>    
    <td width="100%" valign="top" align="center" colspan=3>    
	<table BORDER="0" CELLSPACING="1" CELLPADDING="1" bgcolor="#FFFFFF" dir="<%=dir_var%>">	
	<tr>		
	<td width="100%" nowrap align="center" class="title_sort"><!--תאור--><%=arrTitles(61)%></td>
	<td width="100" nowrap align="center" valign="top" class="title_sort"><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td>
	<td width="150" nowrap align="center" valign="top" class="title_sort"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
	<td width="70" nowrap align="center" class="title_sort"><!--שעת סיום--><%=arrTitles(62)%></td>
	<td width="70" nowrap align="center" class="title_sort"><!--שעת התחלה--><%=arrTitles(63)%></td>
	<td width="60" nowrap align="center" valign="top" class="title_sort"><!--תאריך--><%=arrTitles(32)%></td>
	<td width="70" align="center" dir="<%=dir_obj_var%>" nowrap class="title_sort"><!--'סט--><%=arrTitles(13)%>&nbsp;</td>
	</tr>
<%

 while not rs_meetings.EOF
	meetingID = rs_meetings(1)
	company_name = rs_meetings(4)
	contact_name = rs_meetings(5)
	meetingDate = rs_meetings(6) 
	status = rs_meetings(7)
	startTime = rs_meetings(8)
	endTime = rs_meetings(9)
	
	users_names= ""
	sqlstr = "Select FIRSTNAME + ' ' + LASTNAME From meeting_to_users Inner Join Users On Users.User_ID = meeting_to_users.User_ID " &_
	" Where meeting_ID = " & meetingID & " ORDER BY FIRSTNAME + ' ' + LASTNAME"
	set rs_participants = con.getRecordSet(sqlstr)
	if not rs_participants.eof then
		users_names= rs_participants.getString(,,",",",")
	end if
	set rs_participants = nothing
	If Len(users_names) > 1 Then
		users_names= Left(users_names,Len(users_names)-1)
	End If	
		
    If trim(status) = "1" Then
       href = "href=""javascript:addmeeting('" & meetingID & "')"""   
    Else
       href = "href=""javascript:closemeeting('" & meetingID & "')"""     
    End If	
%>
	<tr class="card">
		<td align="<%=align_var%>" valign="top" ><a class="link_categ" <%=href%>><%=users_names%></a></td>	    
		<td align="center" valign="top"><a class="link_categ" <%=href%>><%=contact_name%></a></td>
		<td align="center" valign="top"><a class="link_categ" <%=href%>><%=company_name%></a></td>
		<td align="center" valign="top"><a class="link_categ" <%=href%>><%=endTime%></a></td>
		<td align="center" valign="top"><a class="link_categ" <%=href%>><%=startTime%></a></td>
		<td align="<%=align_var%>" valign="top"><a class="link_categ" <%=href%>><%=FormatMediumDateShort(meetingDate)%></a></td>
		<td align="center" valign="top" dir="<%=dir_obj_var%>" nowrap><a class="task_status_num<%=status%>" <%=href%> style="width:64"><%=arr_Status_M(status)%></a></td>	  
	</tr>	
<% 
   rs_meetings.moveNext
   Wend  
   NumberOfPages = Fix((recCountM / PageSizeM)+0.9)
   'Response.Write recCount
   if NumberOfPages > 1 then
	  %>
	  <tr>
		<td width="100%" align=middle colspan=11 nowrap class="card">
			<table border="0" cellspacing="0" cellpadding="2" dir=ltr>               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRowM") <> nil Then
	               numOfRowM = Request.QueryString("numOfRowM")
	           Else numOfRowM = 1
	           End If	           
            %>	         
	         <tr>
	         <%if numOfRowM <> 1 then%> 
			 <td valign="center" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(20)%>" href="<%=urlSort%>&PageM=<%=10*(numOfRowM-1)-9%>&numOfRowM=<%=numOfRowM-1%>#table_meetings" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowM-1)) <= NumberOfPages Then
	                  if CInt(PageM)=CInt(i+10*(numOfRowM-1)) then %>
		                 <td align="middle" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRowM-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&PageM=<%=i+10*(numOfRowM-1)%>&amp;numOfRowM=<%=numOfRowM%>#table_meetings" ><%=i+10*(numOfRowM-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRowM) then%>  
					<td valign="center" align="<%=align_var%>"><A class=pageCounter title="<%=arrTitles(19)%>" href="<%=urlSort%>&PageM=<%=10*(numOfRowM) + 1%>&numOfRowM=<%=numOfRowM+1%>#table_meetings" >&gt;&gt;</a></td>
				<%end if%>
				</tr> 				  	             
			</table>
		</td>
	</tr>
	<!--tr><td  class="title_form" align=center colspan=9><%=recCount%> :ואצמנ תומושר כ"הס</td></tr-->										 
	<%End If%> 
 </table></td></tr>
 </table></td></tr>
 <%  End If
 set rs_meetings = Nothing 
  End If
%>	
</table>
<!---------------------------------------------------------------------------------------------------------->
</td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:106;" href="javascript:void(0)" onclick="javascript:window.open('newcontact.asp?companyID=<%=companyID%>','','top=50,left=10,resizable=0,width=720,height=465');">&nbsp;<span id=word38 name=word38><!--הוסף--><%=arrTitles(38)%></span>&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;</a></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:106;" href="javascript:void(0)" onclick="javascript:window.open('../projects/editProject.asp?companyID=<%=companyID%>','addProject','top=50,left=120,resizable=1,scrollbars=1,width=480,height=500');"><span id="word39" name=word39><!--הוסף--><%=arrTitles(39)%></span>&nbsp;<%=trim(Request.Cookies("bizpegasus")("Projectone"))%></a></td></tr>
<tr><td align="center" colspan=2><a class="button_edit_1" style="width:106;" href="javascript:void(0)" onclick="addtask('','<%=companyID%>','')"><span id="word40" name=word40><!--הוסף--><%=arrTitles(40)%></span>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></a></td></tr>
<tr><td align="center" colspan=2><a class="button_edit_1" style="width:106;" href="javascript:void(0)" onclick='window.open("adddoc.asp?companyID=<%=companyID%>", "Documents" ,"scrollbars=1,toolbar=0,top=150,left=120,width=450,height=200,align=center,resizable=0");'><span id="word41" name=word41><!--הוסף מסמך--><%=arrTitles(41)%></span>&nbsp;</a></td></tr>
<%If trim(is_meetings) = "1" Then%>
<tr><td align="center" colspan=2><a class="button_edit_1" style="width:106;" href="javascript:void(0)" onclick="addmeeting('')">&nbsp;<!--הוסף פגישה --><%=arrTitles(64)%>&nbsp;</a></td></tr>  
<%End If%>
<%If trim(private_flag) = "0" Then%>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:106;" href="javascript:void(0)" onclick="javascript:window.open('editcompany.asp?companyID=<%=companyID%>','','top=50,left=100,resizable=0,width=450,height=500');"><span id="word43" name=word43><!--עדכן פרטי--><%=arrTitles(43)%></span>&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></a></td></tr>	
<%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את ה" & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & " עם של המשימות, טפסים, מסמכים ופגישות"
     Else
		str_confirm = "Are you sure want to delete the " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & " with all tasks, forms, documents and meetings ?"
     End If   
%>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:106;" href="default.asp?delid=<%=companyID%>" ONCLICK='return window.confirm("<%=str_confirm%>")'><span id="word45" name=word45><!--מחק--><%=arrTitles(45)%></span>&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></a></td></tr>	
<%End If%>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:106;" href="javascript:void(0)" onclick="javascript:window.open('printcompany.asp?companyID=<%=companyID%>','','top=100,left=20,resizable=0,width=660,height=500,scrollbars=1,menubar=1');"><span id="word44" name=word44><!--הדפס--><%=arrTitles(44)%></span>&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></a></td></tr>	
<%If trim(SURVEYS)  = "1" Then%>
<tr><td height=5 nowrap></td></tr>
<tr><td align="center" colspan=2><a class="button_edit_2" href="javascript:void(0)" onclick="tfasimDropDown(this)"><img hspace=0 vspace=0 border=0 src="../../images/back_arrow.gif" <%If trim(lang_id) = "2" Then%>style="Filter: FlipH"<%End If%>>&nbsp;<!--צרף טופס--><%=arrTitles(42)%></a></td></tr>
<%End If%>
<tr><td height=10 nowrap></td></tr>
</table></td></tr>
<%End If%>
<tr><td height=10 nowrap></td></tr>
</table>
<DIV ID="task_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; width:100%; height:68; overflow:scroll; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #E6E6E6; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%  sqlstr = "Select activity_type_id, activity_type_name from activity_types Where ORGANIZATION_ID = "&OrgID&" Order By activity_type_id"
	set rsactivity = con.getRecordSet(sqlstr)
	If not rsactivity.eof Then
		TaskTypes = rsactivity.getRows()
	End If
	set rsactivity = Nothing	
	If IsArray(TaskTypes) Then
	For i=0 To Ubound(TaskTypes,2)%> 
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; text-align:right; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>&taskTypeID=<%=TaskTypes(0,i)%>#table_tasks'">
    <%=TaskTypes(1,i)%></DIV>
	<% Next
	   End If %>
	<DIV dir="<%=dir_obj_var%>"  onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>#table_tasks'"><!--כל הרשימה--><%=arrTitles(47)%></DIV>
</div>
</DIV>
<DIV ID="appeal_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; width:100%; height:68; overflow:scroll; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #E6E6E6; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%  'אירגון לא משתמש בקבוצות
	If is_groups = 0 Then
	sqlstr = "Select product_id, product_name from Products Where "&_
	" product_number = '0' and ORGANIZATION_ID=" & OrgID & " order by product_name"
	Else
	sqlstr = "Execute get_products_list '" & OrgID & "','" & UserID & "'"
	End If
	'Response.Write sqlstr
	'Response.End
	set rs_products = con.GetRecordSet(sqlstr)
	if not rs_products.eof then 
		ResProductsList = rs_products.getRows()		
	end if
	set rs_products=nothing				
	If IsArray(ResProductsList) Then
	For i=0 To Ubound(ResProductsList,2)%>
	<DIV dir="<%=dir_obj_var%>" onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; text-align:right; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>&productID=<%=ResProductsList(0,i)%>#table_appeals'">
    <%=ResProductsList(1,i)%>
    </DIV>
<%	Next	
	End If	
%>
	<DIV dir="<%=dir_obj_var%>"  onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>#table_appeals'"><!--כל הרשימה--><%=arrTitles(48)%></DIV>
</div>
</DIV>
<div id=div_comp_desc name=div_comp_desc style="display:none;visibility:hidden;position:absolute;left:30;top:435;width:350;height:100;z-index:11;">
<table cellpadding=0 cellspacing=1 border=1 bgcolor="#ffffff" width=100%>
 <tr>
        <td width="100%" bgcolor="#0F2771" align="<%=align_var%>">
            <table border="0" width="100%" cellspacing="0" cellpadding="0">    
            <tr>
				<td colspan=2 width="100%" bgcolor="#616161" height="1"></td>
			</tr>	
            <tr>
                <td width=20 nowrap align=center class=title_form><INPUT type=image src="../../images/close_icon.gif" border="0" onClick="return closeDesc();" vspace=0 hspace=0>              
                </td>            
                <td width=330 nowrap class=title_form align="<%=align_var%>">&nbsp;<!--פרטים נוספים--><%=arrTitles(49)%>&nbsp;</td>
             </tr>
             <tr>
				<td colspan=2 width="100%" bgcolor="#616161" height="1"></td>
			</tr>
			<tr>
            <td width="100%" colspan=2 align="<%=align_var%>" class=card style="padding:5px" dir="<%=dir_obj_var%>"><%=trim(company_desc)%></td>
            </tr>	
            </table>
        </td>  
  </tr>                     
</table>
</div>
<%If trim(SURVEYS)  = "1" Then%>
<!--#include file="tfasim_inc.asp"-->
<%End If%>	
</body>
<%set con=Nothing%>
</html>


<%@ Language=VBScript%>
<%Response.Buffer = True%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language=javascript>
<!--
    function addtask(contactID,companyID,taskID)
	{
		h = parseInt(520);
		w = parseInt(520);
		window.open("../tasks/addtask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskID=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}

	function closeTask(contactID,companyID,taskID)
	{
		h = parseInt(510);
		w = parseInt(520);
		window.open("../tasks/closetask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskId=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=50,left=100,width="+w+",height="+h+",align=center,resizable=0");
	}
	
	var oPopup = window.createPopup();
	function taskDropDown(obj)
	{
		oPopup.document.body.innerHTML = task_Popup.innerHTML; 
		oPopup.document.charset="windows-1255";
		oPopup.show(0, 17, obj.offsetWidth, 68, obj);    
	}
	
	function CheckDelDocument(AppealId,documentId) 
	{
		if (confirm("? האם ברצונך למחוק את המסמך"))
		{         
			document.location.href = "appeal_card.asp?appid="+AppealId+"&delDocumentID="+documentId;		    
		}
		return false;  
	}	
	
//-->
</script>
<%appid = Request.QueryString("appid")
	 quest_id = Request.QueryString("quest_id")		
	
	found_appeal = false
	If IsNumeric(appid) And Len(appid) > 0 Then		
    sqlstr = "Exec dbo.get_appeals '','','','','" & OrgID & "','','','','','','','" & appID & "'"
'	Response.Write("sqlstr=" & sqlstr)
	set app = con.GetRecordSet(sqlstr)
	if not app.eof then
	found_appeal = true
	appeal_date = FormatDateTime(app("appeal_date"), 2) & " " & FormatDateTime(app("appeal_date"), 4)
	companyName = app("company_Name")
	contactName = app("contact_Name")
	projectName = app("project_Name")
	userName = app("user_Name")
	appealUserID = app("user_ID")
	productName = app("product_Name")
	quest_id = app("questions_id")
	companyId = app("company_id")
	contactId = app("contact_id")
	projectId = app("project_id")
	appeal_status = trim(app("appeal_status"))
	appeal_status_name = trim(app("appeal_status_name"))
	appeal_status_color = trim(app("appeal_status_color"))
	closeDate = app("appeal_close_date")
	closeText = app("appeal_close_text")
	response_email = app("appeal_response_email")
	response_user_id = app("response_user_id")
	private_flag = app("private_flag")
	mechanismID = app("mechanism_ID")
	closingID = app("Closing_ID")
	If trim(closingID) <> "" Then
		sqlstr = "Select Closing_Name from Product_Closings WHERE Closing_ID = " & closingID
		set rs_name = con.getRecordSet(sqlstr)
		If not rs_name.eof Then
			ClosingName = trim(rs_name.Fields(0))
		End If
		set rs_name = Nothing
	End If	
	If trim(mechanismId) <> "" Then
		sqlstr = "Select mechanism_Name from mechanism WHERE mechanism_Id = " & mechanismId
		set rs_name = con.getRecordSet(sqlstr)
		If not rs_name.eof Then
			mechanismName = trim(rs_name.Fields(0))
		End If
		set rs_name = Nothing
	End If
	If trim(response_user_id) <> "" Then
		sqlstr = "Select FIRSTNAME + char(32) + LASTNAME From USERS Where USER_ID = " & response_user_id
		set rswrk = con.getRecordSet(sqlstr)
		If not rswrk.eof Then
			response_user_name = trim(rswrk(0))
		End If
		set rswrk = Nothing	
	End If	

	found_quest = false	
	sqlstr =  "Select Langu, Product_Name, RESPONSIBLE From Products WHERE PRODUCT_ID = " & quest_id & " And ORGANIZATION_ID = " & OrgID
	'Response.Write sqlstr
	'Response.End
	set rsq = con.getRecordSet(sqlstr)
	If not rsq.eof Then		
		Langu = trim(rsq(0))
		prodName = trim(rsq(1))		
		RESPONSIBLE = trim(rsq(2))
		found_quest = true
	Else
		found_quest = false		
	End If
	set rsq = Nothing	
	if Langu = "eng" then
		td_align = "left"
		pr_language = "eng"
	else
		td_align = "right"
		pr_language = "heb"
	end if
	
	'האם אני יכול להזין את הטופס
	sqlstr = "Select Products.product_id from Products Inner Join Users_To_Products "&_
	" ON Products.Product_ID = Users_To_Products.Product_ID WHERE Users_To_Products.User_ID = " & UserID &_
	" And Products.product_id = " & quest_id & " and Products.ORGANIZATION_ID = " & OrgID
	set rs_check = con.getRecordSet(sqlstr)
	If not rs_check.eof Then
		is_insert = true
	Else
		is_insert = false	
	End If
	set rs_check = Nothing
	
	Dim canEdit
	canEdit = false
	If trim(is_groups)  = "0" Then 'אין קבוצות
		canEdit = true
	ElseIf trim(appealUserID) = trim(UserID) Then	'אני מיליתי את הטופס
		canEdit = true
	ElseIf trim(EDIT_APPEAL) = "1" Then 'אני רשאי לעדכן טפסים
		canEdit = true
	ElseIf is_insert Then	'אני רשאי להזין טופס זה
		canEdit = true
	Else
		canEdit = false
	End If	
	
	If trim(appeal_status) = "1" Then
		sqlstring="UPDATE appeals set appeal_status = '2' WHERE ORGANIZATION_ID = " & OrgID & " AND appeal_id = " & appid 	
		'Response.Write(sqlstring)
		'Response.End 
		con.ExecuteQuery(sqlstring) 
   End If
   
   If trim(UserID) = trim(RESPONSIBLE) And trim(appeal_status) = "1" Then		
		appeal_status = "2" : status_name  = "בטיפול" : status_num = 2
		xmlFilePath = "../../../download/xml_forms/bizpower_forms.xml"
		'------ start deleting the new message from XML file ------
		set objDOM = Server.CreateObject("Microsoft.XMLDOM")
		objDom.async = false			
		if objDOM.load(server.MapPath(xmlFilePath)) then
				set objNodes = objDOM.getElementsByTagName("FORM")
				for j=0 to objNodes.length-1
					set objTask = objNodes.item(j)
					node_app_id = objTask.attributes.getNamedItem("APPEAL_ID").text										
					if trim(appId) = trim(node_app_id) Then					
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
		' ------ end  deleting the new message from XML file ------
	End If	
	
	set app = nothing
	Else
		found_appeal = false
	End If
	Else
		appeal_date = FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4)
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
		con.executeQuery("delete from appeals_documents where document_id = " & delDocumentID)
		Response.Redirect "appeal_card.asp?appId=" & appId
	End if 		
	
	dim sortby(12)	
	sortby(1) = "message_date"
	sortby(2) = "message_date DESC"
	sortby(3) = "MESSAGE_MAKER"
	sortby(4) = "MESSAGE_MAKER DESC"
	sortby(5) = "message_id"
	sortby(6) = "message_id DESC"
	sortby(7) = "message_status"
	sortby(8) = "message_status DESC"
	sortby(9) = "employee_name"
	sortby(10) = "employee_name DESC"

	sort_mes = Request.QueryString("sort_mes")	
	if sort_mes = nil then
		sort_mes = 6
	end if
%>
</head>
<%   
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 24 Order By word_id"				
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
<body topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0 bgcolor="#E5E5E5">
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr><td width=100% colspan=2>
<!-- #include file="../../logo_top.asp" -->
</td></tr>
<%numOftab = 1%>
<%numOfLink = 0%>
<tr><td width=100% colspan=2>
<!--#include file="../../top_in.asp"-->
</td></tr>
<%If found_appeal Then%>
<tr><td class="page_title" colspan=2 dir="<%=dir_obj_var%>"><font color="#6F6DA6"><%=productName%></font></td></tr>
<tr><td width="100%" valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0 dir="<%=dir_var%>">	   
<tr><td width="100%" valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0 dir="<%=dir_var%>">
<%If trim(appID) <> "" Then       
   		urlSort="appeal_card.asp?appID="& appID

        dim sortby_doc(2)			
        sortby_doc(0) = "D.document_id desc"
		sortby_doc(1) = "D.document_name"
		sortby_doc(2) = "D.document_name DESC "		
      
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
		
		sqlstr = "SELECT AD.appeal_id, D.document_id, D.document_name, D.document_desc, D.document_file " & _
		" FROM dbo.appeals_documents AD INNER JOIN  dbo.documents D ON AD.document_id = D.document_id " & _
		" WHERE AD.Appeal_ID = " & AppID& " ORDER BY " & sortby_doc(sort_doc)
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
  <table cellpadding=0 cellspacing=0 dir="<%=dir_obj_var%>" width="100%" >
  <tr><td width=100%><A name="table_doc"></A>  
  <table cellpadding=0 cellspacing=0 width="100%" >
  <tr>  
  <td class="title_form" width="100%">&nbsp;מסמכים מצורפים&nbsp;&nbsp;</td>
  </tr>
  </table></td></tr>  	  
  <tr>
  <td width="100%" valign=top>
   <table width=100% dir="<%=dir_var%>" border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF" >
    <tr> 
      <td align="center" class="title_sort" width=40 nowrap>מחק</td>     
      <td align="<%=align_var%>" class="title_sort" width="100%">&nbsp;תקציר&nbsp;</td>                  
      <td align="<%=align_var%>" width="193" nowrap class="title_sort<%if trim(sort_doc)="1" OR trim(sort_doc)="2" then%>_act<%end if%>"><%if trim(sort_doc)="1" then%><a class="title_sort" href="<%=urlSort%>&sort_doc=<%=sort_doc+1%>#table_doc" title="<%=arrTitles(20)%>"><%elseif trim(sort_doc)="2" then%><a class="title_sort" href="<%=urlSort%>&sort_doc=<%=sort_doc-1%>#table_doc" title="<%=arrTitles(21)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_doc=1#table_doc" title="<%=arrTitles(21)%>"><%end if%>&nbsp;כותרת מסמך&nbsp;<img src="../../images/arrow_<%if trim(sort_doc)="1" then%>bot<%elseif trim(sort_doc)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>
    </tr>   
 <%
      task_types_name = ""
      do while (not rsdoc.EOF and i<=rsdoc.PageSize)       
		documentID = trim(rsdoc("document_ID"))			
		document_name = trim(rsdoc("document_name"))
	    document_desc = trim(rsdoc("document_desc"))
	    document_file = trim(rsdoc("document_file"))
	    'href = " href = ""../../../download/documents/" & document_file & """"
	    href = " target='_blank' href = ""downloadFile.asp?document_file=" & Server.HTMLEncode(document_file) & """"
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
      <tr bgcolor="#E4E3EE">  
		<td align=center valign=middle><input type=image name=word55 src="../../images/delete_icon.gif" border=0 vspace=0 hspace=0 title="Delete document" onclick="return CheckDelDocument('<%=AppID%>','<%=documentID%>')" ID="Image2"></td>              
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>" title="<%=vFix(document_desc)%>">&nbsp;<%=document_desc_short%>&nbsp;</a></td>   
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top><a class="link_categ" <%=href%> dir="<%=dir_obj_var%>">&nbsp;<%=document_name%>&nbsp;</a></td>                         
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
			<table border="0" cellspacing="0" cellpadding="2" >
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
<%End if 
    set rsdoc = Nothing
    End If    
    
     If trim(appID) <> "" Then
		
		urlSort="appeal_card.asp?appID="& appID
		
        if lang_id = "1" then
	        arr_StatusT = Array("","חדש","בטיפול","סגור")	
        else
		    arr_StatusT = Array("","new","active","close")	
        end if 
        title_tasks = trim(Request.Cookies("bizpegasus")("TasksMulti"))
         
        dim sortby_task(12)			
		sortby_task(1) = "COMPANY_NAME"
		sortby_task(2) = "COMPANY_NAME DESC "
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
		Else taskTypeID = ""	
		End If	           
				
		task_status = trim(Request("task_status")) 
		
		If trim(task_status) <> "" Then		
			task_status = trim(Request("task_status")) 
			allTasks = "1"
		ElseIf trim(allTasks) = "1" Then			
			title_ = "פתוחות"
			task_status = "1,2"	
			allTasks = "1"	
		Else			
			title_ = ""	
			task_status = ""
			allTasks = "0"	
		End If	
   
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
    sqlstr = "EXEC dbo.get_tasks_paging " & page_task & "," & PageSize & ",'','','','" & task_status & "','" & UserID & "','" & OrgID & "','" & lang_id & "','" & taskTypeID & "','','','" & sortby_task(sort_task) & "','','','','','','" & appID & "'"
	'Response.Write sqlstr
	'Response.End   
	set tasksList = con.getRecordSet(sqlstr)
    urlSort = urlSort & "&allTasks=" & allTasks & "&task_status=" & task_status
    If not tasksList.eof Then
		recCount = tasksList("CountRecords")
	End If	
      
    If Not tasksList.eof Or taskTypeID <> "" Then%>
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6" valign=top>
  <table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
  <tr><td width=100%><A name="table_tasks"></A>  
  <table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
  <tr>  
  <td class="title_form" width=100% align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=title_tasks%>&nbsp;<font color="#E6E6E6">(<%=appid & " - " & productName%>)</font>&nbsp;</td> 
  </tr>
  </table></td></tr>  	  
  <tr>
  <td width="100%" valign=top dir="<%=dir_var%>">
   <table width=100% border=0 cellpadding=0 cellspacing=1 bgcolor="#FFFFFF" dir="<%=dir_var%>">
    <tr> 
      <td align=center class="title_sort" width=30 nowrap>&nbsp;</td>     
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort" width=270 nowrap><span id=word3 name=word3><!--תוכן--><%=arrTitles(3)%></span>&nbsp;</td>                  
      <td dir="<%=dir_obj_var%>" id=td_prod_type name=td_prod_type width=120 nowrap class="title_sort" align="<%=align_var%>">&nbsp;<span id=word4 name=word4><!--סוגי--><%=arrTitles(4)%></span>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%>&nbsp;<IMG name=word19 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(19)%>" align=absmiddle onmousedown="taskDropDown(td_prod_type)"></td>                  
      <%If trim(is_companies) = "1" Then%>
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=120 nowrap class="title_sort<%if trim(sort_task)="1" OR trim(sort_task)="2" then%>_act<%end if%>"><%if trim(sort_task)="1" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word20 title="<%=arrTitles(20)%>"><%elseif trim(sort_task)="2" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word21 title="<%=arrTitles(21)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=1#table_tasks" name=word21 title="<%=arrTitles(21)%>"><%end if%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="1" then%>bot<%elseif trim(sort_task)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>
      <%End If%>
      <td width="85" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_task)="9" OR trim(sort_task)="10" then%>_act<%end if%>"><%if trim(sort_task)="9" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>" name=word20 title="<%=arrTitles(20)%>"><%elseif trim(sort_task)="10" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>" name=word21 title="<%=arrTitles(21)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=9" name=word21 title="<%=arrTitles(21)%>"><%end if%><span id="word6" name=word6><!--אל--><%=arrTitles(6)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="9" then%>bot<%elseif trim(sort_task)="10" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
      <td width="85" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_task)="7" OR trim(sort_task)="8" then%>_act<%end if%>"><%if trim(sort_task)="7" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>" name=word20 title="<%=arrTitles(20)%>"><%elseif trim(sort_task)="8" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>" name=word21 title="<%=arrTitles(21)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=7" name=word21 title="<%=arrTitles(21)%>"><%end if%><span id=word5 name=word5><!--מאת--><%=arrTitles(5)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="7" then%>bot<%elseif trim(sort_task)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=75 nowrap class="title_sort<%if trim(sort_task)="3" OR trim(sort_task)="4" then%>_act<%end if%>"><%if trim(sort_task)="3" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word20 title="<%=arrTitles(20)%>"><%elseif trim(sort_task)="4" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word21 title="<%=arrTitles(21)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=3#table_tasks" name=word21 title="<%=arrTitles(21)%>"><%end if%><span id="word7" name=word7><!--תאריך יעד--><%=arrTitles(7)%></span><img src="../../images/arrow_<%if trim(sort_task)="3" then%>bot<%elseif trim(sort_task)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>          
      <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=50 nowrap class="title_sort<%if trim(sort_task)="5" OR trim(sort_task)="6" then%>_act<%end if%>"><%if trim(sort_task)="5" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task+1%>#table_tasks" name=word20 title="<%=arrTitles(20)%>"><%elseif trim(sort_task)="6" then%><a class="title_sort" href="<%=urlSort%>&sort_task=<%=sort_task-1%>#table_tasks" name=word21 title="<%=arrTitles(21)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_task=5#table_tasks" name=word21 title="<%=arrTitles(21)%>"><%end if%>&nbsp;<span id="word8" name=word8><!--'סט--><%=arrTitles(8)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_task)="5" then%>bot<%elseif trim(sort_task)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>
    </tr>   
 <%	current_date = Day(date()) & "/" & Month(date()) & "/" & Year(date())
	dim	 IS_DESTINATION
    task_types_name = ""
    while not tasksList.EOF       
		taskId = trim(tasksList(1))		
		COMPANY_NAME = trim(tasksList(4))		
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
		
		numOfLetters = 150

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
		Else
		   d_s=""
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
        End If		      %>           
      <tr>     
		<td align=center class="card<%=class_%>" valign=middle>
		<%If trim(taskID) <> "" And trim(childID) <> "" Then%>
		<input type=image src="../../images/hets4.gif" border=0 hspace=0 vspace=0 onclick='window.open("../tasks/task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=800,height=480,align=center,resizable=0");' ID="Image1" NAME="Image1">
		<%End If%>
		<%If trim(parentID) <> "" And trim(taskID) <> "" Then%>
		<input type=image src="../../images/hets4a.gif" border=0 hspace=0 vspace=0 onclick='window.open("../tasks/task_history.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=800,height=480,align=center,resizable=0");' ID="Image4" NAME="Image4">
		<%End If%>
		</td>              
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%> title="<%=vFix(tel_text)%>">&nbsp;<%=tel_text_short%>&nbsp;</a></td>
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%>>&nbsp;<%=task_types_names%>&nbsp;</a></td> 
        <%If trim(is_companies) = "1" Then%>                        
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%>>&nbsp;<%=COMPANY_NAME%>&nbsp;</a></td>      
        <%End If%>
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%>>&nbsp;<%=reciver_name%>&nbsp;</a></td>
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%>>&nbsp;<%=sender_name%>&nbsp;</a></td>      
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card<%=class_%>" valign=top><a class="link_categ" <%=href%> <%if IS_DESTINATION and task_status <> 3 then%> name=word22 title="<%=arrTitles(22)%>"><span style="width:9px;COLOR: #FFFFFF;BACKGROUND-COLOR: #FF0000;text-align:center"><b>!</b></span><%else%>><%end if%>&nbsp;<%=d_s%>&nbsp;</a></td>            
        <td align="center" dir="<%=dir_obj_var%>" class="card<%=class_%>" valign=top><a class="task_status_num<%=task_status%>"><%=arr_StatusT(task_status)%></A></td>
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
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word23 title="<%=arrTitles(23)%>" href="<%=urlSort%>&page_task=<%=10*(row_task-1)-9%>&amp;row_task=<%=row_task-1%>#table_tasks" >&lt;&lt;</a></td>			                
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
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word24 title="<%=arrTitles(24)%>" href="<%=urlSort%>&page_task=<%=10*(row_task) + 1%>&amp;row_task=<%=row_task+1%>#table_tasks" >&gt;&gt;</a></td>
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
    set tasksList = Nothing%>  

<%     
    dim sortby_response(6)			
	sortby_response(1) = "response_date"
	sortby_response(2) = "response_date DESC "
	sortby_response(3) = "sender_name"
	sortby_response(4) = "sender_name DESC"
	sortby_response(5) = "response_email"
	sortby_response(6) = "response_email DESC"
    
    sort_response = trim(Request.QueryString("sort_response"))
    If Len(sort_response) = 0 Then
		sort_response = 2
	Else sort_response = trim(Request.QueryString("sort_response"))	
    End If                		
   
	if trim(Request.QueryString("pageR"))<>"" then
		pageR=Request.QueryString("pageR")
	else
		pageR=1
	end if  
 
	if trim(Request.QueryString("rowR"))<>"" then
		rowR=Request.QueryString("rowR")
	else
		rowR = 1
	end if  
	
    sqlstr = "Select response_content ,response_date ,response_email, sender_name From appeal_responses_view " &_
    " Where appeal_id = " & appID & " Order BY " & sortby_response(sort_response)
	'Response.Write sqlstr
	'Response.End   
	set responsesList = con.getRecordSet(sqlstr)
    If not responsesList.eof Then
		PageSize = 5
		if Request("pageR")<>"" then
			pageR=request("pageR")
		else
			pageR=1
		end if		
	    responsesList.PageSize = PageSize
		responsesList.AbsolutePage=pageR
		recCountR=responsesList.RecordCount 	
		NumberOfPagesR = responsesList.PageCount
	%>
<tr>
  <td colspan="3" align="left" width="100%" bgcolor="#E6E6E6" valign=top>
  <table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
  <tr><td width=100%><A name="table_responses"></A>  
  <table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
  <tr>  
  <td class="title_form" width=100% align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<!--תגובות--><%=arrTitles(34)%>&nbsp;<font color="#E6E6E6">(<%=appid & " - " & productName%>)</font>&nbsp;</td> 
  </tr>
  </table></td></tr>  	  
  <tr>
  <td width="100%" valign=top dir="<%=dir_var%>">
   <table width=100% border=0 cellpadding=0 cellspacing=1 bgcolor="#FFFFFF" dir="<%=dir_var%>">
    <tr> 
      <td align="<%=align_var%>" class="title_sort" width=100%>&nbsp;<%=arrTitles(33)%>&nbsp;</td>     
      <td width="150" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_response)="5" OR trim(sort_response)="6" then%>_act<%end if%>"><%if trim(sort_response)="5" then%><a class="title_sort" href="<%=urlSort%>&sort_response=<%=sort_response+1%>" name=word20 title="<%=arrTitles(20)%>"><%elseif trim(sort_response)="6" then%><a class="title_sort" href="<%=urlSort%>&sort_response=<%=sort_response-1%>" name=word21 title="<%=arrTitles(21)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_response=5" name=word21 title="<%=arrTitles(21)%>"><%end if%><span id="Span3" name=word6><!--אל--><%=arrTitles(6)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_response)="5" then%>bot<%elseif trim(sort_response)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
      <td width="100" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort_response)="3" OR trim(sort_response)="4" then%>_act<%end if%>"><%if trim(sort_response)="3" then%><a class="title_sort" href="<%=urlSort%>&sort_response=<%=sort_response+1%>" name=word20 title="<%=arrTitles(20)%>"><%elseif trim(sort_response)="4" then%><a class="title_sort" href="<%=urlSort%>&sort_response=<%=sort_response-1%>" name=word21 title="<%=arrTitles(21)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_response=3" name=word21 title="<%=arrTitles(21)%>"><%end if%><span id="Span4" name=word5><!--מאת--><%=arrTitles(5)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort_response)="3" then%>bot<%elseif trim(sort_response)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
      <td width="70" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort_response)="1" OR trim(sort_response)="2" then%>_act<%end if%>"><%if trim(sort_response)="1" then%><a class="title_sort" href="<%=urlSort%>&sort_response=<%=sort_response+1%>#table_responses" name=word20 title="<%=arrTitles(20)%>"><%elseif trim(sort_response)="2" then%><a class="title_sort" href="<%=urlSort%>&sort_response=<%=sort_response-1%>#table_responses" name=word21 title="<%=arrTitles(21)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort_response=1#table_responses" name=word21 title="<%=arrTitles(21)%>"><%end if%><!--תאריך--><%=arrTitles(32)%><img src="../../images/arrow_<%if trim(sort_response)="1" then%>bot<%elseif trim(sort_response)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></td>          
    </tr>   
  <%j=0	
    do while (not responsesList.eof and j<responsesList.PageSize)   
		response_content = responsesList.Fields(0)
		response_date = responsesList.Fields(1)
		response_email = responsesList.Fields(2)
		sender_name = responsesList.Fields(3)
   %> 
          
      <tr>     		  
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card" style="padding-right:5px; padding-left: 5px" valign=top><%=response_content%></td>                         
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card" style="padding-right:5px; padding-left: 5px" valign=top><%=response_email%></td>      
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card" style="padding-right:5px; padding-left: 5px" valign=top><%=sender_name%></td>
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" class="card" style="padding-right:5px; padding-left: 5px" valign=top><%=response_date%></td>      
      </tr>    
<%
    responsesList.MoveNext
    j=j+1
	Loop	  

	if NumberOfPagesR > 1 then	
%>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2" dir=ltr>               
	        <% If NumberOfPagesR > 10 Then 
	              num = 10 : rowR = cInt(NumberOfPagesR / 10)
	           else num = NumberOfPagesR : rowR = 1    	                      
	           End If
            %>
	         
	         <tr>
	         <%if rowR <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word23 title="<%=arrTitles(23)%>" href="<%=urlSort%>&pageR=<%=10*(rowR-1)-9%>&amp;rowR=<%=rowR-1%>#table_responses" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(rowR-1)) <= NumberOfPagesR Then
	                  if CInt(pageR)=CInt(i+10*(rowR-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(rowR-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&pageR=<%=i+10*(rowR-1)%>&amp;rowR=<%=rowR%>#table_responses" ><%=i+10*(rowR-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPagesR > cint(num * rowR) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word24 title="<%=arrTitles(24)%>" href="<%=urlSort%>&pageR=<%=10*(rowR) + 1%>&amp;rowR=<%=rowR+1%>#table_responses" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>		
	<%
	End if 
	%>
	</table></td></tr>
	<%If responsesList.recordCount = 0 Then%>
	<tr><td align=center class=card1>&nbsp;</td></tr>									
	<%End If%>	 
</table></td></tr>
	<%
	End if   
    set responsesList = Nothing%>  
<%end if%>
<tr>
    <td width="100%" align="center" valign=top>
		<!-- start code --> 
		<table WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0" align="center" bgcolor=white>
		<tr bgcolor=#C9C9C9 >
		<td width="50%" nowrap align="<%=align_var%>" valign=top style="padding-right:15px; padding-left:15px;border: 1px #808080 solid">
		<TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1 valign=top>
		    <%If trim(appeal_status) = "3" Then%>
			<tr>
			 <td align="<%=align_var%>" width=100% valign=top><input type="text" style="width:120;" value="<%=vFix(closeDate)%>" class="Form_R" dir="LTR" ReadOnly ID="Text1" NAME="Text1"></td>
			 <td align="<%=align_var%>" width=130 nowrap class="form_title" valign=top><!--תאריך סגירה--><%=arrTitles(26)%></td>   
			</tr>
			<%If trim(ClosingName) <> "" Then%>
			<tr>
			 <td align="<%=align_var%>" width=100% valign=top><span style="width:350" dir="<%=dir_obj_var%>"><%=trim(ClosingName)%></span></td>
			 <td align="<%=align_var%>" width=130 nowrap class="form_title" valign=top><!--מצב סגירה--><%=arrTitles(36)%></td>   
			</tr>			
			<%End If%>
			<%If Len(closeText) > 0 Then%>
			<tr>
			 <td align="<%=align_var%>" width=100% valign=top><span style="width:350" dir="<%=dir_obj_var%>"><%=breaks(trim(closeText))%></span></td>
			 <td align="<%=align_var%>" width=130 nowrap class="form_title" valign=top><!--תוכן סגירה--><%=arrTitles(27)%></td>   
			</tr>			
			<%End If%> 
			<tr>
			 <td align="<%=align_var%>" width=100%><%=response_user_name%></td>
			 <td align="<%=align_var%>" width=130 nowrap class="form_title"><!--ע''י עובד--><%=arrTitles(29)%></td>   
			</tr>			 
			<%End If%>
		</table></td>	
		<td width="50%" nowrap align="<%=align_var%>" valign=top style="padding-right:15px; padding-left:15px;border-top: 1px #808080 solid;border-bottom: 1px #808080 solid">
		<TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=1 valign=top>
			<tr>
			 <td align="<%=align_var%>" width=100%><input type="text" style="width:110;" value="<%=appID%>"  class="Form_R" dir="ltr" ReadOnly></td>
			 <td align="<%=align_var%>" width=130 nowrap class="form_title">ID</td>   
			</tr>
			<tr>
			 <td align="<%=align_var%>" width=100%><input type="text" style="width:110;" value="<%=appeal_date%>"  class="Form_R" dir="ltr" ReadOnly></td>
			 <td align="<%=align_var%>" width=130 nowrap class="form_title"><span id=word9 name=word9><!--תאריך הזנת הטופס--><%=arrTitles(9)%></span></td>   
			</tr>
			<%If Len(userName) > 0 Then%>
			<tr>
			 <td align="<%=align_var%>" width=100%><input type="text" style="width:150;" value="<%=vFix(userName)%>" class="Form_R" dir="<%=dir_obj_var%>" ReadOnly></td>
			 <td align="<%=align_var%>" width=130 nowrap class="form_title"><span id="word10" name=word10><!--הזנה ע"י עובד--><%=arrTitles(10)%></span></td>   
			</tr>
			<%End If%>
			<tr>
			 <td align="<%=align_var%>" width=100%>
			  <span class="status_num" style="background-color:<%=trim(appeal_status_color)%>; text-align:center"><%=appeal_status_name%></span>
			 </td>
			 <td align="<%=align_var%>" width=130 nowrap class="form_title"><!--סטטוס--><%=arrTitles(11)%></td>   
		    </tr>
			<%If trim(companyId) <> "" And trim(private_flag) = "0" Then%>
			<tr>
			 <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=100%>
			 <%If trim(COMPANIES) = "1"  Then%>
			 <A href="../companies/company.asp?companyID=<%=companyId%>" target=_self class="links_down">
			 <%End If%>
			 <%=companyName%>
			 <%If trim(COMPANIES) = "1" And trim(private_flag) = "0" Then%>
			 </a>
			 <%End If%>
			 </td>
			 <td align="<%=align_var%>" width=130 nowrap class="links_down"><b><!--קישור ל--><%=arrTitles(12)%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></b></td>   
			</tr>
			<%End If%>
			<%If trim(contactID) <> "" Then%>
			<tr>
			 <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=100%>
			 <%If trim(COMPANIES) = "1" Then%>
			 <A href="../companies/contact.asp?contactID=<%=contactID%>" target=_self class="links_down">
			 <%End If%>
			 <%=contactName%>
			 <%If trim(COMPANIES) = "1" Then%>
			 </a>
			 <%End If%>
			 </td>
			 <td align="<%=align_var%>" width=130 nowrap class="links_down"><b><!--קישור ל--><%=arrTitles(13)%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></b></td>   
			</tr>					
			<%End If%>
			<%If trim(projectID) <> "" Then%>
			<tr>
			 <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=100%>
			 <%If trim(COMPANIES) = "1" Then%>
			 <A href="../projects/project.asp?companyID=<%=companyId%>&projectID=<%=projectID%>" target=_self class="links_down">
			 <%End If%>
			 <%=projectName%>
			 <%If trim(COMPANIES) = "1" Then%>
			 </a>
			 <%End If%>
			 </td>
			 <td align="<%=align_var%>" width=130 nowrap class="links_down"><b><!--קישור ל--><%=arrTitles(14)%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%></b></td>   
			</tr>
			<%End If%>
			
			<%If trim(mechanismID) <> "" Then%>
			<tr>
			 <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=100%>
			 <%If trim(COMPANIES) = "1" Then%>
			 <A href="../projects/project.asp?companyID=<%=companyId%>&projectID=<%=projectID%>" target=_self class="links_down">
			 <%End If%>
			 <%=mechanismName%>
			 <%If trim(COMPANIES) = "1" Then%>
			 </a>
			 <%End If%>
			 </td>
			 <td align="<%=align_var%>" width=130 nowrap class="links_down"><b><!--קישור ל--><%=arrTitles(14)%><%=arrTitles(35)%></b></td>   
			</tr>
			<%End If%>
									
			<tr><td height=10 colspan=2 nowrap></td></tr> 
			</table></td>			
			</tr>			
			<tr>
				<td width="100%" colspan="2" align="<%=align_var%>">
					<TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=3 bgcolor=White>					
					<!--#INCLUDE FILE="appeal_fields.asp"-->	
				</TABLE>
			</td>
		</tr>
		</table></td></tr>  
</table></td>
<td width=110 nowrap align="<%=align_var%>" valign=top style="border: 1px solid #808080;border-top:none;" class="td_menu">
<table cellpadding=1 cellspacing=1 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:100;line-height:110%;padding:4px;" href="#" onclick='window.open("../tasks/addtask.asp?appealID=<%=appid%>", "AddTask" ,"scrollbars=1,toolbar=0,top=20,left=120,width=430,height=530,align=center,resizable=0");'><!--הוסף--><%=arrTitles(15)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksOne"))%></a></td></tr>
<%If canEdit Then%>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:100;line-height:110%;padding:4px;" href="appeal.asp?quest_id=<%=quest_id%>&appid=<%=appid%>" dir=rtl><span id="word16" name=word16><!--עדכן טופס--><%=arrTitles(16)%></span></a></td></tr>
<%End If%>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:100;line-height:110%;padding:4px;" href="#" onclick='window.open("appeal_close.asp?appId=<%=appId%>","CloseAppeal","scrollbars=1,toolbar=0,top=150,left=50,width=580,height=300,align=center,resizable=0")' dir=rtl><!--סגור--><%=arrTitles(25)%></a></td></tr>	
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:100;line-height:110%;padding:4px;" href="#" onclick='window.open("send_reply.asp?appId=<%=appId%>","SendReplay","scrollbars=1,toolbar=0,top=150,left=50,width=580,height=350,align=center,resizable=0")' dir=rtl><!--שלח תגובה--><%=arrTitles(31)%></a></td></tr>	
<tr><td align="center" colspan=2><a class="button_edit_1" style="width:100;" href="javascript:void(0)" onclick='window.open("adddoc.asp?appealID=<%=appid%>", "Documents" ,"scrollbars=1,toolbar=0,top=150,left=120,width=450,height=200,align=center,resizable=0");'>צרף מסמך&nbsp;</a></td></tr>
<%
	If trim(lang_id) = "1" Then
		str_confirm = "?האם ברצונך למחוק את הטופס המלא"
	Else
		str_confirm = "Are sure want to delete the form?"
	End If	
%>
<%If (trim(EDIT_APPEAL) = "1") Or (trim(appealUserID) = trim(UserID)) Or (trim(is_groups)  = "0") Or isNumeric(appealUserID) = false Then%>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:100;line-height:110%;padding:4px;" href="appeals.asp?appid=<%=appid%>&prodID=<%=quest_id%>&delProd=1" ONCLICK="return window.confirm('<%=str_confirm%>')" dir=rtl><span id="word17" name=word17><!--מחק טופס--><%=arrTitles(17)%></span></a></td></tr>	
<%End If%>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:100;line-height:110%;padding:4px;" href="#" onclick="javascript:window.open('../appeals/view_appeal.asp?quest_id=<%=quest_id%>&appid=<%=appid%>','','top=20,left=50,width=600,height=550,scrollbars=1,resizable=1,menubar=1')"><!--הדפס טופס--><%=arrTitles(30)%></a></td></tr>
<tr><td colspan=2 height=10 nowrap></td></tr>
</table>
</td>
</tr> 
<%End If%>   
</table>
<DIV ID="task_Popup" STYLE="display:none;">
<div dir=<%=dir_obj_var%> style="position:absolute; top:0; left:0; width:100%; height:68; overflow:scroll; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #E6E6E6; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%  sqlstr = "Select activity_type_id, activity_type_name from activity_types WHERE ORGANIZATION_ID = "&OrgID&" Order By activity_type_id"
	set rsactivity = con.getRecordSet(sqlstr)
	while not rsactivity.eof %>
	<DIV dir=<%=dir_obj_var%> onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; text-align:right; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>&taskTypeID=<%=rsactivity(0)%>#table_tasks'">
    <%=rsactivity(1)%>
    </DIV>
	<%
		rsactivity.moveNext
		Wend
		set rsactivity=Nothing
	%>
	<DIV dir="<%=dir_obj_var%>"  onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#E6E6E6'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#E6E6E6; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='<%=urlSort%>#table_tasks'">
    <span id=word18 name=word18><!--כל הרשימה--><%=arrTitles(18)%></span>
    </DIV>
</div>
</DIV>
<%set con=nothing%>
</body>
</html>

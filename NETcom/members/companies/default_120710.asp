<%Response.Buffer = True%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--	
	function refreshCompany(companyObj)
	{
		if ((event.keyCode==13) || (event.keyCode == 9))
			return false;
		compName = new String(companyObj.value);
		if(compName.length > 0)
		{		
			window.document.all("frameCustomers").src = "companies.asp?search_company=" + compName;	
		}
		else
		{
			window.document.all("frameCustomers").src = "companies.asp";
		}	
		window.focus();
		companyObj.focus();	
		return true;	
				
	}

	function refreshContacts(contactObj)
	{
		if ((event.keyCode==13) || (event.keyCode == 9))
			return false;
		contName = new String(contactObj.value);
		if(contName.length > 0)
		{
			window.document.all("frameCustomers").src = "contacts.asp?" + contactObj.id + "=" + contName;	
		}
		else
		{
			window.document.all("frameCustomers").src = "contacts.asp";
		}	
				
	}	
//-->
</script>
</head>
<%TypeId = Request.QueryString("TypeId")	
	if trim(TypeId)="" then  TypeId = 1  end if
  
    delid=Request.QueryString("delid")
    if delid<>nil and delId<>"" then     	
	'-----deleting files----
	set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
	sqlstr = "select * from company_documents_view where company_id = "& delid
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
	con.executeQuery("delete from DOCUMENTS where document_id IN (Select document_id FROM company_documents WHERE company_id = " & delid & ")")
	con.executeQuery("delete from company_documents where company_id = " & delid)
	'-----deleting files----
	
	'-------------------- deleting from tasks xml ---------------------------------------------
	sqlstr = "Select DISTINCT task_id From tasks WHERE company_id = " & delid
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

	sqlstr = "Select DISTINCT task_id From tasks WHERE company_id = " & delid
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
	sqlstr = "Select DISTINCT appeal_id From appeals WHERE company_id = " & delid
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
		
	con.ExecuteQuery "DELETE FROM Contacts WHERE COMPANY_ID =" & delid    
	con.ExecuteQuery "DELETE FROM companies WHERE company_Id=" & delid	
	con.ExecuteQuery "DELETE FROM FORM_PROJECT_VALUE WHERE PROJECT_ID IN (Select PROJECT_ID FROM PROJECTS WHERE company_Id=" & delid & ")"
	con.ExecuteQuery "DELETE FROM projects WHERE company_Id=" & delid	
	con.ExecuteQuery "DELETE FROM tasks WHERE company_Id=" & delid	
	con.ExecuteQuery "DELETE FROM Form_Value WHERE APPEAL_ID IN (Select appeal_id FROM appeals WHERE company_Id=" & delid & ")"
	con.ExecuteQuery "DELETE FROM appeals WHERE company_Id=" & delid
	con.ExecuteQuery "DELETE FROM pricing_to_projects WHERE company_Id=" & delid		
	con.ExecuteQuery "DELETE FROM hours WHERE company_Id=" & delid	
	con.ExecuteQuery "DELETE FROM meetings WHERE company_Id=" & delid	
  end if
  
  search = false : search_company = ""	: search_contact = ""	  
  
  if trim(Request.Form("search_company")) <> "" Or trim(Request.QueryString("search_company")) <> "" then
		search_company = trim(Request.Form("search_company"))
		if trim(Request.QueryString("search_company")) <> "" then
			search_company = trim(Request.QueryString("search_company"))
		end if							
		search = true			
   else
		search_company = ""			
   end if

   if trim(Request.Form("search_contact")) <> "" Or trim(Request.QueryString("search_contact")) <> "" then
			search_contact = trim(Request.Form("search_contact"))
			if trim(Request.QueryString("search_contact")) <> "" then
				search_contact = trim(Request.QueryString("search_contact"))
			end if					
			search = true
	else
			search_contact = ""		
	end if
	
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 2 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
	arrTitlesD = rstitle.getRows()		
	redim arrTitles(Ubound(arrTitlesD,2)+1)
	For i=0 To Ubound(arrTitlesD,2)		
		arrTitles(arrTitlesD(0,i)) = arrTitlesD(1,i)		
	Next
	End If
	set rstitle = Nothing	%>
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td width="100%" align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width="100%" align="<%=align_var%>">
  <%numOftab = 0%>
  <%numOfLink = 0 + (TypeId - 1)%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title">&nbsp;<!--אנא הקלד את שם החברה מימין או את שם איש הקשר לאיתור במאגר הלקוחות--><%=arrTitles(25)%></td></tr>
<tr>    
    <td width="100%" valign="top" align="center">
    <table dir="<%=dir_var%>" border="0" cellpadding="0" cellspacing="1" width="100%">
    <tr>
    <td align="left" width="100%" valign=top >   
    <iframe name="frameCustomers" id="frameCustomers" <%If trim(search_contact) = "" And trim(search_company) = "" And TypeId = 1 Then%> src="contacts.asp" 
   <%ElseIf trim(search_company) <> "" or TypeId = 2 Then%> src="companies.asp?search_company=<%=Server.URLEncode(search_company)%>" 
   <%Else%> src="contacts.asp?search_contact=<%=Server.URLEncode(search_contact)%>" <%End If%>
	ALLOWTRANSPARENCY=true height=420 width="100%" marginwidth="0" marginheight="0" hspace="0" vspace="0" scrolling="0" frameborder="0"></iframe>
    </td>
    <td width=100 nowrap valign=top class="td_menu" style="border: 1px solid #808080" dir="<%=dir_var%>">
    <table cellpadding=1 cellspacing=0 width="100%" border="0" dir="<%=dir_var%>">
    <tr><td align="<%=align_var%>" colspan=2 class="title_search"><!--חיפוש--><%=arrTitles(11)%></td></tr>
    <%If trim(is_companies) = "1" Then%>
    <FORM action="default.asp?TypeId=<%=TypeId%>" method=POST id=form_search name=form_search target="_self">   
    <tr><td colspan=2 align="<%=align_var%>" class="right_bar"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td></tr>
    <tr>
    <td align="<%=align_var%>" width="15" nowrap><input type="image" onclick="form_search.submit();" src="../../images/search.gif" ></td>
    <td align="<%=align_var%>" width="85"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:80;" value="<%=vFix(search_company)%>" name="search_company" ID="search_company" onKeyUp="return refreshCompany(this)" onKeyDown="return refreshCompany(this)"></td></tr>
    </FORM>
    <%End If%>
    <FORM action="default.asp?TypeId=<%=TypeId%>" method=POST id="form_search_contact" name="form_search_contact" target="_self">   
    <tr><td colspan=2 align="<%=align_var%>" class="right_bar"><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td></tr>
    <tr><td align="<%=align_var%>" width="15" nowrap><input type="image" onclick="form_search.submit();" src="../../images/search.gif"></td>
    <td align="<%=align_var%>" width="85" nowrap><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:80;" value="<%=vFix(search_contact)%>" name="search_contact" ID="search_contact"  onKeyUp="refreshContacts(this)" onKeyDown="refreshContacts(this)"></td></tr>
    </FORM>
    <FORM action="default.asp?TypeId=<%=TypeId%>" method=POST id="Form1" name="form_search_cellular" target="_self">   
    <tr><td colspan=2 align="<%=align_var%>" class="right_bar">טלפון נייד</td></tr>
    <tr><td align="<%=align_var%>" width="15" nowrap><input type="image" onclick="form_search.submit();" src="../../images/search.gif" ></td>
    <td align="<%=align_var%>" width="85" nowrap><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:80;" value="<%=vFix(search_cellular)%>" name="search_cellular" ID="search_cellular"  onKeyUp="refreshContacts(this)" onKeyDown="refreshContacts(this)"></td></tr>
    </FORM>    
	<tr><td colspan=2 height=10 nowrap></td></tr>
	<%If trim(is_companies) = "1" Then%>
	<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:95;" href="javascript:void(0)" onclick="javascript:window.open('editcompany.asp','NewCompany','top=50,left=100,resizable=0,width=460,height=420');"><!--הוסף--><%=arrTitles(12)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></a></td></tr>
	<%End If%>
	<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:95;" href="javascript:void(0)" onclick="javascript:window.open('newcontact.asp','NewContact','top=50,left=10,resizable=0,width=750,height=500');"><!--הוסף--><%=arrTitles(13)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></a></td></tr>
	<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:95;line-height:110%;padding:3px" href="javascript:void(0)" onclick="javascript:window.open('upload.asp','Upload','top=20,left=40,resizable=0,width=600,height=500,scrollbars=1');"><!--ייבוא--><%=arrTitles(23)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsMulti"))%></a></td></tr>
	<%If trim(SURVEYS)  = "1" Then%>
	<tr><td colspan=2 height=5 nowrap></td></tr>
	<tr><td align="center" colspan=2><a class="button_edit_2" href="javascript:void(0)" onclick="tfasimDropDown(this)"><img hspace=0 vspace=0 border=0 src="../../images/back_arrow.gif" <%If trim(lang_id) = "2" Then%>style="Filter: FlipH"<%End If%>>&nbsp;<!--צרף טופס--><%=arrTitles(14)%></a></td></tr>
	<%End If%>
	<tr><td colspan=2 height=10 nowrap></td></tr>
	</table>
	</td></tr></table>
	</td></tr></table>
<%If trim(SURVEYS)  = "1" Then%>
<!--#include file="tfasim_not_popup_inc.asp"-->
<%End If%>	
</body>
</html>
<%set con=Nothing%>
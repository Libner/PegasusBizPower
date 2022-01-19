<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%contactID = trim(Request.QueryString("contactID"))
	urlSort="contact_appeals.asp?contactID="&contactID
	
	arr_Status = session("arr_Status")  :  app_status = "3"
	if lang_id = "1" then
	    arr_StatusT = Array("","חדש","בטיפול","סגור")	
    else
		arr_StatusT = Array("","new","active","close")	
    end if 		
    
	arch = trim(Request("arch"))
	if arch = "" then
		arch = 0
	end if

	if arch = "0" then
		set_arch = 1
	else
		set_arch = 0
	end if    
	
	If Request.Form.Count > 0 Then			
		' ------ transfering to archive --------
		if trim(Request.Form("trapp")) <> "" And trim(Request.Form("delete_flag")) = "0" _
		And trim(Request.Form("change_status_flag")) = "0" _
		And (trim(Request.Form("ContactId")) = "" OR trim(Request.Form("ContactId")) = "0") then
			sqlstr = "UPDATE appeals SET appeal_deleted=" & set_arch & " WHERE appeal_id IN (" & Request.Form("trapp") & ")"
			'Response.Write(sqlstr)
			'Response.End
			con.ExecuteQuery(sqlstr)
		end if
		' ------ transfering to archive --------

		'------ start changing appeals statuses --
		if trim(Request.Form("trapp")) <> "" And trim(Request.Form("change_status_flag")) = "1" then
			If isNumeric(Request.Form("cmb_change_status")) And trim(Request.Form("cmb_change_status")) <> "" Then
				int_change_status = cInt(Request.Form("cmb_change_status"))
			Else
				int_change_status = 3
			End If	
			sqlstr = "UPDATE APPEALS SET APPEAL_STATUS=" & int_change_status & ", appeal_close_date = getDate(), " & _
			" response_user_id = " & UserID & " WHERE Contact_Id = " & contactID & _
			" AND appeal_id IN (" & Request.Form("trapp") & ")"
			'Response.Write(sqlstr)
			'Response.End
			con.ExecuteQuery(sqlstr)
		end if
		'------ end changing appeals statuses ---
		
		'--------start move appeals to another contact person
		If trim(Request.Form("trapp")) <> "" And trim(Request.Form("ContactId")) <> "" And trim(Request.Form("ContactId")) <> "0" And isNumeric(Request.Form("ContactId")) then
			sqlstr = "UPDATE APPEALS SET CONTACT_ID = " & trim(Request.Form("ContactId")) & " WHERE APPEAL_ID IN (" & Request.Form("trapp") & ")"
			'Response.Write(sqlstr)
			'Response.End
			con.ExecuteQuery(sqlstr)
		End If
		'--------end move appeals to another contact person

		' ------ delete appeals --------
		if trim(Request.Form("delete_flag")) = "1" then
			sqlstr = "SELECT appeal_id FROM appeals WHERE appeal_id IN (" & Request.Form("trapp") & ")"
			'Response.Write(sqlstr)
			'Response.End
			set rs_delete = con.getRecordSet(sqlstr)
			While not rs_delete.eof
					delappid = rs_delete.Fields(0)		
					
					set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
					'-----deleting files----
					sqlstr = "SELECT D.document_file FROM dbo.appeals_documents AD INNER JOIN  dbo.documents D " & _
					" ON AD.document_id = D.document_id WHERE appeal_id = " & delappid
					'Response.Write sqlstr
					'Response.End
					set files = con.getRecordSet(sqlstr)
					do while not files.eof
						file_path="../../../download/documents/" & trim(files("document_file"))
						if fs.FileExists(server.mappath(file_path)) then
							set a = fs.GetFile(server.mappath(file_path))
							a.delete '------------------------------		
						end if	
					files.movenext
					loop
					set files =nothing
					set fs = nothing
						
					'------ start deleting the new message from XML file ------
					xmlFilePath = "../../../download/xml_forms/bizpower_forms.xml"
					set objDOM = Server.CreateObject("Microsoft.XMLDOM")
					objDom.async = false			
					if objDOM.load(server.MapPath(xmlFilePath)) then
						set objNodes = objDOM.getElementsByTagName("FORM")		
						for j=0 to objNodes.length-1
							set objTask = objNodes.item(j)
							node_app_id = objTask.attributes.getNamedItem("APPEAL_ID").text										
							if trim(delappid) = trim(node_app_id) Then							
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
					
					con.executeQuery("DELETE FROM DOCUMENTS WHERE document_id IN (Select document_id From appeals_documents Where Appeal_ID = " & delappid & ")")
					con.executeQuery("DELETE FROM appeals_documents WHERE Appeal_ID = " & delappid)			
					con.ExecuteQuery("UPDATE tasks SET appeal_id = NULL WHERE appeal_id =" & delappid )
					con.ExecuteQuery("DELETE FROM Form_Value WHERE appeal_id =" & delappid )
					con.ExecuteQuery("DELETE FROM appeals WHERE appeal_id =" & delappid)
				rs_delete.moveNext
				Wend
				Set rs_delete = Nothing
				
			end if
			' ------ delete appeals --------
		   Response.Redirect(urlSort)
	End If
	
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 9 Order By word_id"				
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
	set rstitle = Nothing%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name=vs_defaultClientScript content="JavaScript">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name=ProgId content=VisualStudio.HTML>
<meta name=Originator content="Microsoft Visual Studio .NET 7.1">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
	function checkMove()
	{
			var fl = 0;
			document.form1.trapp.value = '';
			
			var NewContactId = document.getElementById("ContactId").value;
			
			if(isNaN(NewContactId) || NewContactId == '' || NewContactId == '0')
			{
				window.alert("נא לבחור לקוח להעברה");
				return false;
			}
			
			if(document.form1.ids)
			{
			var strid = new String(document.getElementById("ids").value);
			if(strid != "")
			{
				var arrid = strid.split(',');
				for (i=0;i<arrid.length;i++){
					if (document.getElementById('cb'+ arrid[i]).checked)
					{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
						fl = 1;
					}	
				}
				
				if (fl && confirm("? האם ברצונך להעביר את הטפסים המסומנים ללקוח הנבחר")){
					var txtnew = new String(document.form1.trapp.value);
					document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
					return true;
				}
				else if (fl) return false;
			}
					
			window.alert("נא לסמן לפחות טופס אחד");
			return false;
		}	
		
		return true;
	}	

	function checkChangeStatus()
	{
			var fl = 0;
			document.form1.trapp.value = '';
			if(document.form1.ids)
			{
			var strid = new String(document.getElementById("ids").value);
			if(strid != "")
			{
				var arrid = strid.split(',');
				for (i=0;i<arrid.length;i++){
					if (document.getElementById('cb'+ arrid[i]).checked)
					{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
						fl = 1;
					}	
				}
				
				if (fl && confirm("? האם ברצונך להעביר את הטפסים המסומנים לסטאטוס הנבחר")){
					var txtnew = new String(document.form1.trapp.value);
					document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
					window.document.all("change_status_flag").value = "1";
					return true;
				}
				else if (fl) return false;
			}
					
			window.alert("נא לסמן לפחות טופס אחד");
			return false;
		}	
		
		return true;
	}	
	
	function checktransf()
	{
		var fl = 0;
		document.form1.trapp.value = '';
		if(document.form1.ids)
		{
		var strid = new String(document.form1.ids.value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
				if (document.forms('form1').elements('cb'+ arrid[i]).checked)
				{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
					fl = 1;
				}	
			}
		    <%
			If trim(lang_id) = "1" Then
				str_confirm = "? האם ברצונך להעביר את הטפסים המסומנים"
			Else
				str_confirm = "Are you sure want to transfer the selected forms?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
				window.document.all("delete_flag").value = "0";
				window.document.all("add_tasks_flag").value = "0";
				window.document.all("change_status_flag").value = "0";				
				return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן הטפסים להעברה"
			Else
				str_confirm = "Please select forms to transfer !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;
	}	
	return false;	
	}	
	
	function checkDelete(){
		var fl = 0;
		document.form1.trapp.value = '';
		if(document.form1.ids)
		{
		var strid = new String(document.form1.ids.value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
				if (document.forms('form1').elements('cb'+ arrid[i]).checked)
				{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
					fl = 1;
				}	
			}
		    <%
			If trim(lang_id) = "1" Then
				str_confirm = "? האם ברצונך למחוק את הטפסים המסומנים"
			Else
				str_confirm = "Are you sure want to delete the selected forms?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
				window.document.all("delete_flag").value = "1";
				return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן הטפסים למחיקה"
			Else
				str_confirm = "Please select forms to delete !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;
	}			
		return false;	
	}
		
	function checkAddTasks()
	{
		var fl = 0;
		document.form1.trapp.value = '';
		if(document.form1.ids)
		{
		var strid = new String(document.form1.ids.value);
		if(strid != "")
		{
			var arrid = strid.split(',');
			for (i=0;i<arrid.length;i++){
				if (document.forms('form1').elements('cb'+ arrid[i]).checked)
				{	document.form1.trapp.value = document.form1.trapp.value + arrid[i] + ',';
					fl = 1;
				}	
			}
		    <%
			If trim(lang_id) = "1" Then
				str_confirm = "? האם ברצונך ליצור משימה גורפת עבור הטפסים המסומנים"
			Else
				str_confirm = "Are you sure want to add tasks for the selected forms?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
				window.document.all("add_tasks_flag").value = "1";
				h = parseInt(520);
				w = parseInt(520);
				window.open("../tasks/addtasks.asp?appealsId=" + document.form1.trapp.value, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
				return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן הטפסים על מנת ליצור משימה גורפת"
			Else
				str_confirm = "Please select forms to add to tasks !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;
	}	
	}	
	
	
	function cball_onclick() {
		var strid = new String(document.form1.ids.value);
		var arrid = strid.split(',');
		for (i=0;i<arrid.length;i++)
			document.forms('form1').elements('cb'+ arrid[i]).checked = document.form1.cb_all.checked ;	
	}	

	function openContactsList()
	{
		window.open('../appeals/contacts_list.asp?companyId=','winList','top=50, left=50, width=740, height=600, scrollbars=1');
		return false;
	}	
//-->
</script>
</head>
<body style="margin: 0px;" onload="self.focus();"  >
<%
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
	sortby_app(11) = "status_order, appeal_id DESC"
	sortby_app(12) = "status_order DESC, appeal_id DESC"
	sortby_app(13) = "Company_NAME, appeal_id DESC"
	sortby_app(14) = "Company_NAME DESC, appeal_id DESC"

	sort_app = Request("sort_app")	
	if sort_app = nil then
		sort_app = 4
	end if
	
	search = false
	If trim(Request("productID")) <> nil Then
		productID = trim(Request("productID"))
		where_product = " AND QUESTIONS_ID = " & productID
		search = true
	Else 
	    productID = ""	:    where_product = ""    :    search = false
	End If
	
	If trim(contactID) <> "" Then
		sqlstr = "Select contact_Name from contacts WHERE contact_Id = " & contactID
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			contactName =  trim(rs_pr("contact_Name"))
		End If
		set rs_pr = Nothing
	End If			

	sqlstr = "Exec dbo.get_appeals '','','','','" & OrgID & "','" & sortby_app(sort_app) & "','','','','" & contactID & "','','','" & productID & "','" & UserID & "','','','" & is_groups & "'"
    'Response.Write sqlStr
	set app=con.GetRecordSet(sqlStr)
	app_count = app.RecordCount	
	if Request("page_app")<>"" then
		page_app=request("page_app")
	else
		page_app=1
	end if
	if not app.eof then
		app.PageSize = 15
		app.AbsolutePage=page_app
		recCount=app.RecordCount 		
		NumberOfPagesApp = app.PageCount		
		i=1
		j=0
		ids = "" 'list of appeal_id
	end if
	if not app.eof Or search = true then %>
  <form id="form1" name="form1" action="contact_appeals.asp?contactID=<%=contactID%>" method="post">
  <table cellpadding="0" cellspacing="0" dir="<%=dir_var%>" width="100%" border="0" >  
  <tr><td width="100%"><A name="table_appeals"></A>  
  <table cellpadding=0 cellspacing=0 width="100%" dir="<%=dir_var%>" border="0" >
  <tr>
    <td valign="top" class="title_form" width="100%" align="<%=align_var%>" dir=<%=dir_obj_var%>>&nbsp;<!--טפסים מצורפים--><%=arrTitles(28)%>&nbsp;<font color="#E6E6E6">(<%=contactName%>)</font>&nbsp;</td>
  </tr>
  <tr>  
  <td class="title_form" nowrap>
  <table cellpadding="2" cellspacing="0" border="0" >
	<tr>
	<td align="center" ><a class="button_edit_1" style="width:90px; line-height:110%; padding:5px"  href='javascript:void(0)' onclick="if (checktransf()) {document.form1.submit()}">העברה לארכיון</a></td>
	<td width="10" nowrap valign="top"></td>
	<td align="center" ><a class="button_edit_1" style="width:80px; line-height:110%; padding:5px"  href='javascript:void(0)' onclick="if (checkDelete()) {document.form1.submit()}">מחק טפסים</a></td>
	<td width="10" nowrap ></td>
	<td align="center" ><a class="button_edit_1" style="width:104px; line-height:110%; padding:5px"  href='javascript:void(0)' onclick="return checkAddTasks();">צור משימה גורפת</a></td>
	<td width="15" nowrap></td>
	<td  nowrap ><input type="button" value="העבר לסטאטוס" class="button_edit_2" 
	style="width: 90px; display: inline;" ID="btnChangeStatus" NAME="btnChangeStatus"
	onclick="if (checkChangeStatus()) {document.form1.submit()} ">&nbsp;<select ID="cmb_change_status" 
	dir="<%=dir_obj_var%>" class="search" style="width: 75px; display: inline;" NAME="cmb_change_status">			
	<%For i=1 To Ubound(arr_Status)%>
	<option value="<%=arr_Status(i, 0)%>" <%If trim(arr_Status(i, 0)) = trim(app_status) Then%> selected <%End If%> ><%=arr_Status(i, 1)%></option>
	<%Next%></select>
	</td>
	<td width="20" nowrap></td>
	<td nowrap  align="left"><a class="button_edit_2" style="width:80px; line-height:110%; padding:5px"  href='javascript:void(0)' onclick="if (checkMove()) {document.form1.submit()}">העבר טפסים</a></td>
	<td nowrap  align="left"><input type="text" name="ContactName" id="ContactName" value="" ReadOnly class="Form_R" style="width: 200px;" dir="rtl"></td>
	<td nowrap  align="left"><input type="button" value="בחר לקוח" class="button_edit_1" style="width: 74px" onclick="return openContactsList();" ID="btnOpenPP" NAME="btnOpenPP"></td>
	</tr>
	</table> </td></tr>
  </table></td></tr>  	  	
  <tr>
	<td width="100%" align="center" valign="top">
	<table BORDER="0" CELLSPACING="1" CELLPADDING="1" bgcolor="#FFFFFF" dir="<%=dir_var%>" >	
    <tr>
	    <td width="40" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"><!--הדפס--><%=arrTitles(29)%></td>
		<td width="60" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"><%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%></td>
		<%key_table_width=400%>
		<td class="title_sort">&nbsp;</td>		
		<td width="100%" id="td_app_prod" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort" nowrap>&nbsp;<!--סוג טופס--><%=arrTitles(31)%>&nbsp;</td>				
		<td width="50" nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=3 or sort_app=4 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=4 then%>3<%elseif sort_app=3 then%>4<%else%>4<%end if%>#table_appeals" target="_self">&nbsp;ID&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="3" then%>bot<%elseif trim(sort_app)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="60" nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=1 or sort_app=2 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=2 then%>1<%elseif sort_app=1 then%>2<%else%>2<%end if%>#table_appeals" target="_self">&nbsp;<!--תאריך--><%=arrTitles(32)%>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="1" then%>bot<%elseif trim(sort_app)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>			
		<td width="50" nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort<%if sort_app=11 or sort_app=12 then%>_act<%end if%>" nowrap><a class="title_sort" HREF="<%=urlSort%>&productID=<%=productID%>&sort_app=<%if sort_app=12 then%>11<%elseif sort_app=11 then%>12<%else%>12<%end if%>#table_appeals" target="_self">&nbsp;<!--'סט--><%=arrTitles(33)%>&nbsp;<img src="../../images/arrow_<%if trim(sort_app)="11" then%>bot<%elseif trim(sort_app)="12" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="20" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"><INPUT type="checkbox" LANGUAGE="javascript" onclick="return cball_onclick()" title="<%=arrTitles(20)%>" id="cb_all" name="cb_all"></td>
	</tr>
	<%Do while (Not app.eof And j<app.PageSize)
		If j Mod 2 = 0 Then
			tr_bgcolor = "#E6E6E6"
		Else
			tr_bgcolor = "#C9C9C9"
		End If		
		appid = trim(app("appeal_id"))
		ids = ids & appid		
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
		mechanismId = app("mechanism_id")			
		User_Name = trim(app("User_Name"))
		prod_id = app("product_id")
		quest_id = trim(app("questions_id"))

		sqlstr = "EXECUTE get_appeal_tasks '" & OrgID & "','" & appid & "'"	
	    set rsmess = con.getRecordSet(sqlstr)
		if not rsmess.eof then
			mes_new = rsmess("mes_new") : mes_work = rsmess("mes_work") : mes_close = rsmess("mes_close") 
		else
			mes_new = 0 : mes_work = 0 : mes_close = 0
		end if
		If trim(SURVEYS)  = "1" Then
		    href_ = " HREF=""../appeals/appeal_card.asp?quest_id=" & quest_id & "&appid=" & appid & """  target=""_new"" "
		Else
			href_ = " nohref"
		End If 
		appeal_status = trim(app("appeal_status"))	
		appeal_status_name = trim(app("appeal_status_name"))	
		appeal_status_color = trim(app("appeal_status_color"))	%>
		<tr bgcolor="<%=tr_bgcolor%>" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='<%=tr_bgcolor%>';">
		    <td align="center" nowrap>&nbsp;<a href="javascript:void(0)" onclick="javascript:window.open('../appeals/view_appeal.asp?quest_id=<%=quest_id%>&appid=<%=appid%>','','top=100,left=100,width=500,height=500,scrollbars=1,resizable=1,menubar=1')"><IMG SRC="../../images/print_icon.gif" BORDER=0 hspace=0 vspace=0></a>&nbsp;</td>
			<td align="center" nowrap>						
			&nbsp;
			<a HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&task_status=3" target="_new" style="WIDTH:10pt;" class="task_status_num3" title="<%=arr_StatusT(3)%>"><%=mes_close%></a>
			<a HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&task_status=2" target="_new" style="WIDTH:10pt;" class="task_status_num2" title="<%=arr_StatusT(2)%>"><%=mes_work%></a>
			<a HREF="../appeals/appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&task_status=1" target="_new" style="WIDTH:10pt;" class="task_status_num1" title="<%=arr_StatusT(1)%>"><%=mes_new%></a>
			&nbsp;		
			</td>			
			<td align="<%=align_var%>">
		    <!--#include file="../appeals/key_fields_t.asp"-->
			</td>		
			<td nowrap align="<%=align_var%>"><a class="link_categ" <%=href_%> dir="<%=dir_obj_var%>">&nbsp;<%=product_name%>&nbsp;</a></td>					
			<td nowrap align="center"><a class="link_categ" <%=href_%>  dir="<%=dir_var%>"><%=appid%></a></td>
			<td align="center"><a class="link_categ" <%=href_%> >&nbsp;<%=day(app("appeal_date"))%>/<%=month(app("appeal_date"))%>/<%=mid(year(app("appeal_date")),3,2)%>&nbsp;</a></td>						
			<td nowrap align="center"><a class="status_num" style="background-color:<%=trim(appeal_status_color)%>" <%=href_%> target="_new"><%=appeal_status_name%></a></td>
			<td align="center" valign="top"><INPUT type="checkbox" id=cb<%=appid%> name=cb<%=appid%>></td>			
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
		<input type="hidden" name="trapp" value="" ID="trapp">
		<input type="hidden" name="delete_flag" value="0" ID="delete_flag">
		<input type="hidden" name="add_tasks_flag" value="0" ID="add_tasks_flag">	
		<input type="hidden" name="change_status_flag" value="0" ID="change_status_flag">		
		<input type="hidden" name="ContactId" value="0" ID="ContactId">			
		</td></tr>		
		<% if cInt(NumberOfPagesApp) > 1 then	   %>
		<tr class="card">
		<td width="100%" align="center" nowrap class="card" dir=ltr>
			<table border="0" cellspacing="0" cellpadding="2" >               
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
			 <td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=10*(numOfRowApp-1)-9%>&numOfRowApp=<%=numOfRowApp-1%>#table_appeals" name=word56 title="<%=arrTitles(56)%>" target="_self"><b><<</b></a></td>			                
			  <%end if%>
	             <td><font size="2" color="#060165">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowApp-1)) <= NumberOfPagesApp Then
	                  if CInt(page_app)=CInt(i+10*(numOfRowApp-1)) then %>
		                 <td align="center"><span class="pageCounter"><%=i+10*(numOfRowApp-1)%></span></td>
	                  <%else%>
	                     <td align="center"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=i+10*(numOfRowApp-1)%>&numOfRowApp=<%=numOfRowApp%>#table_appeals" target="_self"><%=i+10*(numOfRowApp-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td><font size="2" color="#060165">]</font></td>
				<%if NumberOfPagesApp > cint(num * numOfRowApp) then%>  
					<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&productID=<%=productID%>&sort_app=<%=sort_app%>&page_app=<%=10*(numOfRowApp) + 1%>&numOfRowApp=<%=numOfRowApp+1%>#table_appeals" name=word57 title="<%=arrTitles(57)%>" target="_self"><b>>></b></a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>
	<%End If%>	
	<%If app.recordCount = 0 Then%>
	<tr><td align="center" class="card1">&nbsp;</td></tr>									
	<%End If%>		
	</form>		 			 
	</table>
	<%Else%><p class="titlep" align="center" >לא נמצאו טפסים מצורפים לאיש הקשר</p>
	<%End If%>	
<%set app = nothing	%>
</body>
</html>
<%Set con = Nothing%>
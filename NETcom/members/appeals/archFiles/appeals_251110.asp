<%@ Language=VBScript%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="stylesheet" type="text/css">
</head>
<%app_status = trim(Request("app_status"))

	arch = trim(Request("arch"))
	if arch = "" then
		arch = 0
	end if

	if arch = "0" then
		set_arch = 1
	else
		set_arch = 0
	end if

	where_arch = " and appeal_deleted=" & arch

	if (Request("prodId") <> nil) then 'and Request.QueryString("delProd") = nil) then
		prodId = Request("prodId")
		where_product_id = " and questions_id=" & prodId
		set prod = con.GetRecordSet("Select product_name, ADD_CLIENT from products where product_id=" & prodId)
		if not prod.eof then
			title_prodName = trim(prod("product_name"))
			ADD_CLIENT = trim(prod("ADD_CLIENT"))
		end if
		set prod = nothing
	else
		sqlstr = "Select Top 1 product_id, product_name, ADD_CLIENT from Products Where "&_
		" product_number = '0' and ORGANIZATION_ID=" & OrgID & " order by product_name"
		set rs_prod = con.getRecordSet(sqlstr)
		If not rs_prod.eof Then
			prodId = rs_prod(0)
			title_prodName = rs_prod(1)
			ADD_CLIENT = trim(prod("ADD_CLIENT"))
		End If
		set rs_prod = Nothing
	end if

	if Request.QueryString("clientId") <> nil then
		clientId = Request.QueryString("clientId")
		where_client_id = " and contact_id=" & clientId
	else
		where_client_id = ""	
	end if

	if trim(Request("orgname")) <> "" then
		orgname = trim(Request("orgname")) 
		where_orgname = " and LOWER(COMPANY_NAME) LIKE '%"& LCase(trim(Request.Form("orgname"))) &"%'"		
	else
		where_orgname = ""
	end if
	
	if trim(Request("clname")) <> "" then
		clname = trim(Request("clname"))
		where_clname = " and LOWER(CONTACT_NAME) LIKE '%"& LCase(trim(Request.Form("clname"))) &"%'"		
	else
		where_clname = ""
	end if
	
	if trim(Request("usname")) <> "" then
		usname = trim(Request("usname"))
		where_usname = " and LOWER(USER_NAME) LIKE '%"& LCase(trim(Request.Form("usname"))) &"%'"		
	else
		where_usname = ""
	end if

	if trim(Request("search_id")) <> "" then
		search_id = trim(Request("search_id"))
		where_id = " and APPEAL_ID = "& trim(Request.Form("search_id"))
	else
		where_id = ""
	end if

If IsNumeric(prodId) Then
   set rs_fields=con.GetRecordSet("SELECT Field_Id,Field_Title,Field_Type FROM FORM_FIELD WHERE PRODUCT_ID=" & prodId &" And ORGANIZATION_ID = "& OrgID &" AND FORM_FIELD.Field_Key = 1 Order by Field_Order DESC")
   If Not rs_fields.EOF Then
	    arr_fields = rs_fields.getRows()
   End If
   Set rs_fields = Nothing
   if IsArray(arr_fields) Then
		count_fields = Ubound(arr_fields, 2)
   Else
		count_fields = 0
   End If
   
  urlSort = "appeals.asp?arch=" & arch & "&prodId=" & prodId & "&app_status=" & app_status & "&search_id=" & search_id & "&clname=" & Server.URLEncode(clname) & "&orgname=" & Server.URLEncode(orgname) & "&usname=" & Server.URLEncode(usname)
  
	'dynamic fields search 
	 if IsArray(arr_fields) Then
	If count_fields >= 0 Then
		str_where_values = ""				
		If Request.Form.Count > 0 Then
			For ss=0 To count_fields							
				srch_val = ""
				field_type = trim(arr_fields(2, ss))
				field_title = trim(arr_fields(1, ss))
				
				if Not isNULL( Request.Form("inpsearch" & trim(arr_fields(0, ss))) ) Then
					If Len(Request.Form("inpsearch" & trim(arr_fields(0, ss)))) > 0 Then
						srch_val = trim(Request.Form("inpsearch" & trim(arr_fields(0, ss))))
					End If	
				End If
				Execute("srch_val" & trim(arr_fields(0, ss)) & "= """ & vFix(srch_val) & """")
				urlSort = urlSort & "&srch_val" & trim(arr_fields(0, ss)) & "=" & Server.URLEncode(Eval("srch_val" & trim(arr_fields(0, ss))))
				If trim(srch_val) <> "" Then
					str_where_values = str_where_values & " AND Appeal_ID IN (SELECT Appeal_ID FROM  FORM_VALUE WHERE FIELD_ID = " & trim(arr_fields(0, ss)) 
					If trim(field_type) = "5" And trim(srch_val) = trim(field_title) Then
						str_where_values = str_where_values & " AND FIELD_VALUE = 'on' )"
					Else
						str_where_values = str_where_values & " AND FIELD_VALUE Like '%" &  sFix(srch_val) & "%')"						
					End If
				End If	
				'Response.Write(Eval("srch_val" & trim(arr_fields(0, ss))) & "<br>") 
			Next
			Else
				For ss=0 To count_fields	
					srch_val = ""
					field_type = trim(arr_fields(2, ss))
					field_title = trim(arr_fields(1, ss))

					if Not isNULL( Request.QueryString("srch_val" & trim(arr_fields(0, ss))) ) Then
						If Len(Request.QueryString("srch_val" & trim(arr_fields(0, ss)))) > 0 Then
							srch_val = trim(Request.QueryString("srch_val" & trim(arr_fields(0, ss))))
						End If	
					End If
					Execute("srch_val" & trim(arr_fields(0, ss)) & "= """ & vFix(srch_val) & """")
					urlSort = urlSort & "&srch_val" & trim(arr_fields(0, ss)) & "=" & Server.URLEncode(Eval("srch_val" & trim(arr_fields(0, ss))))
					If trim(srch_val) <> "" Then
						str_where_values = str_where_values & " AND Appeal_ID IN (SELECT Appeal_ID FROM  FORM_VALUE WHERE FIELD_ID = " & trim(arr_fields(0, ss)) 
						If trim(field_type) = "5" And trim(srch_val) = trim(field_title) Then
							str_where_values = str_where_values & " AND FIELD_VALUE = 'on' )"
						Else
							str_where_values = str_where_values & " AND FIELD_VALUE Like '%" &  sFix(srch_val) & "%')"						
						End If					
					End if	
					Next
				End if 
			End if 
		End if 
	End If
	
	dim sortby(16)	
	sortby(1) = " appeal_date, appeal_status, appeal_id DESC"
	sortby(2) = " appeal_date DESC, appeal_status, appeal_id DESC"
	sortby(3) = " appeal_id"
	sortby(4) = " appeal_id DESC"
	sortby(5) = " USERS.FIRSTNAME, USERS.LASTNAME, appeal_id DESC"
	sortby(6) = " USERS.FIRSTNAME DESC, USERS.LASTNAME DESC, appeal_id DESC"
	sortby(7) = " product_name, appeal_id DESC"
	sortby(8) = " product_name DESC, appeal_id DESC"
	sortby(9) = " CONTACT_NAME , appeal_id DESC"
	sortby(10) = "CONTACT_NAME DESC, appeal_id DESC"
	sortby(11) = " COMPANY_NAME, appeal_id DESC"
	sortby(12) = " COMPANY_NAME DESC, appeal_id DESC"

	sort = Request("sort")	
	if sort = nil then
		sort = 2
	end if

' ------ start deleting appeals ----
if Request.QueryString("delProd") <> "" then
	delappid = Request.QueryString("appId")			
	'------ start deleting the new message from XML file ------
	xmlFilePath = "../../../download/xml_forms/bizpower_forms.xml"
	set objDOM = Server.CreateObject("Microsoft.XMLDOM")
	objDom.async = false			
	if objDOM.load(server.MapPath(xmlFilePath)) then
		set objNodes = objDOM.getElementsByTagName("FORM")		
		for j=0 to objNodes.length-1
			set objTask = objNodes.item(j)
			node_app_id = objTask.attributes.getNamedItem("APPEAL_ID").text										
			if trim(Request.QueryString("appId")) = trim(node_app_id) Then							
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
	
	con.executeQuery("DELETE FROM DOCUMENTS WHERE document_id IN (Select document_id From appeals_documents Where Appeal_ID = " & delappid & ")")
	con.executeQuery("DELETE FROM appeals_documents WHERE Appeal_ID = " & delappid)	
	con.ExecuteQuery("UPDATE tasks SET appeal_id = NULL WHERE appeal_id =" & delappid )
	con.ExecuteQuery("DELETE FROM Form_Value WHERE appeal_id =" & delappid )
	con.ExecuteQuery("DELETE FROM contact_to_forms WHERE appeal_id =" & delappid )	
	con.ExecuteQuery("DELETE FROM appeals WHERE appeal_id =" & delappid)
	
	Response.Redirect "appeals.asp?prodId=" & prodId & "&arch=" & arch
end if 

' ------ end deleting appeals --------

' ------ transfering to archive --------
if trim(Request.Form("trapp")) <> "" And trim(Request.Form("delete_flag")) = "0" And trim(Request.Form("change_status_flag")) = "0" then
	sqlstr = "UPDATE appeals SET appeal_deleted=" & set_arch & " WHERE appeal_id IN (" & Request.Form("trapp") & ")"
	'Response.Write(sqlstr)
	'Response.End
	con.ExecuteQuery(sqlstr)
	Response.Redirect urlSort '"appeals.asp?prodId=" & prodId & "&arch=" & arch
end if
' ------ transfering to archive --------

'------ start changing appeals statuses --
if trim(Request.Form("trapp")) <> "" And trim(Request.Form("change_status_flag")) = "1" then
	If isNumeric(Request.Form("cmb_change_status")) And trim(Request.Form("cmb_change_status")) <> "" Then
		cmb_change_status = cInt(Request.Form("cmb_change_status"))
	Else
		cmb_change_status = 1
	End If	
	sqlstr = "UPDATE APPEALS SET APPEAL_STATUS=" & cmb_change_status & " WHERE appeal_id IN (" & Request.Form("trapp") & ")"
	'Response.Write(sqlstr)
	'Response.End
	con.ExecuteQuery(sqlstr)
	Response.Redirect urlSort '"appeals.asp?prodId=" & prodId & "&arch=" & arch
end if
'------ end changing appeals statuses ---

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
		con.ExecuteQuery("DELETE FROM contact_to_forms WHERE appeal_id =" & delappid )
		con.ExecuteQuery("DELETE FROM appeals WHERE appeal_id =" & delappid)
	rs_delete.moveNext
	Wend
	Set rs_delete = Nothing
	Response.Redirect "appeals.asp?prodId=" & prodId & "&arch=" & arch	
end if
' ------ delete appeals --------

	arr_Status = session("arr_Status")
	if lang_id = "1" then
	    arr_StatusT = Array("","חדש","בטיפול","סגור")	
    else
		arr_StatusT = Array("","new","active","close")	
    end if 	
  
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 18 Order By word_id"				
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
	set rstitle = Nothing	%> 	
<script LANGUAGE="javascript" type="text/javascript">
<!--
var oPopup = window.createPopup();

function CheckDelProd() {
  <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את המשוב"
     Else
		str_confirm = "Are you sure want to delete the form?"
     End If   
  %>
  return (confirm("<%=str_confirm%>"))    
}

function cball_onclick() {
	var strid = new String(document.form1.ids.value);
	var arrid = strid.split(',');
	for (i=0;i<arrid.length;i++)
		document.forms('form1').elements('cb'+ arrid[i]).checked = document.form1.cb_all.checked ;
	
}
function checktransf(){
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
	
	function checkChangeStatus()
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
				str_confirm = "? האם ברצונך להעביר את הטפסים המסומנים לסטאטוס הנבחר"
			Else
				str_confirm = "Are you sure want to move the selected forms to the selected status?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
				document.form1.action = "<%=urlSort%>";
				//window.alert(document.form1.action);
				window.document.all("change_status_flag").value = "1";
				return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן הטפסים "
			Else
				str_confirm = "Please select forms !"
			End If   
		  %>			
		window.alert("<%=str_confirm%>");
		return false;
	}			
		return false;	
	}

var prod;
function openPreview(prodId)
{
	prod = window.open("../products/check_form.asp?prodId="+prodId,"Product","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=700, height=400, left=10, top=20");
	if((prod.document != null) && (prod.document.body != null))
	{
		prod.document.title = "Product Preview";
		prod.document.body.style.margintop  = 20;
	}	
	return false;
}

//-->
</script>
<body>
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>" >
<tr><td width="100%">
<!-- #include file="../../logo_top.asp" -->
</td></tr>
<%numOfTab = 1%>
<%numOfLink = 0%>
<tr><td width="100%">
<!--#include file="../../top_in.asp"-->
</td></tr>
 <tr><td><table cellpadding=0 cellspacing=0 width="100%">
 <tr><td class="page_title" style="border-right:none"><a class="button_edit_1" style="width:90;"  href='appeals.asp?prodId=<%=prodId%>&arch=<%if arch = "1" then%>0<%else%>1<%end if%>'><%if arch = "0" then%><!--ארכיון--><%=arrTitles(15)%><%else%><!--חזרה לטפסים--><%=arrTitles(16)%><%end if%></a></td>
 <td class="page_title" style="border-left:none" width="100%">
 <SELECT dir="<%=dir_obj_var%>" size=1 ID="prodId" name="prodId" class="sel" style="width:310px;font-size:10pt" onChange="document.location.href='appeals.asp?arch=<%=arch%>&sort=<%=sort%>&prodID='+this.value">	
<%
	If is_groups = 0 Then
		sqlstr = "SELECT product_id, product_name FROM dbo.Products WHERE (product_number = '0') " & _
		" AND (ORGANIZATION_ID=" & OrgID & ") ORDER BY product_name"
		' משתמש אשר שייך לקבוצה אבל אינו אחראי באף קבוצה
		Else
		sqlstr = "exec dbo.get_products_list '" & OrgID & "','" & UserID & "'"
	End If
	'Response.Write sqlstr
	'Response.End
	set rs_products = con.GetRecordSet(sqlstr)
	if not rs_products.eof then 
		ResProductsList = rs_products.getRows()		
	end if
	set rs_products=nothing				
	If IsArray(ResProductsList) Then
	For i=0 To Ubound(ResProductsList,2)
		prod_Id = ResProductsList(0,i)   	
		product_name = ResProductsList(1,i)
%>
	<OPTION value="<%=prod_Id%>" <%If trim(prodId) = trim(prod_Id) Then%> selected <%End If%>> <%=product_name%> </OPTION>
<%	Next	
 End If	
%>
</SELECT>&nbsp;&nbsp;&nbsp;</td></tr>	
</table></td></tr>	    		       	       
 <tr><td valign="top">
<FORM action="appeals.asp?prodId=<%=prodId%>&sort=<%=sort%>&arch=<%=arch%>" method=POST id="form1" name="form1" target="_self">   
 <table border="0"  width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">  
  <tr>    
    <td width="100%" valign="top" align="center">
	<table width="100%" cellspacing="0" cellpadding="0" align="center" border="0" bgcolor="#ffffff">
	<tr>
	<td width="100%" align="center" valign="top">
<!-- start code --> 
<%if Request("Page")<>"" then
		Page=request("Page")
	else
		Page=1
	end if
	If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
		 PageSize = RowsInList
	Else	
    	 PageSize = 10
	End If

    sqlstr = "exec dbo.get_appeals_paging @Page=" & Page & ", @RecsPerPage=" & PageSize & ", @company_Name ='" & sFix(orgname) & _
    "', @contact_name='" & sFix(clname) & "', @project_name='" & project_name & "', @appeal_status='" & app_status & _
    "', @OrgID='" & OrgID & "', @sort='" & sortby(sort) & "', @start_date='" & start_date & "', @end_date='" & end_date  & _
    "', @company_id='" & CompanyId & "', @contact_id='" & ContactId & "', @project_id='" & ProjectId &_
    "', @appeal_id='" & search_id & "', @product_id='" & prodId & "', @UserId='" & UserID & "', @archive='" & arch & _
    "', @user_name='" & sFix(usname) & "', @IsGroups='" & is_groups & "', @str_where_values='" & sFix(str_where_values) & "'"	
    'Response.Write sqlStr
    'Response.End
	set app=con.GetRecordSet(sqlStr)	%>
	<!-- start search row -->
	<input type="hidden" name="trapp" value="" ID="trapp">
	<input type="hidden" name="delete_flag" value="0" ID="delete_flag">
	<input type="hidden" name="add_tasks_flag" value="0" ID="add_tasks_flag">	
	<input type="hidden" name="change_status_flag" value="0" ID="change_status_flag">		
	<table border="0" cellspacing="1" cellpadding="0" width="100%">		
	<tr style="background-color: #8A8A8A">
		<td dir="<%=dir_obj_var%>" align="left" colspan="2" style="padding-left: 5px;">		
		<input type="submit" value="חפש" class="but_menu" style="width: 50px" ID="Submit1" NAME="Submit1">&nbsp;&nbsp;
		<input type="button" value="הצג הכל" class="but_menu" style="width: 60px" onclick="document.location.href='appeals.asp?prodId=<%=prodId%>&sort=<%=sort%>&arch=<%=arch%>'" ID="Button1" NAME="Button1">
		</td>		
		<td width="95"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:90px;" value="<%=vFix(usname)%>" name="usname" ID="Text1" maxlength="50"></td>	
		<%If ADD_CLIENT <> "" Then 'טופס מקושר לפרטים קשרי לקוחות%>
		<td width="100" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:90px;" value="<%=vFix(clname)%>" name="clname" ID="Text2" maxlength="50"></td>
		<%End If%>
		<%If ADD_CLIENT = "2" Then 'טופס מקושר לפרטים מלאים קשרי לקוחות%>
		<td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width: 99%;" value="<%=vFix(orgname)%>" name="orgname" ID="Text3" maxlength="50"></td>
		<%End If%>
		<%If count_fields >= 0 And IsArray(arr_fields) Then%>
		<%For ff=0 To count_fields%>
		<td align="<%=align_var%>" dir="<%=dir_obj_var%>"><input id="inpsearch<%=arr_fields(0, ff)%>" name="inpsearch<%=arr_fields(0, ff)%>" value="<%=vFix(Eval("srch_val" & trim(arr_fields(0, ff))))%>" class="search" dir="<%=dir_obj_var%>" style="width: 100%" maxlength="50"></td>
		<%Next %>
		<%End If%>
		<td width="40" nowrap align="center"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:40px;" value="<%=vFix(search_id)%>" name="search_id" ID="Text4" maxlength="50"></td>
		<td colspan="2" nowrap align="center" dir="<%=dir_obj_var%>">
		<select ID="app_status" dir="<%=dir_obj_var%>" class="search" style="width: 65px"
		onchange="document.location.href='appeals.asp?prodId=<%=prodId%>&sort=<%=sort%>&arch=<%=arch%>&app_status=' + this.value;" NAME="app_status">	
		<option dir="<%=dir_obj_var%>" value="" ><!--הכל--><%=arrTitles(19)%></option>
		<%For i=1 To Ubound(arr_Status)%>
		<option value="<%=arr_Status(i, 0)%>" <%If trim(arr_Status(i, 0)) = trim(app_status) Then%> selected <%End If%> ><%=arr_Status(i, 1)%></option>
		<%Next%>    
		</select>
         </td>
		<td width="20" nowrap dir="<%=dir_obj_var%>" align="center">&nbsp;</td>
	</tr>
	<!-- end search row -->		
<%	if not app.eof then %>
	<tr>
		<td width="55" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"><!--משימות--><%=arrTitles(4)%></td>
		<!--td width="40" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"><span id="word5" name=word5>עדכון</span></td-->
		<td width="58" nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" name=word25 title="<%=arrTitles(25)%>" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self">&nbsp;<!--תאריך--><%=arrTitles(6)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>		
		<td width="95"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort=5 or sort=6 then%>_act<%end if%>"><a class="title_sort" name=word25 title="מיון" HREF="<%=urlSort%>&sort=<%if sort=6 then%>5<%elseif sort=5 then%>6<%else%>6<%end if%>" target="_self">&nbsp;<!--עובד--><%=arrTitles(7)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="5" then%>down<%elseif trim(sort)="6" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
		<%If ADD_CLIENT <> "" Then 'טופס מקושר לפרטים קשרי לקוחות%>
		<td width="100" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort=9 or sort=10 then%>_act<%end if%>"><a class="title_sort" name=word25 title="מיון" HREF="<%=urlSort%>&sort=<%if sort=10 then%>9<%elseif sort=9 then%>10<%else%>10<%end if%>" target="_self">&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="9" then%>down<%elseif trim(sort)="10" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<%End If%>
		<%If ADD_CLIENT = "2" Then 'טופס מקושר לפרטים מלאים קשרי לקוחות%>
		<td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort=11 or sort=12 then%>_act<%end if%>"><a class="title_sort" name=word25 title="<%=arrTitles(25)%>" HREF="<%=urlSort%>&sort=<%if sort=12 then%>11<%elseif sort=11 then%>12<%else%>12<%end if%>" target="_self">&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="11" then%>down<%elseif trim(sort)="12" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<%End If%>
		<!--td width="100%" id=td_prod name=td_prod align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort=7 or sort=8 then%>_act<%end if%>"><a class="title_sort" name=word25 title="<%=arrTitles(25)%>" HREF="<%=urlSort%>&sort=<%if sort=7 then%>8<%elseif sort=8 then%>7<%else%>7<%end if%>" target="_self">&nbsp;<span id="word8" name=word8><--סוג טופס--<%=arrTitles(8)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort)="7" then%>down<%elseif trim(sort)="8" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td-->
		<%If count_fields >= 0 And IsArray(arr_fields)  Then
			For i=0 To count_fields	
				Field_Id = arr_fields(0, i)
				Field_Title = trim(arr_fields(1, i))
				Field_Type = trim(arr_fields(2, i))
				If Len(Field_Title) > 22 Then
					Field_Title_S = Left(Field_Title,20) & ".."
				Else
					Field_Title_S = Field_Title
				End If	%>
		<%if Field_Type <> "10" then  'question%>	
		<td class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=middle title="<%=vFix(Field_Title)%>" width="<%=Fix(100/count)-5%>%" nowrap><%=Field_Title_S%></td>
		<%end if%>
		<%Next%>
		<%end if%>
		<td width="40" nowrap align="center" class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>"><a class="title_sort" title="<%=arrTitles(25)%>" HREF="<%=urlSort%>&sort=<%if sort=4 then%>3<%elseif sort=3 then%>4<%else%>4<%end if%>" target="_self">&nbsp;ID&nbsp;<img src="../../images/arrow_<%if trim(sort)="3" then%>down<%elseif trim(sort)="4" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
		<td width="48" nowrap  class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>" >&nbsp;<!--'סט--><%=arrTitles(9)%>&nbsp;</td>
		<td width="20" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"><%if not app.eof then%><INPUT type="checkbox" LANGUAGE="javascript" onclick="return cball_onclick()" title="<%=arrTitles(20)%>" id="cb_all" name="cb_all"><%end if%></td>
		<td width="20" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="center"></td>
	</tr>
<%recCount = app("CountRecords")
	 aa = 0
	 do while not app.eof
		If aa Mod 2 = 0 Then
			tr_bgcolor = "#E6E6E6"
		Else
			tr_bgcolor = "#C9C9C9"
		End If	
	
		appid=app("appeal_id")		
		COMPANY_NAME = app("COMPANY_NAME")
		If trim(COMPANY_NAME) = "" Or IsNull(COMPANY_NAME) Then
			If trim(lang_id) = "1" Then
			COMPANY_NAME = "אין"
			Else
			COMPANY_NAME = "No"
			End If
		End If	
		CONTACT_NAME = app("CONTACT_NAME")
		If trim(CONTACT_NAME) = "" Or IsNull(CONTACT_NAME) Then
		    If trim(lang_id) = "1" Then
			CONTACT_NAME = "אין"
			Else
			CONTACT_NAME = "No"
			End If
		End If	
		User_Name = app("User_Name")
		If trim(User_Name) = "" Or IsNull(User_Name) Then
			If trim(lang_id) = "1" Then
			User_Name = "אינטרנט"
			Else
			User_Name = "No"
			End If
		End If	
		ids = ids & appid 		
		prod_id = app("product_id")
		quest_id = trim(app("questions_id"))
		contactID = trim(app("contact_id"))	
		mes_new = trim(app("mes_new"))
		mes_work = trim(app("mes_work"))
		mes_close = trim(app("mes_close"))
		appeal_status = trim(app("appeal_status"))	
		appeal_status_name = trim(app("appeal_status_name"))	
		appeal_status_color = trim(app("appeal_status_color"))	%>
		<tr bgcolor="<%=tr_bgcolor%>" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='<%=tr_bgcolor%>';">
			<td align="center" valign="top">
			<a HREF="appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" style="WIDTH:10pt;" class="task_status_num3" title="<%=arr_StatusT(3)%>"><%=mes_close%></a>
			<a HREF="appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" style="WIDTH:10pt;" class="task_status_num2" title="<%=arr_StatusT(2)%>"><%=mes_work%></a>
			<a HREF="appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" style="WIDTH:10pt;" class="task_status_num1" title="<%=arr_StatusT(1)%>"><%=mes_new%></a>
			</td>
			<td align="center" valign="top"><a class="link_categ" HREF="appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self"><%=day(app("appeal_date"))%>/<%=month(app("appeal_date"))%>/<%=mid(year(app("appeal_date")),3,2)%></a></td>			
			<td align="<%=align_var%>" valign=top><a class="link_categ" HREF="appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=User_NAME%>&nbsp;</a></td>
			<%If ADD_CLIENT <> "" Then 'טופס מקושר לפרטים קשרי לקוחות%>
			<td align="<%=align_var%>" valign=top><a class="link_categ" HREF="appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=CONTACT_NAME%>&nbsp;</a></td>
			<%End If%>
			<%If ADD_CLIENT = "2" Then 'טופס מקושר לפרטים מלאים קשרי לקוחות%>
			<td align="<%=align_var%>" valign=top><a class="link_categ" HREF="appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" dir="<%=dir_obj_var%>">&nbsp;<%=COMPANY_NAME%>&nbsp;</a></td>
			<%End If%>
		    <!--#include file="key_fields.asp"-->		
			<td align="center" valign="top"><a class="link_categ" HREF="appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0" target="_self" dir="<%=dir_obj_var%>"><%=appid%></a></td>
            <td align="center" valign="top"><a class="status_num" style="background-color:<%=trim(appeal_status_color)%>" href="appeal_card.asp?quest_id=<%=quest_id%>&appid=<%=appid%>&allTasks=0"><%=appeal_status_name%></a></td>			
			<td align="center" valign="top"><INPUT type="checkbox" id=cb<%=appid%> name=cb<%=appid%>></td>
			<td align="center" valign="top"><%If trim(contactID) <> "" Then%><img src="../../images/forms_icon.gif" border="0" 
			style="cursor: pointer;"	alt="היסטוריית טפסים של הלקוח" 
			onclick="window.open('contact_appeals.asp?contactID=<%=contactID%>','winCA','top=20, left=20, width=900, height=500, scrollbars=1');"  ><%End If%></td>
		</tr>
<%		app.movenext
		j=j+1 : aa = aa + 1
		if not app.eof then
		ids = ids & ","
		end if
		loop  %>
   </table>		
   <input type="hidden" name="ids" value="<%=ids%>" ID="ids">
   </td></tr>
   <%	NumberOfPages = Fix((recCount / PageSize)+0.9)	%>
   <%	End If 	%>		
   <% if NumberOfPages > 1 then%>
		<tr class="card">
		<td width="100%" align="center" nowrap class="card" dir="ltr">
			<table border="0" cellspacing="0" cellpadding="2" >               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRow") <> nil Then
	               numOfRow = Request.QueryString("numOfRow")
	           Else numOfRow = 1
	           End If
	         %>
	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" name=22 title="לדפים הקודמים"><b><<</b></a></td>			                
			  <%end if%>
	             <td><font size="2" color="#060165">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="center"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="center"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=i+10*(numOfRow-1)%>&numOfRow=<%=numOfRow%>"><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td><font size="2" color="#060165">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" name=23 title="לדפים הבאים"><b>>></b></a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>				
	<%End If%>		
	<%If app.recordCount > 0 Then%>
	<tr><td align="center" height="20" class="card"><font style="color:#6E6DA6;font-weight:600"><!--נמצאו--><%=arrTitles(10)%>&nbsp;<%=recCount%>&nbsp;<!--רשומות--><%=arrTitles(11)%></font></td></tr>								
	<%Else%>	
	<tr>
		<td colspan="<%=count_fields+9%>" align="center" class="title_sort1" dir="<%=dir_obj_var%>"><!--לא נמצאו טפסים המלאים--><%=arrTitles(12)%></td>
	</tr>
<% 	end if ' if not app.eof
    set app = nothing%>				
</table>
 </td>	
<td width="100" nowrap align="<%=align_var%>" valign="top" class="td_menu">
<table cellpadding="2" cellspacing="0" width="100%" border="0">
<tr><td colspan="2" height="10" nowrap></td></tr>
<tr><td colspan="2" align="center"><a class="button_edit_1" style="width:90px; line-height:110%; padding:3px"  href='javascript:void(0)' onclick="if (checktransf()) {document.form1.submit()}"><%if arch="0" then%><!--העברה לארכיון--><%=arrTitles(17)%><%else%><!--לטיפול חוזר--><%=arrTitles(18)%><%end if%></a></td></tr>
<tr><td colspan="2" align="center"><a class="button_edit_1" style="width:90px; line-height:110%; padding:3px"  href='javascript:void(0)' onclick="if (checkDelete()) {document.form1.submit()}"><!--מחק טפסים--><%=arrTitles(26)%></a></td></tr>
<tr><td colspan="2" align="center"><a class="button_edit_1" style="width:90px; line-height:110%; padding:3px"  href='javascript:void(0)' onclick="return checkAddTasks();">צור משימה גורפת</a></td></tr>
<tr><td colspan="2" height="10" nowrap></td></tr>
<tr><td colspan="2" height="1" bgcolor="#8A8A8A"></td></tr>
<tr><td colspan="2" height="5" bgcolor="#B2B2B2"></td></tr>
<tr><td colspan="2" align="right" bgcolor="#B2B2B2"><select ID="cmb_change_status" dir="<%=dir_obj_var%>" 
class="search" style="width: 75px" NAME="cmb_change_status">			
<%For i=1 To Ubound(arr_Status)%>
<option value="<%=arr_Status(i, 0)%>" <%If trim(arr_Status(i, 0)) = trim(app_status) Then%> selected <%End If%> ><%=arr_Status(i, 1)%></option>
<%Next%>    
</select></td></tr>
<tr><td colspan="2" align="right" bgcolor="#B2B2B2"><input type="button" value="העבר לסטאטוס" class="button_edit_2"
 onclick="if (checkChangeStatus()) {document.form1.submit()} " ></td></tr>
<tr><td colspan="2" height="5" bgcolor="#B2B2B2"></td></tr>
<tr><td colspan="2" height="1" bgcolor="#8A8A8A"></td></tr>
<tr><td colspan="2" height="10" nowrap></td></tr>
</table></td></tr>
</table></form></td></tr>
</table>
</center></div>
</body>
</html>
<%set con=nothing%>
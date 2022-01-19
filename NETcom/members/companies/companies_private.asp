<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!--#include file="../../../include/title_meta_inc.asp"-->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
 	var oPopup_Type = window.createPopup();
	function company_typeDropDown(obj){
	    oPopup_Type.document.body.innerHTML = company_type_Popup.innerHTML; 
	    oPopup_Type.document.charset="windows-1255";
	    oPopup_Type.show(0-120+obj.offsetWidth, 17, 120, 148, obj);    
	}
	
	var oPopup_Status = window.createPopup();
	function StatusDropDown(obj)
	{
	    oPopup_Status.document.body.innerHTML = Status_Popup.innerHTML; 
	    oPopup_Status.show(0, 17, 60, 112, obj);    
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
//-->
</SCRIPT>
</head>

<%

if trim(Request.QueryString("company_type"))<>"" then
    company_type=Request.QueryString("company_type")
end if 

if trim(Request.QueryString("status"))<>"" then
    status=Request.QueryString("status")
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
  if trim(sort)="" then  sort = 0  end if
  
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
  end if
 'Response.Write ("workerAdmin="& workerAdmin)%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 4 Order By word_id"				
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
<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td width=100% align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width=100% align="<%=align_var%>">
  <%numOftab = 0%>
  <%numOfLink = 2%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title">&nbsp;</td></tr>		   
<%
search = false		

If trim(company_type)<>""  Then		
   if trim(company_type) <> "all" Then	
	  where_company_type = " And company_id In (Select company_id From company_to_types WHERE type_id=" & company_type & ")"	
	  search = true
   else
	   where_company_type = "" 	  
   end If  
Else 
    where_company_type = ""    
End If

If trim(status)<>""  Then		
   if trim(status) = "all" Then
	  where_status = ""	
   else	
	  where_status = " And status = " & status	
	  search = true 
   end If	   
Else 
   where_status = ""   
End If

if trim(Request.Form("search_company")) <> "" Or trim(Request.QueryString("search_company")) <> "" then
		search_company = trim(Request.Form("search_company"))
		if trim(Request.QueryString("search_company")) <> "" then
			search_company = trim(Request.QueryString("search_company"))
		end if					
		where_company = " And company_Name LIKE '%"& sFix(search_company) &"%'"
		search = true			
else
		where_company = ""			
end if


if trim(Request.Form("search_contact")) <> "" Or trim(Request.QueryString("search_contact")) <> "" then
		search_contact = trim(Request.Form("search_contact"))
		if trim(Request.QueryString("search_contact")) <> "" then
			search_contact = trim(Request.QueryString("search_contact"))
		end if					
		where_contact = " And CONTACT_NAME LIKE '%"& sFix(search_contact) &"%'"					
		search = true
else
		where_contact = ""		
end if

urlSort="companies_private.asp?search_company="& Server.URLEncode(search_company) &"&search_contact="& Server.URLEncode(search_contact) &"&company_type="& company_type & "&status=" & status

dim sortby(8)	
sortby(0) = "rtrim(ltrim(company_name))"
sortby(1) = "rtrim(ltrim(company_name))"
sortby(2) = "rtrim(ltrim(company_name)) DESC"
sortby(3) = "rtrim(ltrim(city_Name))"
sortby(4) = "rtrim(ltrim(city_Name)) DESC"
sortby(5) = "CONTACT_NAME"
sortby(6) = "CONTACT_NAME DESC"
sortby(7) = "type_name"
sortby(8) = "type_name DESC"
%>
<tr>    
    <td width="100%" valign="top" align="center">
    <table dir="<%=dir_var%>" border="0" cellpadding="0" cellspacing="1" width="100%">
    <tr>
    <td align="left" width="100%" valign=top dir="<%=dir_var%>" >
    <table width="100%" cellspacing="1" cellpadding="1" bgcolor=#FFFFFF>       
    <tr> 	      
	    <td id=td_group name=td_group width="120" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort"><span id=word3 name=word3><!--קבוצה--><%=arrTitles(3)%></span>&nbsp;<IMG name=word18 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(18)%>" align=absmiddle onclick="return false" onmousedown="company_typeDropDown(td_group)"></td>
	    <td width="170" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;<span id="word4" name=word4><!--אימייל--><%=arrTitles(4)%></span>&nbsp;</td>
	    <td width="75"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;<span id="word5" name=word5><!--פקס--><%=arrTitles(5)%></span></td>
  	    <td width="75"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;<span id="word6" name=word6><!--טלפון--><%=arrTitles(6)%></span></td>	    
  	    <td width="100"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word18 title="<%=arrTitles(18)%>"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word19 title="<%=arrTitles(19)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=3" name=word19 title="<%=arrTitles(19)%>"><%end if%><span id="word7" name=word7><!--עיר--><%=arrTitles(7)%></span>&nbsp;<img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
        <td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word19 title="<%=arrTitles(19)%>"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word20 title="<%=arrTitles(20)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=1" name=word20 title="<%=arrTitles(20)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	     	              
        <td id=td_status name=td_status width="50" dir="<%=dir_obj_var%>" align="<%=align_var%>" class="title_sort" nowrap>&nbsp;<span id="word8" name=word8><!--סטטוס--><%=arrTitles(8)%></span>&nbsp;<IMG name=word18 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(18)%>" align=absmiddle onmousedown="StatusDropDown(td_status)"></td>
	</tr>   
<%    
   If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
	    PageSize = RowsInList
   Else	
     	PageSize = 10
   End If
   
   if trim(lang_id) = "1" then
	   arr_status = Array("","עתידי","פעיל","סגור","פונה")
   else
	   arr_status = Array("","new","active","close","appeal")
   end if	   
   
   'sqlSelect="SELECT contact_ID, company_ID,company_name,prefix_phone,prefix_fax,"&_
   '"phone,fax,email,city_Name,status FROM companies_private_view  "
   'sqlSelect=sqlSelect & " WHERE ORGANIZATION_ID = " & trim(OrgID) & " AND private = '1' "    
   'sqlSelect=sqlSelect & where_company_type & where_company & where_contact & where_status
   'sqlSelect=sqlSelect & " order by "& sortby(sort)   
   sqlSelect = "get_companies_private " & Page & "," & PageSize & ",'" & OrgID & "','" & company_type & "','" & sFix(search_company) & "','" & sFix(search_contact) & "','" & status & "','" & sortby(sort) & "'"
   'Response.Write (sqlSelect) 
   'Response.End
   set companyList=con.GetRecordSet(sqlSelect) 
   'companyList.open sqlSelect,con,3,1
   If not companyList.EOF then		
		recCount = companyList("CountRecords")   
		do while not companyList.EOF
			contactID=companyList(1)
			companyID=companyList(2)	  
			If recCount = 1 And search = true Then
				Response.Redirect "company_private.asp?companyID=" & companyID & "&contactID=" & contactID
			End If
			companyName=rtrim(ltrim(companyList(3)))    
			email=companyList(6)
			cityName=companyList(7)
			status_company=trim(companyList(8))
			phone=companyList(4)
			If trim(phone) = "-" Then
				phone = ""
			End if	
			fax=companyList(5)
			If trim(fax) = "-" Then
				fax = ""
			End if
			If isNumeric(trim(companyID)) Then
					types=""
					sqlstr="Select type_Name From company_to_types_view Where company_id = " & companyID & " Order By type_Name"
					set rssub = con.getRecordSet(sqlstr)		   
						
					While not rssub.eof
						types = types & rssub("type_Name") & ","
						rssub.moveNext
					Wend		    
					set rssub=Nothing
					If Len(types) > 0 Then
						types = Left(types,(Len(types)-1))
					End If
					If Len(types) > 20 Then
						short_types = Left(types, 18) &  ".."
					Else short_types=types
					End If
			End If   
       %>
        <tr>    	      	     
	      <td class="card" valign=top dir="<%=dir_obj_var%>" align="<%=align_var%>" title="<%=vFix(types)%>"><a class="link_categ" href="company_private.asp?companyId=<%=companyId%>&contactID=<%=contactID%>">&nbsp;<%=short_types%>&nbsp;</a></td>	      
	      <td class="card" valign=top dir="<%=dir_obj_var%>" align="<%=align_var%>">&nbsp;<%If Len(email) > 0 Then%><a class="card" href="mailto:<%=email%>"><%=email%></a><%End If%>&nbsp;</td>
	      <td class="card" valign=top dir="<%=dir_obj_var%>" align="center"><a class="link_categ" href="company_private.asp?companyId=<%=companyId%>&contactID=<%=contactID%>" style="font-size:11px"><%=fax%></a></td>
	      <td class="card" valign=top dir="<%=dir_obj_var%>" align="center"><a class="link_categ" href="company_private.asp?companyId=<%=companyId%>&contactID=<%=contactID%>" style="font-size:11px"><%=phone%></a></td>	     
	      <td class="card" valign=top dir="<%=dir_obj_var%>" align="<%=align_var%>"><a class="link_categ" href="company_private.asp?companyId=<%=companyId%>&contactID=<%=contactID%>">&nbsp;<%=cityName%>&nbsp;</a></td>
          <td class="card" valign=top dir="<%=dir_obj_var%>" align="<%=align_var%>" dir=rtl><a class="link_categ" href="company_private.asp?companyId=<%=companyId%>&contactID=<%=contactID%>" style="line-height:120%;padding-top:3px;padding-bottom:3px">&nbsp;<%=companyName%>&nbsp;</a></td>	                  
          <td class="card" valign=top dir="<%=dir_obj_var%>" align="center"><a class=status_num<%=status_company%> href="company_private.asp?companyId=<%=companyId%>&contactID=<%=contactID%>"><%=arr_status(status_company)%></a></td>
	    <%'end if%>     
      </tr> 
	  <%companyList.movenext
	  i=i+1
	  loop
	  
	  NumberOfPages = Fix((recCount / PageSize)+0.9)
	  if NumberOfPages > 1 then
	  urlSort = urlSort & "&sort=" & sort
	  %>
	  <tr>
		<td width="100%" align=middle colspan="8" nowrap  bgcolor="#e6e6e6" dir=ltr>
			<table border="0" cellspacing="0" cellpadding="2">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>
	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word21 title="<%=arrTitles(21)%>" href="<%=urlSort%>&page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word22 title="<%=arrTitles(22)%>" href="<%=urlSort%>&page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>			
	<%companyList.close 
	set companyList=Nothing
	End if
	%>
	<tr>
	   <td colspan="8" height=18 bgcolor="#e6e6e6" align=center class="card" dir="ltr" style="color:#6F6DA6;font-weight:600"><span id=word9 name=word9><!--נמצאו--><%=arrTitles(9)%></span>&nbsp;<%=recCount%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesMulti"))%></td>
	</tr>
	<% 
	Else %>
	   <tr>
	   <td colspan="8" align=center class="title_sort1" dir="<%=dir_var%>"><span id="word10" name=word10><!--לא נמצאו--><%=arrTitles(10)%></span>&nbsp;&nbsp;</td>
	   </tr>
<% End If%>
</table></td>
<td width=100 nowrap valign=top class="td_menu" style="border: 1px solid #808080">
<table cellpadding=1 cellspacing=0 width=100%>
<tr><td align="<%=align_var%>" colspan=2 class="title_search"><span id=word11 name=word11><!--חיפוש--><%=arrTitles(11)%></span></td></tr>
<FORM action="companies_private.asp?sort=<%=sort%>" method=POST id=form_search name=form_search target="_self">   
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td></tr>
<tr>
<td align="<%=align_var%>"><input type="image" onclick="form_search.submit();" src="../../images/search.gif"></td>
<td align="<%=align_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:80;" value="<%=vFix(search_company)%>" name="search_company" ID="Text1"></td></tr>
</FORM>
<tr><td colspan=2 height=10 nowrap></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:98;" href="#" onclick="javascript:window.open('editcompany_private.asp','','top=50,left=100,resizable=0,width=460,height=500');"><span id=word12 name=word12><!--הוסף--><%=arrTitles(12)%></span> <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></a></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:95;line-height:110%;padding:3px" href="#" onclick="javascript:window.open('upload.asp','Upload','top=50,left=40,resizable=0,width=600,height=450');"><span id="word23" name=word23><!--ייבוא--><%=arrTitles(23)%></span> <%=trim(Request.Cookies("bizpegasus")("CompaniesMulti"))%></a></td></tr>
<tr><td colspan=2 height=10 nowrap></td></tr>
<%If trim(SURVEYS)  = "1" Then%>
<%
	If is_groups = 0 Then
	sqlstr = "Select product_id, product_name from Products Where isNULL(FORM_CLIENT,0) = 1 And "&_
	" product_number = '0' and ORGANIZATION_ID=" & OrgID & " order by product_name"
	Else
	sqlstr = "Select DISTINCT Products.product_id, Products.product_name from Products Inner Join Users_To_Products "&_
	" ON Products.Product_ID = Users_To_Products.Product_ID WHERE Users_To_Products.User_ID = " & UserID &_
	" And isNULL(FORM_CLIENT,0)=1 And Products.product_number = '0' and Products.ORGANIZATION_ID=" & OrgID & " order by Products.product_name"
	End If
	'Response.Write sqlstr
	'Response.End
	set rs_products = con.GetRecordSet(sqlstr)
	if not rs_products.eof then 
		ClientProductsList = rs_products.getString(,,",",";")		   
		arr_products = Split(ClientProductsList,";")
	end if
	set rs_products=nothing	
	If IsArray(arr_products) Then
%>
<tr><td align="center" colspan=2 class="title_form"><span id=word14 name=word14><!--צרף טופס--><%=arrTitles(14)%></span></td></tr>
<%    
	For i=0 To Ubound(arr_products)-1	
	arr_prod = Split(arr_products(i),",")
	If IsArray(arr_prod) Then
    quest_id = arr_prod(0)   	
    product_name = arr_prod(1)
%>
<tr><td align="<%=align_var%>" colspan=2><a class="link1" href="#" style="width:106;line-height:110%;padding:4px;direction:rtl;font-weight:500" onclick="javascript:window.open('../appeals/edit_appeal.asp?quest_id=<%=quest_id%>&from=comp_list','','top=100,left=100,width=500,height=500,scrollbars=1,resizable=1')" ><%=product_name%></a></td></tr>
<%  End If
	Next	
	End If
	set rs_products=nothing
End If%>	
</table></td></tr>
</table></td></tr></table>
<DIV ID="company_type_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute;overflow:auto; top:0; left:0; width:120; height:148; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; SCROLLBAR-FACE-COLOR:#D3D3D3;SCROLLBAR-HIGHLIGHT-COLOR: #D3D3D3; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%set company_typeList=con.GetRecordSet("SELECT type_ID,type_Name FROM company_type  WHERE Organization_Id = " & trim(Request.Cookies("bizpegasus")("OrgID")) & " ORDER BY type_Name")
do while not company_typeList.EOF
   prcompany_typeID=company_typeList(0)
   prcompany_typeName=company_typeList(1)%>
   <DIV onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; border:1px solid black; padding-right:3px;  padding-left:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='companies_private.asp?sort=<%=sort%>&company_type=<%=prcompany_typeID%>'">
    <%=prcompany_typeName%>
    </DIV>              
<%company_typeList.MoveNext
loop
company_typeList.close
set company_typeList=Nothing%>
    <DIV onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; border:1px solid black; padding-right:3px;  padding-left:3px;cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='companies_private.asp?sort=<%=sort%>'">
   כל הקבוצות
    </DIV>
</DIV> 
</div>
<DIV ID="Status_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; width:60; border-top:1px solid black;">
<%For i=1 To Ubound(arr_status)%>
	<DIV onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='companies_private.asp?sort=<%=sort%>&status=<%=i%>'">
    <%=arr_status(i)%>
    </DIV>
<%Next%>       
    <DIV onmouseover="this.style.background='#6F6DA6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='companies_private.asp?sort=<%=sort%>'">
    <span id=word15 name=word15><!--הכל--><%=arrTitles(15)%></span>
    </DIV>
</div>
</DIV>
</body>
<%set con=Nothing%>
</html>

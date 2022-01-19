<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
	appId = trim(Request("appid"))
	quest_id = trim(Request("quest_id"))
	companyID = trim(Request("companyID"))
	contactID = trim(Request("contactID"))
	projectID = trim(Request("projectID"))
	mechanismID = trim(Request("mechanismID"))
	
	PathCalImage = "../../"	
	
	arr_Status = session("arr_Status")
	
	userName = user_name 'cokie
	
	If IsNumeric(appid) Then		
    'Response.Write("sqlstr=" & sqlstr)
    sqlstr = "exec dbo.get_appeals '','','','','" & OrgID & "','','','','','','','" & appID & "'"
	set app = con.GetRecordSet(sqlstr)
	if not app.eof then
	
	appeal_date = FormatDateTime(app("appeal_date"), 2) & " " & FormatDateTime(app("appeal_date"), 4)
	quest_id = trim(app("questions_id"))	
	companyID = app("company_ID")
	contactID = app("contact_ID")
	projectID = app("project_ID")
	mechanismID = app("mechanism_ID")
	companyName = app("company_Name")
	contactName = app("contact_Name")
	projectName = app("project_Name")
	userName = app("user_Name")	
	appeal_status = app("appeal_status")
	private_flag = app("private_flag")
	UserIdOrderOwner = trim(app("User_Id_Order_Owner"))
		AppCountryID = trim(app("appeal_CountryId"))

	
	
	If trim(appeal_status) = "1" And trim(UserID) = trim(RESPONSIBLE) Then
		sqlstring="UPDATE appeals set appeal_status = '2' WHERE ORGANIZATION_ID = " & OrgID & " AND appeal_id = " & appid 	
		'Response.Write(sqlstring)
		'Response.End 
		con.ExecuteQuery(sqlstring) 
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
	End If
	Else
		appeal_date = FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4)
	End If	

	title = false
	If trim(companyId) <> "" Then
		sqlstr = "Select company_Name from companies WHERE company_Id = " & companyId
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			companyNameT = rs_pr("company_Name")
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(contactID) <> "" Then
		sqlstr = "Select contact_Name from contacts WHERE contact_Id = " & contactID
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			contactNameT = " - " & rs_pr("contact_Name")
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(projectID) <> "" Then
		sqlstr = "Select project_Name from projects WHERE project_Id = " & projectID
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			projectNameT = " - " & rs_pr("project_Name")
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(quest_id) <> "" Then
		sqlstr =  "Select Langu, PRODUCT_DESCRIPTION,RESPONSIBLE, ADD_CLIENT, FILE_ATTACHMENT, ATTACHMENT_TITLE From Products WHERE PRODUCT_ID = " & quest_id &_
		" And ORGANIZATION_ID = " & OrgID
		'Response.Write sqlstr
		'Response.End
		set rsq = con.getRecordSet(sqlstr)
		If not rsq.eof Then		
			Langu = trim(rsq(0))
			PRODUCT_DESCRIPTION = trim(rsq(1))
			RESPONSIBLE = trim(rsq(2))
			addClient = trim(rsq(3))
			attachment_file = trim(rsq(4))
			attachment_title = trim(rsq(5))			
			found_quest = true
		Else
			found_quest = false	
		End If
		set rsq = Nothing	
		if Langu = "eng" then
			dir_align = "ltr" :  td_align = "left"  :  	pr_language = "eng"
		else
			dir_align = "rtl"  :  td_align = "right"  :  pr_language = "heb"
		end if			
	End If
	
	If trim(appID) = "" Then private_flag = "0"
	compId = trim(Request.QueryString("companyID"))
	contID = trim(Request.QueryString("contactID"))
	
	'איש קשר פרטי
	If trim(compID) <> "" And trim(contID) <> "" And trim(addClient) = "1" Then
		sqlSelect="SELECT contacts.contact_id,contacts.company_id,contacts.contact_name,company_name," &_
		" contacts.phone,contacts.cellular,contacts.email FROM contacts "&_
		" Inner Join Companies On contacts.company_id = companies.company_id "&_
		" WHERE contacts.ORGANIZATION_ID = " & trim(OrgID) & " AND isNULL(private_flag,0) = '1' " &_
	    " And contacts.company_id = " & compId & " AND contacts.contact_id = " & contID
	   	set companyList = con.getRecordSet(sqlSelect)
	   	If not companyList.eof Then		
			contactID=companyList(0)
			companyID=companyList(1)
			contactName=companyList(2)
			companyName=rtrim(ltrim(companyList(3)))    						
			contactPhone = companyList(4)						
			contactCellular = companyList(5)		
			contactEmail=companyList(6)						
			set companyList = Nothing
		End If
	ElseIf trim(compID) <> "" And trim(addClient) = "2" Then
	    If trim(compID)<>"" then   
		sqlStr = "SELECT company_name,company_id,address,city_Name, phone1,phone2,zip_code,"&_
		" fax1,fax2,url,email,date_update, status,company_desc,private_flag FROM companies WHERE company_id="& compID
		set rs_company=con.GetRecordSet(sqlStr)
		if not rs_company.EOF then	
			company_name  =rs_company("company_name")
  			companyID	  =rs_company("company_id")
			address	      =rs_company("address")
			zip_code	  =rs_company("zip_code")					
			cityName      =rs_company("city_Name")				
			phone1	      =rs_company("phone1")
			phone2	      =rs_company("phone2")
			fax1          =rs_company("fax1")
			fax2	      =rs_company("fax2")
			company_email =rs_company("email")
			company_desc  =rs_company("company_desc")
			url	          =rs_company("url")				
			date_update	  =rs_company("date_update")    
			status        =rs_company("status")
			private_flag  =rs_company("private_flag")				
		end if
		set rs_company = Nothing			
		End If  
		
		If trim(contID) <> "" Then
			sql="SELECT contact_ID,email,prefix_phone,phone,fax, cellular, CONTACT_NAME,"&_
			" messanger_name,date_update,contact_city_name, contact_address,"&_
			" contact_zip_code,contact_desc, responsible_id FROM contacts WHERE contact_ID=" & contID
			set listContact=con.GetRecordSet(sql)
			if not listContact.EOF then 
				contactID=listContact("contact_ID")      
				contact_name=listContact("CONTACT_NAME")     
				contact_email=listContact("email")				
				contact_phone=listContact("phone")				
				contact_cellular=listContact("cellular")				
				contact_fax=listContact("fax")       
				messangerName=listContact("messanger_name")
				contact_address = listContact("contact_address")
				contact_city_Name = listContact("contact_city_Name")
				contact_zip_code = listContact("contact_zip_code")
				contact_desc = listContact("contact_desc")
				responsible_id = listContact("responsible_id")			        
			End if   
			set listContact = Nothing
			If isNumeric(trim(contactId)) Then
				types=""
				sqlstr = "get_contact_types '" & contactId & "'"
				set rstype = con.getRecordSet(sqlstr)		   						
				If not rstype.eof Then
					types = rstype.getString(,,",",",")						
				End If		    
				set rstype=Nothing		
				If Len(types) > 0 Then
					types = Left(types,(Len(types)-1))
				End If
				
				contact_types=""
				sqlstr = "get_contact_types_id '" & contactId & "'"
				set rstype = con.getRecordSet(sqlstr)		   						
				If not rstype.eof Then
					contact_types = rstype.getString(,,",",",")						
				End If		    
				set rstype=Nothing		
				If Len(contact_types) > 0 Then
					contact_types = Left(contact_types,(Len(contact_types)-1))
				End If		
			End If	   			
		End If		
	End If		
  
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 19 Order By word_id"				
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
	set rsbuttons=nothing %> 	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">						
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
//start of pop-ups related functions//
	function openCompaniesList()
	{
		window.open('../appeals/companies_list.asp','winList','top=50, left=50, width=740, height=600, scrollbars=1');
		return false;
	}	
	
	function removeCompany()
	{
		window.document.getElementById("companyID").value = "";
		window.document.getElementById("CompanyName").value = "";
		
		window.document.getElementById("project_id").value = "";
		window.document.getElementById("projectName").value = "";
		
		window.document.getElementById("contactID").value = "";
		window.document.getElementById("contactName").value = "";		
		
		window.document.getElementById("contacter_body").style.display = "none";
		window.document.getElementById("project_body").style.display = "none";	 
		return false;
	}
	
	function openProjectsList()
	{
		var companyId = document.getElementById("companyID").value;
		window.open('../appeals/projects_list.asp?companyId=' + companyId,'winList','top=50, left=50, width=740, height=600, scrollbars=1');
		return false;
	}	
	
	function chooseProject(ProjectId)
	{
			var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");
				
			xmlHTTP.open("POST", '../tasks/getmechanism.asp', false);
			xmlHTTP.send('<?xml version="1.0" encoding="UTF-8"?><request><projectID>' + ProjectId + '</projectID></request>');	 
				
			result = new String(xmlHTTP.responseText);						
			arr_mechanism = result.split(";");					
				
			if(window.document.getElementById("mechanism_id").options)
			{
				window.document.getElementById("mechanism_id").options.length = 1;
			}
						
			for (i=0; i < arr_mechanism.length-1; i++)
			{			
				arr_mechanismer = new String(arr_mechanism[i]);
				arr_mechanismer = arr_mechanismer.split(",");						
				window.document.getElementById("mechanism_id").options[window.document.getElementById("mechanism_id").options.length]=new Option(arr_mechanismer[1],arr_mechanismer[0]); 
			}			
	}
	
	function removeProject()
	{		
		window.document.getElementById("project_id").value = "";
		window.document.getElementById("projectName").value = "";
		
		if(window.document.getElementById("mechanism_id").options)
		{
			window.document.getElementById("mechanism_id").options.length = 1;
		}		
		window.document.getElementById("mechanism_body").style.display = "none";		
		
		return false;
	}	
	
	function openContactsList()
	{
		var companyId = document.getElementById("companyID").value;
		window.open('../appeals/contacts_list.asp?companyId=' + companyId,'winList','top=50, left=50, width=740, height=600, scrollbars=1');
		return false;
	}	
	
	function removeContact()
	{				
		window.document.getElementById("contactID").value = "";
		window.document.getElementById("contactName").value = "";		
		return false;
	}		
//end of pop-ups related functions//

function CheckFields(action)
{
	if(action == "send")
	{
		window.document.form1.send_task.value = 1;
	}	
		<%If trim(appID) = "" Then%>
		<%If trim(COMPANIES) = "1" Then%>
		<%If (trim(addClient) = "1" Or trim(addClient) = "2") And trim(companyID) = "" Then%>
		if (window.document.all("contact_name").value=='')
		{
		   <%
				If trim(lang_id) = "1" Then
					str_alert = "נא למלא שם"
				Else
					str_alert = "Please, insert the full name"
				End If	
			%>
 			window.alert("<%=str_alert%>");
 			window.document.all("contact_name").focus();
			return false;
		}
		if (window.document.all("email").value=='')
		{
		}
		else if (!checkEmail(document.all("email").value))
		{
			<%
				If trim(lang_id) = "1" Then
					str_alert = "כתובת דואר אלקטרוני לא חוקית"
				Else
					str_alert = "The email address is not valid!"
				End If	
			%>
				window.alert("<%=str_alert%>");
				document.all("email").focus();
				return false;
		} 
		<%End If%>	
		<%If trim(addClient) = "2" Then%>
	   if(window.document.all("companyID").value == "")
       {         
			if (window.document.all("company_name").value=='')
			{
			<%
					If trim(lang_id) = "1" Then
						str_confirm = "" & Space(22) & "שים לב, לא הכנסת שם חברה \n\n ?האם ברצונך להוסיף איש קשר ללא חברה"
					Else
						str_confirm = "Are you sure want to add customer without company?"
					End If	
				%>
 				var answer = window.confirm("<%=str_confirm%>");
 				if(answer == false)
 				{
 					window.document.all("company_name").focus();
					return false;
				}	
			}
		}		
		<%End If%>
		<%End If%>
		<%End If%>		
		if(click_fun())
			document.form1.submit();
				
	return false;		
}		

function checkEmail(addr)
{
	addr = addr.replace(/^\s*|\s*$/g,"");
	if (addr == '') {
	return false;
	}
	var invalidChars = '\/\'\\ ";:?!()[]\{\}^|';
	for (i=0; i<invalidChars.length; i++) {
	if (addr.indexOf(invalidChars.charAt(i),0) > -1) {
		return false;
	}
	}
	for (i=0; i<addr.length; i++) {
	if (addr.charCodeAt(i)>127) {     
		return false;
	}
	}

	var atPos = addr.indexOf('@',0);
	if (atPos == -1) {
	return false;
	}
	if (atPos == 0) {
	return false;
	}
	if (addr.indexOf('@', atPos + 1) > - 1) {
	return false;
	}
	if (addr.indexOf('.', atPos) == -1) {
	return false;
	}
	if (addr.indexOf('@.',0) != -1) {
	return false;
	}
	if (addr.indexOf('.@',0) != -1){
	return false;
	}
	if (addr.indexOf('..',0) != -1) {
	return false;
	}
	var suffix = addr.substring(addr.lastIndexOf('.')+1);
	if (suffix.length != 2 && suffix != 'com' && suffix != 'net' && suffix != 'org' && suffix != 'edu' && suffix != 'int' && suffix != 'mil' && suffix != 'gov' & suffix != 'arpa' && suffix != 'biz' && suffix != 'aero' && suffix != 'name' && suffix != 'coop' && suffix != 'info' && suffix != 'pro' && suffix != 'museum') {
	return false;
	}
return true;
}

function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=157pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		 return false;  		
}

function refreshCompany(companyObj,quest_id)
{
	if ((event.keyCode==13) || (event.keyCode == 9))
		return false;
	compName = new String(companyObj.value);
	if(compName.length > 0)
	{
		window.document.all("frameCompanies").src = "companies.asp?comp_name=" + compName + "&quest_id=" + quest_id;	
		window.document.all("frameCompanies").style.display = "inline";		
	}
	else
	{
		window.document.all("frameCompanies").style.display = "none";
	}	
	window.focus();
	companyObj.focus();	
	return true;	
			
}

function refreshContacts(contactObj,quest_id)
{
	if ((event.keyCode==13) || (event.keyCode == 9))
		return false;
	contName = new String(contactObj.value);
	
	if(contName.length > 0)
	{
		window.document.all("framecontacts").src = "contacts.asp?cont_name=" + contName + "&quest_id=" + quest_id;	
		window.document.all("framecontacts").style.display = "inline";
	}
	else
		window.document.all("framecontacts").style.display = "none";
			
}

function refreshContactsPr(contactObj,quest_id)
{
	if ((event.keyCode==13) || (event.keyCode == 9))
		return false;
	contName = new String(contactObj.value);
	if(contName.length > 0)
	{
		window.document.all("frameContactsPr").src = "contacts_private.asp?cont_name=" + contName + "&quest_id=" + quest_id;	
		window.document.all("frameContactsPr").style.display = "inline";
	}
	else
		window.document.all("frameContactsPr").style.display = "none";
			
}

function ClearCompany()
{
	if(window.document.all("companyId") != null)
	{
		window.document.all("companyId").value = "";		
	}	
	
	if(window.document.all("company_Id") != null)
	{
		window.document.all("company_Id").value = "";		
	}	
	
		
	if(window.document.all("COMPANY_NAME") != null)	
	{
		window.document.all("COMPANY_NAME").value = "";
		window.document.all("COMPANY_NAME").readOnly = false;
		window.document.all("COMPANY_NAME").className = "Form";
	}	
		
	if(window.document.all("cityName") != null)
	{		
		window.document.all("cityName").value = "";		
		window.document.all("cityName").readOnly = false;
		window.document.all("cityName").className = "Form";
	}	
		
	if(window.document.all("address") != null)
	{
		window.document.all("address").value = "";	
		window.document.all("address").readOnly = false;
		window.document.all("address").className = "Form";
	}			


	if(window.document.all("phone1") != null)
	{
		window.document.all("phone1").value = "";			
		window.document.all("phone1").readOnly = false;
		window.document.all("phone1").className = "Form";
	}			

		
	if(window.document.all("phone2") != null)
	{
		window.document.all("phone2").value = "";				
		window.document.all("phone2").readOnly = false;
		window.document.all("phone2").className = "Form";
	}
		
	if(window.document.all("fax1") != null)
	{
		window.document.all("fax1").value = "";			
		window.document.all("fax1").readOnly = false;
		window.document.all("fax1").className = "Form";
	}		
	
	if(window.document.all("company_email") != null)	
	{
		window.document.all("company_email").value = "";
		window.document.all("company_email").readOnly = false;
		window.document.all("company_email").className = "Form";
	}	
		
	if(window.document.all("url") != null)	
	{
		window.document.all("url").value = "";
		window.document.all("url").readOnly = false;
		window.document.all("url").className = "Form";
	}
	
	if(window.document.all("company_desc") != null)	
	{
		window.document.all("company_desc").value = "";
		window.document.all("company_desc").style.height = "85px";
		window.document.all("company_desc").style.lineHeight = "16px";
		window.document.all("company_desc").readOnly = false;
		window.document.all("company_desc").className = "Form";
	}
	
	if(window.document.all("cities_list") != null)	
	{
		window.document.all("cities_list").disabled = false;
		window.document.all("cities_list").style.display = "inline";
		window.document.all("cities_list").style.visible = true;
	}
	return false;
}
//-->
</script>
</head>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 1%>
<%numOfLink = 2%>
<!--#include file="../../top_in.asp"-->
<div align="center">
<form id="form1" name="form1" action="addappeal.asp?from=appeal" method="POST" style="margin: 0px;" enctype="multipart/form-data" >
<table border="0" width="100%" cellspacing="0" cellpadding="0" bgcolor=#e6e6e6 dir="<%=dir_var%>">
<tr><td class="page_title" dir="<%=dir_obj_var%>">
<input type=hidden name=appid id=appid value="<%=appid%>">
<input type=hidden name=send_task id="send_task" value="0">
&nbsp;&nbsp;
<%If title Then%>
<font color="#4D4A91">(<%=companyNameT & contactNameT & projectNameT%>)</font>
<%End If%>
&nbsp;&nbsp;
<SELECT dir="<%=dir_obj_var%>" size=1 ID=quest_id name=quest_id class="sel" style="width:310px;font-size:10pt" <%If IsNumeric(appid) Then%> disabled <%Else%> onChange="document.location.href='appeal.asp?quest_id='+this.value" <%End If%>>	
<%If trim(lang_id) = "1" Then%>
 <OPTION value="">בחר טופס</OPTION>
<%Else%>
 <OPTION value="">Choose form</OPTION>
<%End If%> 
<%
	If is_groups = 0 Then
	sqlstr = "SELECT product_id, product_name FROM PRODUCTS WHERE (product_number = '0') " &_
	" AND (SHOW_INSITE = 1) AND (ORGANIZATION_ID=" & OrgID & ") ORDER BY product_name"
	Else
	'sqlstr = "Select DISTINCT Products.product_id, Products.product_name from Products Inner Join Users_To_Products "&_
	'" ON Products.Product_ID = Users_To_Products.Product_ID WHERE Users_To_Products.User_ID = " & UserID &_
	'" And Products.product_number = '0' and Products.ORGANIZATION_ID=" & OrgID & " order by Products.product_name"
	sqlstr = "exec dbo.get_products_list_enter '" & OrgID & "','" & UserID & "'"
	End If
	'Response.Write sqlstr
	'Response.End
	set rs_products = con.GetRecordSet(sqlstr)
	if not rs_products.eof then 
		arr_products = rs_products.getRows()
	end if
	set rs_products=nothing	
	
	If IsArray(arr_products) Then
		For i=0 To  Ubound(arr_products,2) 				
				prod_Id = trim(arr_products(0,i))
				product_name = trim(arr_products(1,i))%>
	<OPTION value="<%=prod_Id%>" <%If trim(quest_id) = trim(prod_Id) Then%> selected <%End If%>><%=product_name%> </OPTION>
<%Next	
 End If	%>
</SELECT>
<%If IsNumeric(appid) Then%><input type="hidden" name="quest_id" id="quest_id" value="<%=quest_id%>"><%End If%>
</td></tr>		   
 <!-- start code --> 
 <%If trim(quest_id) <> "" And found_quest Then%> 		
 <tr><td width="100%" align="center" valign=top><table WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="3" align="center">	 
		<%if trim(PRODUCT_DESCRIPTION) <> "" then%>	
		<tr><td class="form_makdim" bgcolor="#DADADA" dir="<%=dir_obj_var%>" width="100%" align="<%=td_align%>"><%=breaks(PRODUCT_DESCRIPTION)%>	</td></tr>		
		<%end if%>	
		<%If Len(attachment_title) > 0 Then%>
		<tr>
			<td class="form_makdim" bgcolor=#DADADA dir="<%=dir_obj_var%>" width=100% align=<%=td_align%>>
			<a class="file_link" href="<%=strLocal%>/download/products/<%=attachment_file%>" target=_blank><%=attachment_title%></a>
			</td>
		</tr>	
		<%End If%>	
		<tr>
		<td align="<%=align_var%>" bgcolor=#C9C9C9 style="border-bottom: 1px solid #808080;border-top: 1px solid #808080;padding-right:15px;padding-left:15px">
		<table cellpadding=2 cellspacing=0 border=0 width=100% align="<%=align_var%>">
		<tr><td height=10 colspan=2 nowrap></td></tr> 
		<%If trim(appID) <> "" Then%>
		<tr>
			 <td align="<%=align_var%>" width=100%><input type="text" style="width:110;" value="<%=appID%>" class="Form_R" dir="ltr" ReadOnly></td>
			 <td align="<%=align_var%>" width=130 nowrap class="form_title">ID</td>   
		</tr>	
		<%End If%>	
		<tr>
			 <td align="<%=align_var%>" width=100%><input type="text" style="width:110;" value="<%=appeal_date%>" id="appeal_date" class="Form_R"  name="appeal_date" dir="ltr" ReadOnly></td>
			 <td align="<%=align_var%>" width=130 nowrap class="form_title"><!--תאריך הזנת הטופס--><%=arrTitles(4)%></td>   
		</tr>
		<%If Len(userName) > 0 Then%>
		<tr>
			 <td align="<%=align_var%>" width=100%><input type="text" style="width:150;" value="<%=vFix(userName)%>" class="Form_R" dir="<%=dir_obj_var%>" ReadOnly ID="Text1" NAME="Text1"></td>
			 <td align="<%=align_var%>" width=130 nowrap class="form_title"><!--הזנה ע"י עובד--><%=arrTitles(5)%></td>   
		</tr>
		<%End If%>
		<tr>		   
			 <td align="<%=align_var%>" width="100%">
			  <select id="appeal_status" name="appeal_status" dir="<%=dir_obj_var%>" class="norm">
			   <%For i=1 To Ubound(arr_Status)%><option value="<%=arr_Status(i, 0)%>" 
			   style="color: #FFFFFF; background-color:<%=trim(arr_Status(i, 2))%>"
			   <%if trim(arr_Status(i, 0)) = trim(appeal_status) then%> selected<%end if%>>&nbsp;<%=arr_Status(i, 1)%>&nbsp;</option>
			   <%Next%>
			 </select>
			 </td>
			 <td align="<%=align_var%>" nowrap class="form_title"><!--סטטוס--><%=arrTitles(6)%></td>   
		</tr>
			<%if quest_id=16504 then%>
				<tr >
			<td align="right" width="100%" dir="<%=dir_obj_var%>">
			<select name="CRM_Country"  dir="<%=dir_obj_var%>" class="norm" style="width: 250px" ID="CRM_Country">
			<option value="0">בחר יעד</option>
			
			<%set CountryList=con.GetRecordSet("SELECT Country_Id,Country_Name FROM Countries  ORDER BY Country_Name")
			do while not CountryList.EOF
			selCountryID=CountryList(0)
			selCountryName=CountryList(1)%>
			
			<option value="<%=selCountryID%>" <%if trim(AppCountryID)=trim(selCountryID) then%> selected <%end if%>><%=selCountryName%></option>
	<% CountryList.MoveNext
			Loop
			Set CountryList=Nothing%>
			</select>	</td>
			<td align="<%=align_var%>"  width=130 class="form_title" dir=rtl>יעד</td>
			</tr> 

			<%end if%>
			<%If trim(quest_id) = 16735 Then 'הוסף טופס רישום חתום%>
			<tr><td align="<%=align_var%>" width="100%">
			<script type="text/javascript"> 
			<!--
				var selectList_cache; // Variable for remember previous choice 
				function validateList(val) 
				{ 
					<%If JobId <> "463" And trim(appID) <> "" Then%>
					alert("לא ניתן להחליף שיוך הרישום"); 
					document.getElementById("UserIdOrderOwner").value=selectList_cache; 			
					<%End If%>
				}
			   //-->	 
			</script><select name="UserIdOrderOwner" dir="<%=dir_obj_var%>" class="norm" style="width: 250px" 
			ID="UserIdOrderOwner" onclick="selectList_cache=this.value" onchange="validateList(this.value);return false">
			<%set UserList=con.GetRecordSet("SELECT User_ID, UserName = (FIRSTNAME + ' ' + LASTNAME) FROM Users WHERE ORGANIZATION_ID = " & OrgID & " AND ACTIVE = 1 ORDER BY Case When User_ID = "&UserID&" Then 0 Else 1 End, FIRSTNAME, LASTNAME")
			do while not UserList.EOF
			selUserID=UserList(0)
			selUserName=UserList(1)%>
			<option value="<%=selUserID%>" <%if trim(UserIdOrderOwner)=trim(selUserID) then%> selected <%end if%>><%=selUserName%></option>
	<% UserList.MoveNext
			Loop
			Set UserList=Nothing%>
			</select>
			</td>
			<td align="<%=align_var%>" nowrap class="form_title"><b>למי שייך הרישום</b></td>
			</tr>	 
			<%End If%>		
		<tr><td height=10 colspan=2 nowrap></td></tr> 		
		</table></td></tr>
		<%If trim(COMPANIES) = "1" Then%>						
		<%If trim(appid) = "" And trim(addClient)="1" Then%>	  
		<tr><td align="<%=align_var%>" class="title_form"><!--פרטי איש קשר--><%=arrTitles(32)%></td></tr>
		<tr>
		<td align="<%=align_var%>" bgcolor=#C9C9C9 style="border-bottom: 1px solid #808080;padding-right:15px;padding-left:15px">
		<table cellpadding=2 cellspacing=0 border=0 width=100% align="<%=align_var%>">
		<tr><td colspan=3 height=10 nowrap></td></tr>
		<tr>		
			<%If appid = nil And trim(addClient)="1" Then%>	
			<td width=100% align="<%=align_var%>" rowspan=5 valign=top>
			<iframe style="display:none" name="frameContactsPr" id="frameContactsPr" src="contacts_private.asp" ALLOWTRANSPARENCY=true height="140" width=550 marginwidth="0" marginheight="0" hspace="0" vspace="0" scrolling="yes" frameborder="1" ></iframe>  
			<input type=hidden name="companyId" id="companyId" value="<%=companyID%>">
			<input type=hidden name="contactID" id="contactID" value="<%=contactID%>">			
			</td>
			<%End If%>			
			<td align="<%=align_var%>" width=270 nowrap valign=top>           
			<%If trim(companyID) <> "" And trim(appId) = "" Then%>
			<img src="../../images/delete_icon.gif" onclick='window.document.location.href="appeal.asp?quest_id=<%=quest_id%>"' title="שנה" class="hand" style="position:relative;top:3px">
			<%End If%>
			<input type=text dir="<%=dir_align%>" name="contact_name" value="<%=vFix(contactName)%>" maxlength=50 style="width:200" <%If trim(companyID) = "" Then%> class="Form" onKeyUp="refreshContactsPr(this,'<%=quest_id%>')" onKeyDown="refreshContactsPr(this,'<%=quest_id%>')" <%Else%> ReadOnly class="Form_R" <%End If%>>
			</td>
			<td align="<%=align_var%>" nowrap class="form_title" valign=top>					
			&nbsp;<b><!--שם מלא--><%=arrTitles(14)%></b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
		</tr> 			    
		<tr>
			<td align="<%=align_var%>" dir=ltr width=100%>			
			<input type=text dir="ltr" name="phone" value="<%=vFix(contactPhone)%>" maxlength="20" style="width:100" class="Form" ID="phone">                
			</td>
			<td align="<%=align_var%>" nowrap class="form_title"><b><!--טלפון--><%=arrTitles(16)%></b></td>
		</tr>
		<tr>
			<td align="<%=align_var%>" dir=ltr width=100%>			
			<input type=text dir="ltr" name="cellular" value="<%=vFix(contactCellular)%>" style="width:100" maxlength=20 class="Form" ID="cellular">              
			</td>
			<td align="<%=align_var%>" nowrap class="form_title"><b><!--טלפון נייד--><%=arrTitles(17)%></b></td>
		</tr>
		<tr>
			<td align="<%=align_var%>" width=100%>
				<input type=text name="email" dir=ltr value="<%=vFix(contactEmail)%>" style="width:200" maxlength=50 class="Form" ID="email">              
			</td>
			<td align="<%=align_var%>" nowrap class="form_title"><b>Email</b></td>
		</tr>
		<tr><td colspan=2 height=10 nowrap></td></tr>
		</table></td></tr>
		<%ElseIf appid = nil And trim(addClient)="2" Then%>	
		<tr><td align="<%=align_var%>" class="title_form"><!--פרטי איש קשר--><%=arrTitles(32)%></td></tr>
		<tr>
		<td align="<%=align_var%>" bgcolor=#C9C9C9 style="padding-right:15px;padding-left:15px">
		<table cellpadding=2 cellspacing=0 border=0 width=100% align="<%=align_var%>">		
		<tr>
			<td width=100% align="<%=align_var%>" rowspan=10 valign=top>
			<iframe style="display:none" name="frameContacts" id="frameContacts" src="contacts.asp" ALLOWTRANSPARENCY=true height="255" width=500 marginwidth="0" marginheight="0" hspace="0" vspace="0" scrolling="yes" frameborder="1" ></iframe>  
			<input type=hidden name="contactID" id="contactID" value="<%=contactID%>">			
			</td>
			<td align="<%=align_var%>" width=270 nowrap valign=top>           
			<%If trim(contactID) <> "" And trim(appId) = "" Then%>
			<img src="../../images/delete_icon.gif" onclick='window.document.location.href="appeal.asp?quest_id=<%=quest_id%>"' title="שנה" class="hand" style="position:relative;top:3px">
			<%End If%>
			<input type=text dir="<%=dir_align%>" name="contact_name" value="<%=vFix(contact_name)%>" maxlength=50 style="width:200" <%If trim(contactID) = "" Then%> class="Form" onKeyUp="refreshContacts(this,'<%=quest_id%>')" onKeyDown="refreshContacts(this,'<%=quest_id%>')" <%Else%> ReadOnly class="Form_R" <%End If%>>
			</td>
			<td align="<%=align_var%>" width=120 nowrap class="form_title" valign=top>					
			&nbsp;<b><!--שם מלא--><%=arrTitles(14)%></b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
		</tr> 		
		<tr>
			<td align="<%=align_var%>" width=100%>
			<INPUT type=image src="../../images/icon_find.gif" name=messangers_list id="messangers_list" onclick='window.open("../companies/messangers_list.asp","MessangersList","left=300,top=100,width=250,height=350,scrollbars=1");return false;' title="<%=arrTitles(21)%>">
			<input type=text dir="<%=dir_obj_var%>" name="messangerName" value="<%=vFix(messangerName)%>" maxlength=50 class="Form" ID="Text3" onfocus="document.all('frameContacts').style.display='none'">         
			</td>
			<td align="<%=align_var%>" nowrap class="form_title">
			&nbsp;<b><!--עיסוק--><%=arrTitles(15)%></b></td>
		</tr>	
		<tr>
			<td align="<%=align_var%>" width=100%>
			<input type=text dir="<%=dir_obj_var%>" name="contact_address" value="<%=vFix(contact_address)%>"  style="width:180" class="Form" ID="contact_address"></td>
			<td align="<%=align_var%>" nowrap class="form_title"><!--כתובת--><%=arrTitles(22)%></td>
		</tr>             
		<tr>
			<td align="<%=align_var%>" width=100%>
			<INPUT type=image src="../../images/icon_find.gif" onclick='window.open("../companies/cities_list.asp?fieldID=contact_city_Name","CitiesList","left=100,top=50,width=450,height=450,scrollbars=1");return false;' title="<%=arrTitles(21)%>">&nbsp;                             
			<input type=text dir="<%=dir_obj_var%>" class="Form" name="contact_city_Name" id="contact_city_Name" value="<%=contact_city_Name%>">
			</td>
			<td align="<%=align_var%>" nowrap class="form_title"><!--עיר--><%=arrTitles(23)%></td>
		</tr>
      <tr>
          <td align="<%=align_var%>" dir=ltr width=100%>
		   <input type=text dir="ltr" name="contact_zip_code" value="<%=vFix(contact_zip_code)%>" size="10" maxlength=10 <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="contact_zip_code">
           </td>
          <td align="<%=align_var%>" class="form_title">&nbsp;<!--מיקוד--><%=arrTitles(35)%>&nbsp;</td>
      </tr>		    			    
		<tr>
			<td align="<%=align_var%>" dir=ltr width=100%>			
			<input type=text dir="ltr" name="phone" value="<%=vFix(contact_phone)%>" maxlength="20" class="Form" ID="phone">                
			</td>
			<td align="<%=align_var%>" nowrap class="form_title"><b><!--טלפון--><%=arrTitles(16)%></b></td>
		</tr>
		<tr>
			<td align="<%=align_var%>" dir=ltr width=100%>			
			<input type=text dir="ltr" name="fax" value="<%=vFix(contact_fax)%>" maxlength="20" class="Form" ID="fax">                
			</td>
			<td align="<%=align_var%>" nowrap class="form_title"><b><!--פקס--><%=arrTitles(13)%></b></td>
		</tr>
		<tr>
			<td align="<%=align_var%>" dir=ltr width=100%>			
			<input type=text dir="ltr" name="cellular" value="<%=vFix(contact_cellular)%>" maxlength=20 class="Form" ID="cellular">              
			</td>
			<td align="<%=align_var%>" nowrap class="form_title"><b><!--טלפון נייד--><%=arrTitles(17)%></b></td>
		</tr>
		<tr>
			<td align="<%=align_var%>" dir=ltr width=100%>
				<input type=text name="email" dir=ltr value="<%=vFix(contact_email)%>" style="width:180" maxlength=50 class="Form" ID="email">              
			</td>
			<td align="<%=align_var%>" nowrap class="form_title"><b>Email</b></td>
		</tr>
       <tr>
           <td align="<%=align_var%>" width=100%>
           <INPUT type=image src="../../images/icon_find.gif" name=types_list id="types_list" onclick='window.open("../companies/types_list.asp?contactID=<%=contactID%>&contact_types=" + contact_types.value,"TypesList","left=300,top=100,width=250,height=350,scrollbars=1");return false;' title="<%=arrTitles(21)%>">&nbsp;
           <span class="Form_R" dir="<%=dir_obj_var%>" style="width:180;line-height:16px" name="types_names" id="types_names"><%=types%></span>
           <input type=hidden name=contact_types id=contact_types value="<%=contact_types%>">
           </td>
           <td align="<%=align_var%>" nowrap valign=top><b><!--סיווג--><%=arrTitles(24)%></b></td>
      </tr>  
      <tr>
      <td align="<%=align_var%>" width=100% valign=top colspan=2>
      <textarea class="Form" dir="<%=dir_obj_var%>" style="width:250;" rows=5 name="contact_desc" id="contact_desc"><%=trim(contact_desc_short)%></textarea>      
      </td>
      <td align="<%=align_var%>" nowrap valign=top><b><!--פרטים נוספים--><%=arrTitles(25)%></b></td>
      </tr>  
      <tr>
        <td align="<%=align_var%>" width=100% valign=top colspan=2>                        
		<select name="responsible_id" dir="<%=dir_obj_var%>" class="norm" style="width:250" ID="responsible_id">
			<option value=""><!-- בחר --><%=arrTitles(37)%></option>
			<%set rs_users=con.GetRecordSet("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME FROM Users WHERE ORGANIZATION_ID = " & OrgID & " AND ACTIVE = 1 ORDER BY FIRSTNAME + ' ' + LASTNAME")
			If Not rs_users.eof Then
				arr_users = rs_users.getRows()
			End If
			Set rs_users = Nothing
			If isArray(arr_users) Then
			For uu=0 To Ubound(arr_users,2)%>
			<option value="<%=arr_users(0,uu)%>" <%if trim(responsible_id)=trim(arr_users(0,uu)) then%> selected <%end if%>><%=arr_users(1,uu)%></option>
			<%Next
			  End If%>
		</select>	          
		</td>
        <td align="<%=align_var%>" nowrap valign=top><b>&nbsp;<!--אחראי--><%=arrTitles(36)%>&nbsp;</b></td>
      </tr>             			
	  <tr><td colspan=3 height=10 nowrap></td></tr>	
	  </table></td></tr>
   	  <tr><td align="<%=align_var%>" class="title_form"><!--פרטי החברה--><%=arrTitles(33)%></td></tr>
	  <tr>
	  <td align="<%=align_var%>" bgcolor=#C9C9C9 style="border-bottom: 1px solid #808080;padding-right:15px;padding-left:15px">
	  <table cellpadding=2 cellspacing=0 border=0 width=100% align="<%=align_var%>">					  
	  <tr><td colspan=3 height=10 nowrap></td></tr>
	  <tr>
			<td width=100% align="<%=align_var%>" rowspan=9 valign=top>
			<iframe style="display:none" name="frameCompanies" id="frameCompanies" src="companies.asp" ALLOWTRANSPARENCY=true height="235" width=550 marginwidth="0" marginheight="0" hspace="0" vspace="0" scrolling="yes" frameborder="1" ></iframe>  
			<input type=hidden name="companyId" id="companyId" value="<%=companyID%>">				
			</td>		
			<td align="<%=align_var%>" width=270 nowrap>			
			<img src="../../images/delete_icon.gif" <%If trim(companyID) = "" Or trim(appId) <> "" Then%> style="display:none" <%End If%> name="delete_company" id="delete_company" onclick='return ClearCompany()' title="שנה" class="hand" style="position:relative;top:3px">
		    <input type=text dir="<%=dir_align%>" name="company_name" value="<%=vFix(company_name)%>"  style="width:200" <%If trim(companyID) = "" Then%> class="Form" <%Else%> ReadOnly class="Form_R" <%End If%> onKeyUp="return refreshCompany(this,'<%=quest_id%>')" onKeyDown="return refreshCompany(this,'<%=quest_id%>')">
		    </td>
        	<td align="<%=align_var%>" width=120 nowrap class="form_title">
			&nbsp;<b><!--חברה--><%=arrTitles(3)%></b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font>
			</td> 
	  </tr>	
      <tr>
        <td align="<%=align_var%>" width=100%><input type=text dir="<%=dir_obj_var%>" name="address" value="<%=vFix(address)%>"  style="width:200" <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="address"></td>
        <td align="<%=align_var%>" class="form_title">&nbsp;<!--כתובת--><%=arrTitles(26)%>&nbsp;</td>
     </tr>             
     <tr>
         <td align="<%=align_var%>" width=100%>
         <%If private_flag = "0" Then%>
         <INPUT type=image src="../../images/icon_find.gif" onclick='window.open("../companies/cities_list.asp","CitiesList","left=100,top=50,width=450,height=450,scrollbars=1");return false;' title="<%=arrTitles(21)%>">&nbsp;
         <%End If%>
         <input type=text dir="<%=dir_obj_var%>" name="cityName" value="<%=vFix(cityName)%>" <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="cityName">         
         </td>
         <td align="<%=align_var%>" class="form_title">&nbsp;<!--עיר--><%=arrTitles(27)%>&nbsp;</td>
      </tr>
      <tr>
          <td align="<%=align_var%>" dir=ltr width=100%>
		   <input type=text dir="ltr" name="zip_code" value="<%=vFix(zip_code)%>" size="10" maxlength=10 <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="zip_code">
           </td>
          <td align="<%=align_var%>" class="form_title">&nbsp;<!--מיקוד--><%=arrTitles(35)%>&nbsp;</td>
      </tr>                              
      <tr>
          <td align="<%=align_var%>" dir=ltr width=100%>
		   <input type=text dir="ltr" name="phone1" value="<%=vFix(phone1)%>" size="20" maxlength=20 <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="phone1">
           </td>
          <td align="<%=align_var%>" class="form_title">&nbsp;<!--1 טלפון--><%=arrTitles(28)%>&nbsp;</td>
      </tr>                                
      <tr>  
          <td align="<%=align_var%>" dir=ltr width=100%>
		  <input type=text dir="ltr" name="phone2" value="<%=vFix(phone2)%>" size="20" maxlength=20 <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="phone2">
          </td>
          <td align="<%=align_var%>" class="form_title">&nbsp;<!--2 טלפון--><%=arrTitles(29)%>&nbsp;</td>
      </tr>  
      <tr>  
          <td align="<%=align_var%>" dir=ltr width=100%>            
			<input type=text dir="ltr"  name="fax1" value="<%=vFix(fax1)%>" size="20" maxlength=20 <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="fax1">
           </td>
          <td align="<%=align_var%>" class="form_title">&nbsp;<!--פקס--><%=arrTitles(13)%>&nbsp;</td>
      </tr>                          
       <tr>
         <td align="<%=align_var%>" width=100%>         
         <input dir=ltr type=text name="company_email" value="<%=vFix(company_email)%>" style="width:200" <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> ID="company_email"></td>
         <td align="<%=align_var%>" class="form_title">&nbsp;Email&nbsp;</td>
      </tr>   
      <tr>
         <td align="<%=align_var%>" width=100%>         
         <input dir=ltr type=text name="url" value="<%=vFix(url)%>" style="width:200" <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%>  ID="url"></td>
         <td align="<%=align_var%>" class="form_title">&nbsp;<!--אתר--><%=arrTitles(30)%>&nbsp;</td>
      </tr>
      <tr>
      <td align="<%=align_var%>" colspan=2>
      <textarea dir="<%=dir_obj_var%>" <%If private_flag = "0" Then%>class="Form"<%Else%>class="Form_R" readonly<%End If%> style="width:330" name="company_desc" rows=5 ID="company_desc"><%=trim(company_desc)%></textarea>
      </td>
      <td align="<%=align_var%>" class="form_title" valign=top>&nbsp;<!--פרטים נוספים--><%=arrTitles(25)%>&nbsp;</td>
      </tr>	  
      <%End If%> 
      </table></td></tr>
      <%If trim(addClient)="" Or IsNull(addClient) Or trim(appID) <> "" Then%> 
	   <tr><td bgcolor=#C9C9C9 style="border-bottom: 1px solid #808080;padding-right:15px;padding-left:15px">
	   <table cellpadding=0 cellspacing=0 width=100%>
	   <tr><td height=5 nowrap></td></tr>
		<tr  <%If trim(private_flag) <> "0" Then%> style="display: none;" <%End If%>>
			<td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>" >
			<input type="hidden" name="companyId" id="companyId" value="<%=companyID%>">
			<input type="text" id="CompanyName" name="CompanyName" dir="<%=dir_obj_var%>" style="width: 350px" class="Form_R" readonly 
			value="<%If Len(CompanyName) > 0 Then%><%=vFix(CompanyName)%><%Else%><%End If%>" 	>&nbsp;&nbsp;
			<input type="button" value="בחר מרשימה" class="add" style="width: 90px" onclick="return openCompaniesList();" >
			&nbsp;&nbsp;<input type="button" value="בטל בחירה"	 onclick="return removeCompany();" class="but_site" style="width: 80px" >
			</td>
			<td align="<%=align_var%>" width=120 nowrap class="form_title"><!--קישור ל--><%=arrTitles(7)%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
			</tr> 		
			<tbody name="contacter_body" id="contacter_body" <%If (trim(companyID) = "" Or IsNull(companyID)) Then%> style="display:none" <%End if%>>
			<tr><td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>"><input type="hidden" name="contactID" id="contactID" value="<%=contactID%>">
			<input type="text" id="contactName" name="contactName" dir="<%=dir_obj_var%>" style="width: 350px" class="Form_R" readonly 
			value="<%If Len(contactName) > 0 Then%><%=vFix(contactName)%><%Else%><%End If%>" 	>&nbsp;&nbsp;
			<input type="button" value="בחר מרשימה" class="add" style="width: 90px" onclick="return openContactsList();" >
			&nbsp;&nbsp;<input type="button" value="בטל בחירה"	 onclick="return removeContact();" class="but_site" style="width: 80px">				
			</td>
			<td align="<%=align_var%>" nowrap class="form_title"><!--קישור ל--><%=arrTitles(8)%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td>
			</tr>   
			</tbody>		       
			<tbody name="project_body" id="project_body" <%If trim(companyID) = "" Or IsNull(companyID) Then%> style="display:none" <%End if%>>
			<TR>
				<td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>"><input type="hidden" name="project_id" id="project_id" value="<%=projectId%>">
				<input type="text" id="projectName" name="projectName" dir="<%=dir_obj_var%>" style="width: 350px" class="Form_R" readonly 
				value="<%If Len(projectName) > 0 Then%><%=vFix(projectName)%><%Else%><%End If%>" 	>&nbsp;&nbsp;
				<input type="button" value="בחר מרשימה" class="add" style="width: 90px" onclick="return openProjectsList();" >
				&nbsp;&nbsp;<input type="button" value="בטל בחירה"	 onclick="return removeProject();" class="but_site" style="width: 80px" >				
				</TD>
				<td align="<%=align_var%>" nowrap class="form_title"><!--קישור ל--><%=arrTitles(9)%><%=trim(Request.Cookies("bizpegasus")("Projectone"))	%></td>
			</TR>
			</tbody>
			<tbody name="mechanism_body" id="mechanism_body" <%If trim(projectID) = "" Or IsNull(projectID) Then%> style="display:none" <%End if%>>
			<TR>
				<td align="<%=align_var%>" width=100%>
				<select name="mechanism_id" id="mechanism_id" class="norm" dir="<%=dir_obj_var%>" style="width:350px">
				<%If trim(lang_id) = "1" Then%>
				<option value=""><%=String(23,"-")%> לא מקושר למנגנון <%=String(23,"-")%></option>
				<%Else%>
				<option value=""><%=String(23,"-")%> Not linked to Sub-Project <%=String(23,"-")%></option>
				<%End If%> 
				<%If trim(projectID) <> "" Then%>	
				<%sqlstr = "Select mechanism_id, mechanism_name FROM mechanism WHERE ORGANIZATION_ID = " & OrgID & " AND Project_ID  = " & projectID & " Order BY mechanism_name"
				  set rs_mech = con.getRecordSet(sqlstr)
				  If Not rs_mech.eof Then
					 arr_mech = rs_mech.getRows()
				  End If
				  Set rs_mech = Nothing
				  If isArray(arr_mech) Then
				  For mm=0 To Ubound(arr_mech,2)%>
				<option value="<%=arr_mech(0,mm)%>" <%if trim(arr_mech(0,mm))=trim(mechanismID) then%> selected <%end if%>><%=arr_mech(1,mm)%></option>		
				<%Next
				End If%>	
				<%End If%>	
				</select>
				</TD>
				<td align="<%=align_var%>" nowrap class="form_title"><!--קישור ל--><%=arrTitles(9)%><%=arrTitles(38)%></td>
			</TR>			
			</tbody>
			<tr><td width="100%" height="10" colspan=2></td></tr>
			</table>
		</td>
	 </tr>  
	<%End If%> 
	<%Else%>	
	<%End If%>
	<tr>
		<td width="100%" align="<%=align_var%>" dir="<%=dir_var%>">
			<table width=100% border=0 cellspacing=1 cellpadding=3 bgcolor=White dir="<%=dir_var%>">
			<!-- start fields dynamics -->	
			<%If trim(appId) = "" Then%>			
			<!--#INCLUDE FILE="default_fields.asp"-->
			<%else%>
			<!--#INCLUDE FILE="upd_fields.asp"-->
			<%End If%>
			<!-- end fields dynamics -->		
		    </table>
		</td>
	</tr>		
	<%if is_must_fields then %>
	<tr><td width="100%" height="10" colspan=4></td></tr>
	<tr><td colspan="4" height="10" align="<%=td_align%>" dir="<%=dir_obj_var%>">&nbsp;<font color=red><!--שדות חובה--><%=arrTitles(19)%></font>&nbsp;</td></tr>
	<%end if%>				
	<tr><td width="100%" height="10" colspan=4></td></tr>
	<tr><td align=center colspan=4 dir="<%=dir_var%>">
	<INPUT style="width:150px" type="button" class="but_menu" value="<%=arrTitles(34) & " " & trim(Request.Cookies("bizpegasus")("TasksOne"))%>" onClick="return CheckFields('send')">
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<INPUT style="width:90px" type="button" class="but_menu" value="<%=arrButtons(4)%>" id="Button1" name="Button4" onClick="return CheckFields('save')">
	</td></tr>
	<tr><td width="100%" height="10" colspan=4></td></tr>
</table></td></tr>	
</table>
</form></td></tr> 
<%End If%>     
</div>
</table> 
<% If (isNumeric(quest_id) = true) Then%>
<script type="text/javascript">
<!--
    <%	If trim(lang_id) = "1" Then
			str_alert = "!!!נא למלא את השדה"
		Else
			str_alert = "Please provide the answer!!!"
		End If%> 		 		
	function click_fun()
	{ 
        <% sqlStr = "SELECT Field_Id,Field_Must,Field_Type FROM Form_Field Where product_id=" & quest_id & " Order by Field_Order"
		'Response.Write sqlStr
		'Response.End
		set fields=con.GetRecordSet(sqlStr)
		If not fields.eof Then
			arr_fields = fields.getRows()
		End If
		Set fields = Nothing
		If isArray(arr_fields) Then
		For ff=0 To Ubound(arr_fields,2)
		  Field_Id = arr_fields(0,ff)				
		  Field_Must = trim(arr_fields(1,ff))
		  Field_Type = trim(arr_fields(2,ff))
		  'Response.Write  Field_Type
		  If Field_Must = "True"  Then %>		
       
       field =  document.getElementsByName("field<%=Field_Id%>");  
       <%If trim(Field_Type) = "5" Or trim(Field_Type) = "8" Or trim(Field_Type) = "9" Or trim(Field_Type) = "11" Or trim(Field_Type) = "12" Then%>			
		if(field != null)
		{
			answered = false;
			for(j=0;j<field.length;j++)
			{									
				if(field(j).checked == true)
				{
					answered = true;					
				}	
			}		
			if(answered == false)
			{
				window.alert("<%=str_alert%>");			
				field(0).focus();			 
				return false;
			}
		}	  	 	   
	  <%Else%>
	    if((field != null) && document.all("field<%=Field_Id%>").value == '') 
	      { document.all("field<%=Field_Id%>").focus(); 
		    window.alert("<%=str_alert%>"); 
		    return false; } 
	  <%   End If	 
	     End If
        Next
       End If %>   
       return true;	}
//-->
</script> 
<%End If%> 
</body>
</html>
<%set con=nothing%>
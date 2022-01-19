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
	if trim(request.QueryString ("Is_GroupsTour"))="1" then
	Is_GroupsTour=trim(request.QueryString ("Is_GroupsTour"))
	else
	Is_GroupsTour=0
	end if
		
	'טופס רשום לטיול
	ReservationId = trim(Request("ReservationId"))
	Country_CRM=0
	'response.Write ("ReservationId="& ReservationId)
	'response.end
'response.Write "Country_CRM="&Country_CRM
if trim(ReservationId)<>"" then
sqlstr="SELECT T.Tour_Id,Country_CRM,Insert_Date, TR.Reservation_Id	FROM pegasus.dbo.Tour_Reservations TR LEFT OUTER JOIN pegasus.dbo.Tours T ON TR.Tour_Id = T.Tour_Id	WHERE (Contact_Id = "& contactID &") and  Reservation_Id=" & ReservationId
set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			Country_CRM = trim(rs_pr("Country_CRM"))
		else
		Country_CRM=0
		End If
		set rs_pr = Nothing
end if
if not IsNumeric(Country_CRM) then Country_CRM=0
	title = false
	If trim(companyId) <> "" Then
		sqlstr = "Select company_Name from companies WHERE company_Id = " & companyId
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			companyName = trim(rs_pr("company_Name"))
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(contactID) <> "" Then
		sqlstr = "Select CONTACT_NAME from contacts WHERE contact_Id = " & contactID
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			contactName = trim(rs_pr(0))
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(projectID) <> "" Then
		sqlstr = "Select project_Name from projects WHERE project_Id = " & projectID
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			projectName = trim(rs_pr("project_Name"))
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	from = trim(Request("from"))
	If trim(from) = "" Then
		from = "pop_up"
	End If 	
	PathCalImage = "../../"
	
	arr_Status = session("arr_Status")
	
	sqlstr =  "Select Langu, PRODUCT_NAME, PRODUCT_DESCRIPTION, RESPONSIBLE, FILE_ATTACHMENT, ATTACHMENT_TITLE From Products WHERE PRODUCT_ID = " & quest_id
	'Response.Write sqlstr
	'Response.End
	set rsq = con.getRecordSet(sqlstr)
	If not rsq.eof Then		
		Langu = trim(rsq(0))
		productName = trim(rsq(1))
		PRODUCT_DESCRIPTION = trim(rsq(2))
		RESPONSIBLE = trim(rsq(3))
		attachment_file = trim(rsq(4))
		attachment_title = trim(rsq(5))		
	End If
	set rsq = Nothing	
	if Langu = "eng" then
		td_align = "left" : pr_language = "eng"
	else
		td_align = "right" : 	pr_language = "heb"
	end if
	
	userName = user_name 'cokie
	If IsNumeric(appid) Then
	sqlstr = "Exec dbo.get_appeals '','','','','" & OrgID & "','','','','','','','" & appid & "','',''"
'	Response.Write("sqlstr=" & sqlstr)
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
	private_flag = trim(app("private_flag"))
	UserIdOrderOwner = trim(app("User_Id_Order_Owner"))
	AppCountryID = trim(app("appeal_CountryId"))
	If trim(appeal_status) = "1" And trim(UserID) = trim(RESPONSIBLE) Then
		sqlstring="UPDATE appeals set appeal_status = '2' WHERE ORGANIZATION_ID = " & OrgID & " AND appeal_id = " & appid 	
		'Response.Write(sqlstring)
		'Response.End 
		con.ExecuteQuery(sqlstring) 
		appeal_status = "2"
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
	
	If trim(appeal_status) = "1" And trim(UserID) = trim(RESPONSIBLE) Then
		sqlstring="UPDATE appeals set appeal_status = '2' WHERE ORGANIZATION_ID = " & OrgID & " AND appeal_id = " & appid 	
		'Response.Write(sqlstring)
		'Response.End 
		con.ExecuteQuery(sqlstring) 
		appeal_status = "2"
    End If			
	
	title = false
	If trim(companyId) <> "" Then
		sqlstr = "Select company_Name from companies WHERE company_Id = " & companyId
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			companyName = trim(rs_pr("company_Name"))
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(contactID) <> "" Then
		sqlstr = "Select CONTACT_NAME from contacts WHERE contact_Id = " & contactID
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			contactName = trim(rs_pr(0))
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(projectID) <> "" Then
		sqlstr = "Select project_Name from projects WHERE project_Id = " & projectID
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			projectName = trim(rs_pr("project_Name"))
			title = true		
		End If
		set rs_pr = Nothing
	End If	
%>
<%   
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
<html>
	<head>
		<title>
			<%=productName%>
		</title>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script>
function selectGuide(comboDepartureId)
{
//alert(comboDepartureId.value)
DepartureId=comboDepartureId.value
var d = document.getElementById('Guide_Id');
 d = comboDepartureId.value;

 var xmlhttp;
if (window.XMLHttpRequest)
  {// code for IE7+, Firefox, Chrome, Opera, Safari
  xmlhttp=new XMLHttpRequest();
  }
else
  {// code for IE6, IE5
  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
  }
  
      if(isNaN(DepartureId) == false)
   { // var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");

	 xmlhttp.open("POST", 'selectToursGuide.asp', false);
	 xmlhttp.send('<?xml version="1.0" encoding="UTF-8"?><request><DepartureId>' + DepartureId + '</DepartureId></request>');	 
	 
			result = new String(xmlhttp.responseText);						
//alert(result)
if (result !='')
{
 var element = document.getElementById('Guide_Id');
 element.value = result;
//alert(document.getElementById("Guide_Id'))
//var ddl = document.getElementById("Guide_Id');
//var opts = ddl.options.length;
//alert(opts)
}


		
}
}
function selectToursDate(comboTourId) 
{ 
   TourId = comboTourId.value;
 var xmlhttp;
if (window.XMLHttpRequest)
  {// code for IE7+, Firefox, Chrome, Opera, Safari
  xmlhttp=new XMLHttpRequest();
  }
else
  {// code for IE6, IE5
  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
  }
  
      if(isNaN(TourId) == false)
   { // var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");

	 xmlhttp.open("POST", 'selectToursDate.asp', false);
	 xmlhttp.send('<?xml version="1.0" encoding="UTF-8"?><request><TourId>' + TourId + '</TourId></request>');	 
	 
			result = new String(xmlhttp.responseText);						
		
		arr_ToursDate= result.split(";");		
		window.document.all("ToursDate_Id").length = 1;
	
	
			for (i=0; i < arr_ToursDate.length-1; i++)
		{			
			arr_TourDate = new String(arr_ToursDate[i]);
			arr_TourDate = arr_TourDate.split(",")			
			window.document.all("ToursDate_Id").options[window.document.all("ToursDate_Id").options.length]=
            new Option(arr_TourDate[1],arr_TourDate[0]); 
	    }
	    
	 
		
	    /*ToursDate_Id.style="dispaly:block"*/
	 }
}
</script>
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
	
	function openRelationList()
	{
	
		var CRMCountryId = <%=Country_CRM%>
		var contactID = <%=contactID%>
		if (CRMCountryId=='0')
		{alert("הטיול אינו מקושר ליעד")
		return false;
		}
		// document.getElementById("companyID").value;
		window.open('../appeals/listInterestedForm.asp?CRMCountryId=' + CRMCountryId +'&contactID=' + contactID,'winList','top=50, left=50, width=850, height=600, scrollbars=1');
	return false;
	}	
	
	function chooseProject(ProjectId)
	{
			var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");
				
			xmlHTTP.open("POST", '../tasks/getmechanism.asp', false);
			xmlHTTP.send('<?xml version="1.0" encoding="UTF-8"?><request><projectID>' + ProjectId + '</projectID></request>');	 
				
			result = new String(xmlHTTP.ResponseText);						
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
	function removeRelation()
	{		
		window.document.getElementById("RelationId").value = "";
		window.document.getElementById("RelationName").value = "";
		
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

function popupcal(elTarget){
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

function CheckFields()
{
//alert ("chk")
	if(window.document.form1.document_name.value != '' ||
	window.document.form1.document_file.value != '')
		if(checkfile() == false)
			return false;
					
	if(click_fun() == false)
		return false;
		
		<%if quest_id=16735 then%>
		if (window.document.getElementById("RelationId").value=='')
	{
	alert("אנא בחר טפסי מתעניין לטיול")
	return false;
	}
		<%end if%>
	<%if quest_id=16504 then%>
//	alert(window.document.getElementById("CRM_Country").value)
		if (window.document.getElementById("CRM_Country").value=='0')
	{
	alert("אנא בחר לאיזה יעד נסיעה הינך מתעניין")
	return false;
	}
	<%end if%>
	
		
	document.form1.submit();				
	return false;		
}	   	 

function checkfile()
{
	  if(document.form1.document_name.value=='')
	  {
			<%If trim(lang_id) = "1" Then%>
			window.alert("נא למלא כותרת מסמך");
			<%Else%>
			window.alert("Please insert document title");
			<%End If%>
			document.form1.document_name.focus();
			return false;		
	  }
	  if (document.form1.document_file.value=='')
	  {
			<%If trim(lang_id) = "1" Then%>
			window.alert('! נא לבחור קובץ');
			<%Else%>
			window.alert('Please choose the document!');
			<%End If%>
			document.form1.document_file.focus();
			return false;
	  }
	  else	
	  {
			var fname=new String();
			var fext=new String();
			var extfiles=new String();
			fname=document.form1.document_file.value;
			indexOfDot = fname.lastIndexOf('.')
			fext=fname.slice(indexOfDot+1,-1)		
			fext=fext.toUpperCase();
			extfiles='HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT,ZIP,MSG,TIFF,TIF';		
			if(extfiles.indexOf(fext)>-1)
				return true;
			else
			    <%If trim(lang_id) = "1" Then%>
				window.alert(':סיומת הקובץ - אחת מרשימה' + '\n' + 'HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT,ZIP,MSG,TIFF,TIF');
				<%Else%>
				window.alert('The file ending should be one of the these:' + '\n' + 'HTML,HTM,JPG,JPEG,GIF,BMP,PNG,DOC,XLS,PDF,TXT,ZIP,MSG,TIFF,TIF');
				<%End If%>
			return false;
	  }  
	 return true;
}
//-->
		</script>
	</head>
	<body style="margin:0px; background-color:#E6E6E6">
		<div align="center">
			<form id="form1" name="form1" action="addappeal.asp?from=<%=from%>" method="POST" ENCTYPE="multipart/form-data">
			<input type=hidden id="Is_GroupsTour" name="Is_GroupsTour" value="<%=Is_GroupsTour%>">
				<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
					<tr>
						<td class="title_form" style="font-size:14pt;line-height:24px" width="100%" align="center"><%=productName%>
							<input type="hidden" name="quest_id" id="quest_id" value="<%=quest_id%>"> <input type="hidden" name="ReservationId" id="ReservationId" value="<%=ReservationId%>">
						</td>
					</tr>
					<!-- start code -->
					<%If trim(quest_id) <> "" Then%>
					<%	set prod = con.GetRecordSet("Select Langu from products where product_id=" & quest_id)
	if not prod.eof then					
	if prod("Langu") = "eng" then
		dir_align = "ltr"  :  	td_align = "left"  :  	pr_language = "eng"
	else
		dir_align = "rtl"   :  	td_align = "right"  :  	pr_language = "heb"
	end if
end if
set prod = nothing%>
					<tr>
						<td width="100%" align="center" valign="top">
							<table WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0" align="center">
								<%if trim(PRODUCT_DESCRIPTION) <> "" then%>
								<tr>
									<td align=<%=td_align%> width="100%">
										<TABLE BORDER="0" CELLSPACING="1" CELLPADDING="3" width="100%" ID="Table1">
											<TR>
												<td class="form_makdim" dir="<%=dir_obj_var%>" width="100%" align=<%=td_align%>>
													<%=breaks(PRODUCT_DESCRIPTION)%>
												</td>
											</TR>
											<%If Len(attachment_title) > 0 Then%>
											<TR>
												<td class="form_makdim" dir="rtl" width="100%" align=<%=td_align%>>
													<a class="file_link" href="<%=strLocal%>/download/products/<%=attachment_file%>" target=_blank>
														<%=attachment_title%>
													</a>
												</td>
											</TR>
											<%End If%>
										</TABLE>
									</td>
								</tr>
								<%end if%>
								<tr>
									<td align="<%=align_var%>" bgcolor="#C9C9C9" style="border-bottom: 1px solid #808080;border-top: 1px solid #808080;padding-right:15px;padding-left:15px">
										<table cellpadding=2 cellspacing=0 border=0 width=450 align="<%=align_var%>">
											<tr>
												<td height="10" colspan="2" nowrap></td>
											</tr>
											<tr>
												<td align="<%=align_var%>" width="100%"><input type="text" style="width:110;" value="<%=appeal_date%>" id="appeal_date" class="Form_R"  name="appeal_date" dir="<%=dir_var%>" ReadOnly></td>
												<td align="<%=align_var%>" width=130 nowrap class="form_title"><!--תאריך הזנת הטופס--><%=arrTitles(4)%></td>
											</tr>
											<tr>
												<td align="<%=align_var%>" width="100%"><input type="text" style="width:150;" value="<%=vFix(userName)%>" class="Form_R" dir="<%=dir_obj_var%>" ReadOnly ID="Text1" NAME="Text1"></td>
												<td align="<%=align_var%>" width=130 nowrap class="form_title"><!--הזנה ע"י עובד--><%=arrTitles(5)%></td>
											</tr>
											<tr>
												<td align="<%=align_var%>" width="100%">
													<select id="appeal_status" name="appeal_status" dir="<%=dir_obj_var%>" class="norm">
														<%For i=1 To Ubound(arr_Status)%>
														<option value="<%=arr_Status(i, 0)%>" 
			   style="color: #FFFFFF; background-color:<%=trim(arr_Status(i, 2))%>"
			   <%if cstr(arr_Status(i, 0)) = cstr(appeal_status) then%> selected <%end if%>>&nbsp;<%=arr_Status(i, 1)%>&nbsp;</option>
														<%Next%>
													</select>
												</td>
												<td align="<%=align_var%>" nowrap class="form_title"><!--סטטוס--><%=arrTitles(6)%></td>
											</tr>
											<%If appid = nil Then%>
											<%if trim(COMPANIES) = "1" THEN%>
											<tr  <%If trim(private_flag) <> "0" Then%> style="display: none;" <%End If%> >
												<td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>">
													<input type="hidden" name="companyId" id="companyId" value="<%=companyID%>"> <input type="text" id="CompanyName" name="CompanyName" dir="<%=dir_obj_var%>" style="width: 350px" class="Form_R" readonly 
			value="<%If Len(companyName) > 0 Then%><%=vFix(companyName)%><%Else%><%End If%>" >&nbsp;&nbsp;<input type="image" alt="בחר מרשימה" src="../../images/icon_find.gif" onclick="return openCompaniesList();"
														align="absmiddle">&nbsp;&nbsp;<input type="image" alt="בטל בחירה" onclick="return removeCompany();" src="../../images/delete_icon.gif"
														align="absmiddle">
												</td>
												<td align="<%=align_var%>" nowrap class="form_title"><!--קישור ל--><%=arrTitles(7)%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
											</tr>
											<%End If%>
											<tbody name="contacter_body" id="contacter_body" <%If (trim(companyID) = "" Or IsNull(companyID)) Then%> style="display:none" <%End if%>>
												<tr>
													<td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>"><input type="hidden" name="contactID" id="contactID" value="<%=contactID%>">
														<input type="text" id="contactName" name="contactName" dir="<%=dir_obj_var%>" style="width: 350px" class="Form_R" readonly 
			value="<%If Len(contactName) > 0 Then%><%=vFix(contactName)%><%Else%><%End If%>" 	>&nbsp;&nbsp;<input type="image" alt="בחר מרשימה" src="../../images/icon_find.gif" onclick="return openContactsList();"
															align="absmiddle">&nbsp;&nbsp;<input type="image" alt="בטל בחירה" onclick="return removeContact();" src="../../images/delete_icon.gif"
															align="absmiddle">
													</td>
													<td align="<%=align_var%>" nowrap class="form_title"><!--קישור ל--><%=arrTitles(8)%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td>
												</tr>
											</tbody>
											<tbody name="project_body" id="project_body" <%If trim(companyID) = "" Or IsNull(companyID) Then%> style="display:none" <%End if%>>
												<tr>
													<td align="<%=align_var%>" width="100%" dir="<%=dir_obj_var%>"><input type="hidden" name="project_id" id="project_id" value="<%=projectId%>">
														<input type="text" id="projectName" name="projectName" dir="<%=dir_obj_var%>" style="width: 350px" class="Form_R" readonly 
				value="<%If Len(projectName) > 0 Then%><%=vFix(projectName)%><%Else%><%End If%>" 	>&nbsp;&nbsp;<input type="image" alt="בחר מרשימה" src="../../images/icon_find.gif" onclick="return openProjectsList();"
															align="absmiddle">&nbsp;&nbsp;<input type="image" alt="בטל בחירה" onclick="return removeProject();" src="../../images/delete_icon.gif"
															align="absmiddle">
													</td>
													<td align="<%=align_var%>" nowrap class="form_title"><!--קישור ל--><%=arrTitles(9)%><%=trim(Request.Cookies("bizpegasus")("Projectone"))	%></td>
												</tr>
											</tbody>
											<tbody name="mechanism_body" id="mechanism_body" <%If trim(projectID) = "" Or IsNull(projectID) Then%> style="display:none" <%End if%>>
												<TR>
													<td align="<%=align_var%>" width="100%">
														<select name="mechanism_id" id="mechanism_id" class="norm" dir="<%=dir_obj_var%>" style="width:350px">
															<%If trim(lang_id) = "1" Then%>
															<option value=""><%=String(23,"-")%>
																לא מקושר למנגנון
																<%=String(23,"-")%>
															</option>
															<%Else%>
															<option value=""><%=String(23,"-")%>
																Not linked to Sub-Project
																<%=String(23,"-")%>
															</option>
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
													</td>
													<td align="<%=align_var%>" nowrap class="form_title"><!--קישור ל--><%=arrTitles(9)%><%=arrTitles(38)%></td>
												</TR>
											</tbody>
											<%If quest_id = 16735 Then 'הוסף טופס רישום חתום%>
											<%''if true then%>
											
											<%''end if 'if true%>
											<tr>
												<td align="<%=align_var%>" width="100%">
													<select name="UserIdOrderOwner"  dir="<%=dir_obj_var%>" class="norm" style="width: 250px" ID="UserIdOrderOwner">
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
												<td align="<%=align_var%>" width=130 nowrap class="links_down"><b>למי שייך הרישום</b></td>
											</tr>
											<%End If%>
											<tr>
												<td width="100%" height="10" colspan="2"></td>
											</tr>
										</table>
									</td>
								</tr>
								<%Else%>
								<%If isNumeric(companyId) And trim(private_flag) = "0" Then%>
								<tr>
									<td align="<%=align_var%>" width="100%">
										<%If trim(COMPANIES) = "1" Then%>
										<A href=# onclick='window.opener.location.href="../companies/company.asp?companyID=<%=companyId%>"' target=_self class="link1" dir=rtl>
											<%End If%>
											<%=companyName%>
											<%If trim(COMPANIES) = "1" Then%>
										</A>
										<%End If%>
									</td>
									<td align="<%=align_var%>" width=130 nowrap class="form_title"><!--קישור ל--><%=arrTitles(9)%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
								</tr>
								<%End If%>
								<%If isNumeric(contactId) Then%>
								<tr>
									<td align="<%=align_var%>" width="100%">
										<%If trim(COMPANIES) = "1" Then%>
										<A href=# onclick='window.opener.location.href="../companies/contact.asp?contactID=<%=contactID%>"' target=_self class="link1" dir=rtl>
											<%End If%>
											<%=contactName%>
											<%If trim(COMPANIES) = "1" And trim(private_flag) = "0" Then%>
										</A>
										<%End If%>
									</td>
									<td align="<%=align_var%>" width=130 nowrap class="form_title"><!--קישור ל--><%=arrTitles(9)%><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td>
								</tr>
								<%End If%>
								<%If isNumeric(projectID) And trim(projectName) <> "" Then%>
								<tr>
									<td align="<%=align_var%>" width="100%">
										<%If trim(COMPANIES) = "1" Then%>
										<A href=# onclick='window.opener.location.href="../projects/project.asp?companyID=<%=companyId%>&projectID=<%=projectID%>"' target=_self class="link1" dir=rtl>
											<%End If%>
											<%=projectName%>
											<%If trim(COMPANIES) = "1" Then%>
										</A>
										<%End If%>
									</td>
									<td align="<%=align_var%>" width=130 nowrap class="form_title"><!--קישור ל--><%=arrTitles(9)%><%=trim(Request.Cookies("bizpegasus")("Projectone"))%></td>
								</tr>
								<%End If%>
								<%If trim(mechanismID) <> "" Then%>
								<tr>
									<td align="<%=align_var%>" dir="<%=dir_obj_var%>" width="100%">
										<%If trim(COMPANIES) = "1" Then%>
										<A href="../projects/project.asp?companyID=<%=companyId%>&projectID=<%=projectID%>" target=_self class="links_down">
											<%End If%>
											<%=mechanismName%>
											<%If trim(COMPANIES) = "1" Then%>
										</A>
										<%End If%>
									</td>
									<td align="<%=align_var%>" width=130 nowrap class="links_down"><b><!--קישור ל--><%=arrTitles(14)%><%=arrTitles(35)%></b></td>
								</tr>
								<%End If%>
								<%End If%>
							</table>
						</td>
					</tr>
					<!--start adding attachment -->
<%if quest_id=17857 then%>
<tr>
<td align="<%=align_var%>" dir="<%=dir_obj_var%>" width="100%">   
			<table cellpadding=2 cellspacing=0 border=0 width=450 align="<%=align_var%>" ID="Table3">
			<tr>
<td align="center" dir="<%=dir_obj_var%>" width="100%">   
     <SELECT dir="<%=dir_obj_var%>" size=1 ID=Tour_id name=Tour_id class="sel" style="width:310px;font-size:10pt"   onchange="selectToursDate(this)">	
<%If trim(lang_id) = "1" Then%>
 <OPTION value="">בחר טיול</OPTION>
<%Else%>
 <OPTION value="">Choose Tour</OPTION>
<%End If%> 
<%
sqlstrTour = "SELECT Tour_Id, (Category_Name + ' - ' + Tour_Name) as Tour_Name  FROM dbo.Tours T " & _
        " INNER JOIN dbo.Tours_Categories TC ON T.Category_Id = TC.Category_Id  " & _
        " INNER JOIN dbo.Tours_SubCategories TSC ON T.SubCategory_Id = TSC.SubCategory_Id " & _
         " ORDER BY Category_Order, Tour_Name"
 ''       " WHERE (Tour_Vis = 1) AND (SubCategory_Vis = 1) AND (Category_Vis = 1) " & _
   

	'Response.Write sqlstr
	'Response.End
	set rs_tours= conPegasus.Execute(sqlstrTour)
   do while not rs_tours.EOF 
		Tour_Id =rs_tours("Tour_Id")
		Tour_name =rs_tours("Tour_Name")%>

	<OPTION value="<%=Tour_Id%>" <%If trim(pTourId) = trim(Tour_Id) Then%> selected <%End If%>><%=Tour_name%> </OPTION>
			<%rs_tours.moveNext
		loop
		rs_tours.close
		set rs_tours=Nothing	
%>
</SELECT>
	</TD>
				</TR>	
			<tr><td height=10></td></tr>	
				<TR>
			<tr><td align="center" dir="<%=dir_obj_var%>" width="100%">   

			
				<select name="ToursDate_Id" id="ToursDate_Id" class="sel" dir=rtl  style="width:310px;font-size:10pt;display:block;" onchange="selectGuide(this)">
				<option value="">בחר תאריך יציאה</option>
				</select>
				</TD>
						

				</TR>
							<tr><td height=10></td></tr>	
			
			<tr><td align="center" dir="<%=dir_obj_var%>" width="100%">   

			
				<select name="Guide_Id" id="Guide_Id" class="sel" dir=rtl  style="width:310px;font-size:10pt;display:block;">
				<%If trim(lang_id) = "1" Then%>
 <OPTION value="">בחר מדריך</OPTION>
<%Else%>
 <OPTION value="">Choose Gוuide</OPTION>
<%End If%> 
				       <%
sqlstrGuide = "SELECT Guide_Id, (Guide_FName + ' - ' + Guide_LName) as Guide_Name  FROM Guides  where Guide_Vis=1 ORDER BY Guide_LName, Guide_FName"
    

	'Response.Write sqlstr
	'Response.End
	set rs_Guide= conPegasus.Execute(sqlstrGuide)
   do while not rs_Guide.EOF 
		Guide_Id =rs_Guide("Guide_Id")
		Guide_Name =rs_Guide("Guide_Name")%>

	<OPTION value="<%=Guide_Id%>" <%If trim(pGuideId) = trim(Guide_Id) Then%> selected <%End If%>><%=Guide_Name%> </OPTION>
			<%rs_Guide.moveNext
		loop
		rs_Guide.close
		set rs_Guide=Nothing	
%>
				<option value="">בחר מדריך</option>
				</select>
				</TD>
	

				</TR>
						<tr>
				<td align="<%=align_var%>" width="100%" style="padding-right:130px" class="card">
				<input dir="<%=dir_obj_var%>" type="file" name="f1" id="f1" style="width:340" maxlength=100 >
				</td>
			</tr>
			<tr><td height=10></td></tr>	
		
				<tr><td height=10></td></tr>	</table></td></tr>
<%end if%>
			<tr <%if 	quest_id="17857" then%>style="display:none"<%end if%>>
						<td bgcolor="#C9C9C9" style="border-bottom: 1px solid #808080;padding-right:15px;padding-left:15px">
							<table cellpadding="0" cellspacing="0" width="100%" border="0">
								<tr>
									<td height="5" colspan="2" nowrap></td>
								</tr>
								<tr>
									<td align="<%=align_var%>" width="100%">
										<input dir="<%=dir_obj_var%>" type=text class="Form" name="document_name" id="document_name" maxlength=100 style="width:300">
									</td>
									<td align="<%=align_var%>" width="150" nowrap class="form_title">כותרת קובץ</td>
								</tr>
								<tr>
									<td align="<%=align_var%>" width="100%">
										<textarea dir="<%=dir_obj_var%>" class="Form" name="document_desc" id="document_desc" maxlength="255" style="width: 300px"></textarea>
									</td>
									<td align="<%=align_var%>" width="150" nowrap class="form_title">תאור קובץ</td>
								</tr>
								<tr>
									<td align="<%=align_var%>" dir="ltr" width="100%">
										<input dir="<%=dir_obj_var%>" type="file" name="document_file" style="width: 300px" maxlength="100" ID="document_file">
									</td>
									<td align="<%=align_var%>" width="150" nowrap class="form_title">קובץ מצורף</td>
								</tr>
								<tr>
									<td height="5" colspan="2" nowrap></td>
								</tr>
							</table>
						</td>
					</tr>
		
					<!--end adding attachment -->
					<%If quest_id = 16735 Then%>
<%if Is_GroupsTour=1 then%>
<%else%>				
<tr>	<td width="100%" colspan="2"width="100%" colspan="2" class="title_form" bgcolor="#DADADA" align="right" valign="top"
							dir="ltr" style="padding-right:15px;">שיוך טופס מתעניין 
													בטיול</td></tr>
<tr>
												<td align="right" width="100%" dir="<%=dir_obj_var%>" colspan=2>
													<input type="hidden" name="RelationId" id="RelationId" value="<%=RelationId%>"> <input type="text" id="RelationName" name="RelationName" dir="<%=dir_obj_var%>" style="width: 350px" class="Form_R" readonly 
			value="<%If Len(RelationName) > 0 Then%><%=vFix(RelationName)%><%Else%><%End If%>" >&nbsp;&nbsp;<input type="image" alt="בחר מרשימה" src="../../images/icon_find.gif" onclick="return openRelationList();"
														align="absmiddle" ID="RelationImage1" NAME="RelationImage1">&nbsp;&nbsp;<input type="image" alt="בטל בחירה" onclick="return removeRelation();" src="../../images/delete_icon.gif"
														align="absmiddle" ID="RelationImage2" NAME="RelationImage2">
												</td>
											
											</tr>
											<%end if%>
											<%end if%>
					<%if quest_id=16504 then%>
					<tr>
						<td colspan="2" height="2" width="100%" style="bgcolor:#808080"></td>
					</tr>
					<tr>
						<td width="100%" colspan="2" class="title_form" bgcolor="#DADADA" align="right" valign="top"
							dir="ltr" style="padding-right:15px;"><span dir="rtl"><b>שאלות חובה במהלך השיחה המופנות 
									ללקוח מתעניין / מתעניין</b></span>
						</td>
					</tr>
					<tr>
						<td width="100%" class="form" bgcolor="#DADADA" align="right" valign="top" dir="ltr"
							style="padding-right:15px;"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>לאיזה 
									יעד נסיעה הינך מתעניין?</b></span>
						</td>
					</tr>
					<tr>
						<td width="100%" align="right">
							<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#F0F0F0" ID="Table2">
								<tr>
									<td align="right" width="100%" dir="<%=dir_obj_var%>" style="padding-right:15px;">
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
										</select>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					<%end if%>
				
					<tr>
						<td width="100%" align="right">
							<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="White" dir="<%=dir_obj_var%>">
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
					<tr>
						<td width="100%" height="10" colspan="4"></td>
					</tr>
					<tr>
						<td colspan="4" height="10" align="<%=td_align%>" dir="<%=dir_obj_var%>">&nbsp;<font color="red"><span id="word19" name="word19"><!--שדות חובה--><%=arrTitles(19)%>&nbsp;-&nbsp;*</span></font>&nbsp;</td>
					</tr>
					<%end if%>
					<tr>
						<td width="100%" height="10" colspan="4"></td>
					</tr>
					<tr>
						<td align="center" colspan="4"><INPUT style="width:90px" type="submit" class="but_menu" value="<%=arrButtons(4)%>" 
	id="btnSubmit" name="btnSubmit" onClick="return CheckFields()">
						</td>
					</tr>
					<tr>
						<td width="100%" height="10" colspan="4"></td>
					</tr>
				</table>
			</form>
			</td></tr>
			<%End If%>
		</div>
		</table>
		<script language="javascript" type="text/javascript">
<!-- 
    <%	If trim(lang_id) = "1" Then
			str_alert = "!!!נא למלא את השדה"
		Else
			str_alert = "Please provide the answer!!!"
		End If	
    %> 		 	
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
	</body>
</html>
<%set con=nothing%>

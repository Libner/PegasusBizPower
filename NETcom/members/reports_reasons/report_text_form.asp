<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%  Session.LCID = 1037
	quest_id = Request.Form("quest_id")
	fId = Request.Form("FieldId")
	start_date = DateValue(trim(Request("dateStart")))
    end_date = DateValue(trim(Request("dateEnd")))
    start_date_ = Month(start_date) & "/" & Day(start_date) & "/" & Year(start_date)   
    end_date_ = Month(end_date) & "/" & Day(end_date) & "/" & Year(end_date)
    CompanyID = trim(Request("company_id"))
    ProjectID = trim(Request("project_id"))
    mechanismID = trim(Request("mechanism_id"))
    
	reasons=""
	'only for tofes mitan'en
	
			'הגורם ליצירת הטופס
			'---------------P2932 - type of form- REASONS ---------------
				'select creation reason-------------------------------
				'added by Mila 22/10/2019-----------------------------			
	if quest_id<>"" then
		if cint(quest_id)=16504 then
			reasons=Request.Form("reasons")
			reasons=replace(reasons," ","")
		end if
	end if
	'-----------------------------------------------------------
	
    OrgID	 = trim(Request.Cookies("bizpegasus")("OrgID"))
  	User_ID	 = trim(Request.Cookies("bizpegasus")("UserID"))
  	lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))
  	is_groups = trim(Request.Cookies("bizpegasus")("ISGROUPS"))
  	UserID = trim(Request("user_id"))
	If trim(UserID) <> "" Then
		sqlstr = "Select FIRSTNAME + CHAR(32) + LASTNAME FROM USERS WHERE USER_ID = " & UserID
		set rs_user = con.getRecordSet(sqlstr)
		If not rs_user.eof Then
			userName = trim(rs_user(0))
		End If
		set rs_user = Nothing
	End If
	
	strTitle = ""
	If trim(UserID) <> "" Then
		sqlstr = "Select FirstName, LastName From Users WHERE User_ID = " & UserID
		set rs_user = con.getRecordSet(sqlstr)
		if not rs_user.eof then
			strTitle = strTitle & "<br>&nbsp;עובד:&nbsp;<font color=""#666699"">" & trim(rs_user(0)) & " " & trim(rs_user(1))& "</font>"
		end if
		set rs_user = nothing
	End If

	If trim(CompanyID) <> "" Then
		sqlstr = "Select Company_Name From Companies WHERE Company_ID = " & CompanyID
		set rs_com = con.getRecordSet(sqlstr)
		if not rs_com.eof then
			strTitle = strTitle & "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & ":&nbsp;<font color=""#666699"">" & trim(rs_com(0)) & "</font>"
		end if
		set rs_com = nothing
	End If

	If trim(ProjectID) <> "" Then
		sqlstr = "Select Project_Name From Projects WHERE Project_ID = " & ProjectID
		set rs_project = con.getRecordSet(sqlstr)
		if not rs_project.eof then
			strTitle = strTitle & "<br>&nbsp; " & trim(Request.Cookies("bizpegasus")("Projectone")) & ":&nbsp;<font color=""#666699"">" & trim(rs_project(0)) & "</font>"
		end if
		set rs_project = nothing
	End If

	If trim(mechanismID) <> "" Then
		sqlstr = "Select mechanism_name, project_name From mechanism Inner Join Projects On mechanism.project_id = projects.project_id WHERE mechanism_id = " & mechanismID
		set rs_mech = con.getRecordSet(sqlstr)
		if not rs_mech.eof then
			strTitle = strTitle & "<br>&nbsp; מנגנון:&nbsp;<font color=""#666699"">" & trim(rs_mech(1)) & " - " & trim(rs_mech(0)) & "</font>"
		end if
		set rs_mech = nothing
	End If	
	
	'---------------P2932 - type of form- REASONS ---------------
		If reasons<>"" Then
		sqlstr = "SELECT  Reason_Title FROM  Appeals_CreationReasons where QUESTIONS_ID=16504 and Reason_Id in (" & reasons & ") order by Reason_Order"
			set rs_Reason = con.getRecordSet(sqlstr)
			If Not rs_Reason.eof Then
				arr_Reason = rs_Reason.getRows()
			End If
			Set rs_Reason = Nothing
			If isArray(arr_Reason) Then
			
				strTitle = strTitle & "<br>&nbsp; הגורם ליצירת הטופס:&nbsp;"
				'response.Write(Ubound(arr_Reason,2))
				'response.end
				For mm=0 To Ubound(arr_Reason,2)
					if mm>0 then
						strTitle = strTitle & "<font color=""#FF9900"">,&nbsp;</font>"
					end if
					strTitle = strTitle & "<font color=""#FF9900"">" & trim(arr_Reason(0,mm)) & "</font>"
				next
			end if
	end if
		 
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function showcard(Id,gID)
	{
		strOpen="../groups_clients/addpeople.asp?people_Id="+Id+"&groupId="+gID;
		oClient = window.open(strOpen, "Client", "scrollbars=1,toolbar=0,top=50,left=150,width=450,height=350,align=center,resizable=0");
		oClient.focus();
		return false;
	}
//-->
</SCRIPT>
</head>
<body marginwidth="0" marginheight="0" hspace="0" vspace="0" topmargin="0" leftmargin="0" rightmargin="0" >
<div align="center">
<center>
<table border="0" width="780" cellspacing="0" cellpadding="0" align=center>
<tr>
<td width="780" align="center">
<!--#INCLUDE FILE="logo_top.asp"-->
</td></tr> 
</table>
<table border="0" width="780" cellspacing="0" cellpadding="0" align=center>
<!-- start code --> 
	<tr bgcolor="#FF9900">
		<td align="center" height="20" style="color:#000000; font-size:13pt" align=right valign=bottom dir="rtl"><b>			
         <%If trim(lang_id) = "1" Then%>דוח תשובות לשאלה חופשית<%Else%>Open text - Answer report<%End If%><%=strTitle%></b></td>
	</tr>				

  <tr>
    <td width="100%" align="center">
	<table WIDTH="100%" BGCOLOR="#999999" ALIGN="center" BORDER="0" CELLPADDING="1" cellspacing="0">	
			<tr>
			<td align="right" width="100%" valign="top">
<%				set prod = con.GetRecordSet("Select Product_Name, Langu FROM PRODUCTS WHERE PRODUCT_ID=" & quest_id & " AND ORGANIZATION_ID=" & orgID)
					if not prod.eof then
					productName=prod("Product_Name")					
					if prod("Langu") = "eng" then
						td_align = "left"
						pr_language = "eng"
					else
						td_align = "right"
						pr_language = "heb"
					end if
				end if
				set prod = nothing%> 			 
			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="right" bgcolor="#F0F0F0">
			<tr><td class="form" align=center width=100% bgcolor=#999999 colspan=4><INPUT type="hidden" id=product_id name=product_id value="<%=prod_id%>">
					<table width="100%" cellpadding="0" cellspacing="0">
						<tr>
							<td align="center">
							<font color=#FFFFFF size=3 dir="rtl"><b><%=productName%></b></font>
							</td>
						</tr>
						<tr>
						   <td align="center" dir=ltr><font color=#FFFFFF size=3 dir="rtl"><b><%=end_date & " - " & start_date%></b></font></td>
						</tr>
							<%'---------------P2932 - type of form- REASONS ---------------
   	If reasons<>"" Then
		sqlstr = "SELECT  Reason_Title FROM  Appeals_CreationReasons where QUESTIONS_ID=16504 and Reason_Id in (" & reasons & ") order by Reason_Order"
			set rs_Reason = con.getRecordSet(sqlstr)
			If Not rs_Reason.eof Then
				arr_Reason = rs_Reason.getRows()
			End If
			Set rs_Reason = Nothing
			If isArray(arr_Reason) Then
			%>
			<tr ><td align="center" dir=ltr><font color=#FFFFFF size=3 dir="rtl"><b> הגורם ליצירת הטופס:
 				<%For mm=0 To Ubound(arr_Reason,2)
					if mm>0 then%>
					,&nbsp;
					<%end if%>
					<%= trim(arr_Reason(0,mm))%>
				<%next%>
				</b></font></td></tr>
<%
			end if
	end if
	%>
					</table>		
				</td>
		   </tr>
		
	<%
	sqlStr = "SET DATEFORMAT DMY; "&_
	" SELECT FORM_VALUE.FIELD_ID, FORM_VALUE.FIELD_VALUE, FORM_FIELD.FIELD_TYPE, FORM_FIELD.FIELD_SIZE,"&_ 
    " FORM_FIELD.FIELD_ALIGN, FORM_FIELD.FIELD_TITLE, APPEALS.APPEAL_DATE, APPEALS.APPEAL_ID, COMPANY_NAME"&_
    " FROM APPEALS LEFT OUTER JOIN COMPANIES ON APPEALS.COMPANY_ID = COMPANIES.COMPANY_ID INNER JOIN "&_
    " FORM_VALUE ON APPEALS.APPEAL_ID = FORM_VALUE.APPEAL_ID INNER JOIN "&_
    " FORM_FIELD ON FORM_VALUE.FIELD_ID = FORM_FIELD.FIELD_ID INNER JOIN "&_
	" PRODUCTS ON APPEALS.QUESTIONS_ID = PRODUCTS.PRODUCT_ID " &_
	" AND FORM_VALUE.FIELD_ID = " & fId & " AND APPEALS.QUESTIONS_ID = " & quest_id &_
	" AND APPEALS.ORGANIZATION_ID = " & OrgID
	If IsNumeric(UserID) Then
		  sqlStr = sqlStr & " AND APPEALS.USER_ID = " & UserID
	End If  
	If IsNumeric(CompanyID) Then
	  sqlStr = sqlStr & " AND APPEALS.COMPANY_ID = " & CompanyID
	End If 
	If IsNumeric(ProjectID) Then
		  sqlStr = sqlStr & " AND APPEALS.Project_ID = " & ProjectID
	End If
	If IsNumeric(mechanismID) Then
		  sqlStr = sqlStr & " AND APPEALS.Mechanism_ID = " & mechanismID
	End If	 
	If isDate(start_date) Then
		  sqlStr = sqlStr & " And DateDiff(d,APPEAL_DATE,'"&start_date&"') <= 0"
	End If	
	If isDate(end_date) Then
		 sqlStr = sqlStr & " And DateDiff(d,APPEAL_DATE,'"&end_date&"') >= 0"			
	End If 
		'---------------P2932 - type of form- REASONS ---------------
	If reasons<>"" Then 	
		  sqlStr = sqlStr & " AND APPEALS.Reason_Id in (" & reasons & ")"
	end if 
	
	If trim(is_groups) = "1" Then	
		sqlStr = sqlStr & " AND ((appeals.User_Id IN (Select User_ID From Users_To_Groups WHERE Group_ID IN (Select Group_ID From Responsibles_to_Groups WHERE Responsible_ID = " & User_ID & "))"&_
 		" AND (appeals.Questions_Id IN (Select Product_ID From Users_To_Products WHERE User_ID = " & User_ID & "))" &_
 		" Or (appeals.User_Id  = " & User_ID & ")" &_				 		
 		" Or ((appeals.questions_id IN (Select product_id From Users_To_Products WHERE Group_ID IN (Select Group_ID From Responsibles_to_Groups WHERE Responsible_ID = " & User_ID & "))" &_
 		" And products.OPEN_FORM = 1 And appeals.User_Id IS NULL))"	&_ 		
 		" Or (appeals.questions_id IN (Select product_id From Users_To_Products WHERE Group_ID IN (Select Group_ID From Users_to_Groups WHERE User_ID = " & User_ID & "))" &_
 		" And products.OPEN_FORM = 1 And appeals.User_Id IS NULL))" &_ 		
 		" Or (Products.Responsible = " & User_ID & "))"
	End If	
    sqlStr = sqlStr & " Order BY APPEAL_DATE"
	'Response.Write sqlStr
	'Response.End	 	    
	set fld = con.GetRecordSet(sqlStr)
	if not fld.eof then %>
	<tr>
		<td>
		<TABLE align=center BORDER=0 CELLSPACING=0 CELLPADDING=0 width=100%>
		<tr bgcolor=#DADADA dir="rtl" style="font-size:10pt;font-weight:bold;">			
			<%if pr_language = "heb" then%>
			    <td width="15%" align=center valign="top" nowrap style="padding:3px">מספר טופס מלא</td>
			    <td width="15%" align=center valign="top" nowrap style="padding:3px">תאריך</td>
			    <td width="25%" align=right valign="top" nowrap style="padding:3px"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
				<td width="45%" align=right valign="top" dir="rtl" style="padding:3px"><%=fld("Field_Title")%></td>
			<%else%>	
				<td width="45%" align=left valign="top" dir="rtl" style="padding:3px"><%=fld("Field_Title")%></td>
				<td width="25%" align=left valign="top" nowrap style="padding:3px"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
				<td width="15%" align=center valign="top" nowrap style="padding:3px">Date</td>
				<td width="15%" align=center valign="top" nowrap style="padding:3px">Form No</td>
			<%end if%>
		</tr>	
		</TABLE>
		</td>
	</tr>
	 <% 
	 Do while not fld.eof
		Field_Value=fld("Field_Value")
		FIELD_TYPE=fld("FIELD_TYPE")
		if FIELD_TYPE = "5" then
			if pr_language = "heb"  then 
				Field_Value = "כן"
			else
				Field_Value = "Yes"			
			end if	
		end if
		Field_Align=fld("Field_Align")
		appID=fld("APPEAL_ID")
		appDate=fld("APPEAL_DATE")	  
		If IsDate(appDate) Then
			appDate = Day(appDate) & "/" & Month(appDate) & "/" & Year(appDate)
		End If
		compName=fld("COMPANY_NAME")
		If trim(Field_Value) <> "" Then	  
	%>
		<tr>
			   <td width="100%" height="1" nowrap bgcolor=#DADADA></td>
		</tr>
		<tr>
		  <td width="100%">
		  <TABLE align=center BORDER=0 CELLSPACING=1 CELLPADDING=2 width=100%>
			<tr>
			<%if pr_language = "heb" then%>
			    <td class="form" width="15%" align=center valign="top" style="padding:3px"><%=appID%></a></td>
			    <td class="form" width="15%" align=center valign="top" style="padding:3px"><%=appDate%></a></td>
			    <td class="form" width="25%" align=right valign="top" style="padding:3px"><%=compName%></a></td>
				<td class="form" width="45%" align=right valign="top" dir="rtl" style="padding:3px"><%=Field_Value%></td>
			<%else%>	
				<td class="form" width="45%" align=left valign="top" dir="rtl" style="padding:3px"><%=Field_Value%></td>
				<td class="form" width="25%" align=left valign="top" style="padding:3px"><%=compName%></a></td>
				<td class="form" width="15%" align=center valign="top" style="padding:3px"><%=appDate%></td>
				<td class="form" width="15%" align=center valign="top" style="padding:3px"><%=appID%></td>
			<%end if%>
	         </tr>			
	      </table>      				
		</tr>
	<%
		end if
	fld.MoveNext
	loop
	end if 'not fld.eof
	set fld = nothing%>
<!-- end code --> 
</table></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
<!--#INCLUDE FILE="bottom_inc.asp"-->
</table>
</center>
</div>
</body>
<%  set con = nothing %>
</html>

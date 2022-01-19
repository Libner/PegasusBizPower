<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<% 
PathCalImage = "../../"	
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<%
	appId = trim(Request("appid"))
	prodId = trim(Request("prodId"))
	quest_id = trim(Request("quest_id"))
	if trim(lang_id) = "1" then
		 arr_Status = Array("","חדש","בטיפול","סגור")	
	else
		 arr_Status = Array("","new","active","close")	
	end if
    found_feedback = false
    sqlstr = "EXECUTE get_feedbacks '','','','','" & OrgID & "','','','','','','','" & appID & "'"
    'Response.Write("sqlstr=" & sqlstr)
	set app = con.GetRecordSet(sqlstr)
	if not app.eof then
		appeal_status = app("appeal_status")
		quest_id = app("QUESTIONS_ID") 
		in_date = Day(app("appeal_date"))& "/" & Month(app("appeal_date")) & "/" & Year(app("appeal_date"))	
		product_name = trim(app("product_name"))
		peopleId = app("people_id")
		found_feedback = true
		if trim(peopleId)<>"" then	    
			set people=con.GetRecordSet("select * from PEOPLES where PEOPLE_ID=" & peopleId)
			if not people.eof then
			peopleName=trim(people("PEOPLE_NAME"))				
			peoplePHONE=trim(people("PEOPLE_PHONE"))
			peopleFAX=trim(people("PEOPLE_FAX"))
			peopleCELL = trim(people("PEOPLE_CELL"))
			peopleEMAIL=trim(people("PEOPLE_EMAIL"))
			peopleCOMPANY=trim(people("PEOPLE_COMPANY"))
			peopleOFFICE=(trim(people("PEOPLE_OFFICE")))	
			contactID=(trim(people("CONTACT_ID")))	
			If Len(contactID) > 0 Then
				sqlstr = "Select COMPANY_NAME, CONTACTS.COMPANY_ID, CONTACT_NAME, private_flag FROM contacts INNER JOIN COMPANIES ON CONTACTS.COMPANY_ID = COMPANIES.COMPANY_ID WHERE CONTACTS.CONTACT_ID = " & contactID
				set rs_cont = con.getRecordSet(sqlstr)
				If not rs_cont.eof Then
					companyId = trim(rs_cont("COMPANY_ID"))
					companyName = trim(rs_cont("COMPANY_NAME"))
					contactName = trim(rs_cont("CONTACT_NAME"))
					private_flag = trim(rs_cont("private_flag"))
				End If
				set rs_cont = Nothing
			End If
		end if
		people.close
		set people=nothing		
	end if
	product_id = app("product_id")
	sqlstr =  "Select Langu,RESPONSIBLE From Products WHERE PRODUCT_ID = " & quest_id
	'Response.Write sqlstr
	'Response.End
	set rsq = con.getRecordSet(sqlstr)
	If not rsq.eof Then	
		Langu = trim(rsq(0))
		RESPONSIBLE = trim(rsq(1))
	End If
	set rsq = Nothing	
	if Langu = "eng" then
		td_align = "left"
		pr_language = "eng"
	else
		td_align = "right"
		pr_language = "heb"
	end if
	set app = nothing
	
	If trim(appeal_status) = "1" Then
		sqlstring="UPDATE appeals set appeal_status = '2' WHERE ORGANIZATION_ID = " & OrgID & " AND appeal_id = " & appid 	
		'Response.Write(sqlstring)
		'Response.End 
		con.ExecuteQuery(sqlstring) 		
	End If	
	
	If trim(UserID) = trim(RESPONSIBLE)	Then
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
	
	if Request.Form("updseker") <> nil and appid <> nil and quest_id <> nil then
	    
	    appeal_status=trim(Request.Form("appeal_status"))

	    sqlstring="UPDATE appeals set appeal_status = '" & appeal_status & "' WHERE ORGANIZATION_ID = " & OrgID & " AND appeal_id = " & appid 	
		'Response.Write(sqlstring)
		'Response.End 
		con.ExecuteQuery(sqlstring)     
		
		txt_exception = ""
		set fields=con.GetRecordSet("SELECT Field_Id,Field_Title,Field_Type FROM FORM_FIELD Where product_id=" & quest_id & " Order by Field_Order")
			do while not fields.EOF
				Field_Id = fields("Field_Id")
				'Field_questionID = fields("QUESTION_ID")
				importance = Request.Form("importance" & Field_Id)
				rvalue = Request.Form("field" & Field_Id)
				
				if fields("FIELD_TYPE") = "4" then
					Field_Value = rvalue & " " & importance
				else
					Field_Value = trim(rvalue)
				end if
				
				sqlstr="UPDATE FORM_VALUE Set Field_Value='"& Left(sFix(Field_Value),255) &"' Where Appeal_Id=" & appid & " and Field_Id=" & Field_Id 
				'Response.Write(sqlstr)
				con.ExecuteQuery sqlstr
				if fields("FIELD_TYPE") = "2" and Len(Field_Value) > 255 then  ' long text
					sqlstr="UPDATE FORM_TEXT Set Field_Value='"& sFix(Field_Value) &"'  Where Appeal_Id=" & appid & " and Field_Id=" & Field_Id
					'Response.Write(sqlstr)
					con.ExecuteQuery sqlstr	
				end if
				' add new tasks if is an exception
			fields.moveNext()
			loop
		set fields=nothing
		Response.Redirect "feedback.asp?quest_id="&quest_id&"&appid="&appid
	end if 'Request.Form("updseker") <> nil
	
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
	  set rsbuttons=nothing	 	  
%>
<SCRIPT LANGUAGE=javascript>
<!--
// ----------- calendar's functions -----------
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
// ----------- end calendar's functions -----------
function showcard(Id)
  {
    strOpen="../products/showcard.asp?peopleId="+Id;
	window.open(strOpen, "people", "scrollbars=1,toolbar=0,top=50,left=170,width=450,height=230,align=center,resizable=0");
	return false;
  }
//-->
</SCRIPT>

</head>

<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 2%>
<%numOfLink = 3%>
<!--#include file="../../top_in.asp"-->
<%If found_feedback Then%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<tr><td class="page_title" dir=rtl>&nbsp;<%=product_name%></td></tr>		        
   <tr>
		<td width="100%" colspan="4" align="<%=align_var%>" valign=top>
		<form id="form1" name="form1" action="feedback.asp?arch=<%=Request("arch")%>&quest_id=<%=quest_id%>&peopleId=<%=Request("peopleId")%>&prodId=<%=prodId%>&appid=<%=appid%>" method="POST">
			<TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=0 bgcolor="#e6e6e6" bgcolor="#e6e6e6" valign="top" style="border: 1px solid #808080;border-top:none"  dir="<%=dir_var%>">
			 <tr>
			<td align="<%=align_var%>" bgcolor="#C9C9C9" style="padding-right:15px;padding-left:15px">
			<INPUT type="hidden" id=updseker name=updseker value="1">
			<table cellpadding=1 cellspacing=0 border=0 width=450 align="<%=align_var%>">
			 <tr><td colspan=2 height=10></td></tr>
			 <tr>
			 <td align="<%=align_var%>" width=100%><input type="text" style="width:70;" value="<%=in_date%>" id="appeal_date" class="Form_R"  name="appeal_date" dir="<%=dir_obj_var%>" ReadOnly></td>
			 <td align="<%=align_var%>" width=120 nowrap class="form_title"><span id=word4 name=word4><!--תאריך מענה--><%=arrTitles(4)%></span></td>   
			</tr>
			<tr>
			 <td align="<%=align_var%>" width=100%>
			  <select id="appeal_status" name="appeal_status" dir="<%=dir_obj_var%>" class="norm">			   
				<%For i=1 To Ubound(arr_Status)%>
			    <option value="<%=i%>" <%if cstr(i) = cstr(appeal_status) then%> selected<%end if%>>&nbsp;<%=arr_Status(i)%>&nbsp;</option>
			    <%Next%>
			    </select>
			 </td>
			 <td align="<%=align_var%>" nowrap class="form_title"><span id="word6" name=word6><!--סטטוס--><%=arrTitles(6)%></span></td>   
		</tr>	
			<%If Len(trim(peopleCOMPANY)) > 0 And IsNull(peopleCOMPANY) = false Then%>
			<tr>
			 <td align="<%=align_var%>" width=100%><input type="text" size=40 value="<%=vFix(peopleCOMPANY)%>" class="Form_R"  name="appeal_date" dir="<%=dir_obj_var%>" ReadOnly></td>
			 <td align="<%=align_var%>" nowrap class="form_title"><span id="word3" name=word3><!--חברה--><%=arrTitles(3)%></span></td>   
			</tr>
			<%End If%>
			<%If Len(peopleName) > 0 Then%>
			<tr>
			 <td align="<%=align_var%>" width=100%><input type="text" size=40 value="<%=vFix(peopleName)%>" class="Form_R"  name="appeal_date" dir="<%=dir_obj_var%>" ReadOnly></td>
			 <td align="<%=align_var%>" nowrap class="form_title"><span id="word5" name=word5><!--נמען--><%=arrTitles(5)%></span></td>   
			</tr>
			<%End if%>
			<%If Len(trim(peopleEMAIL)) > 0 Then%>
			<tr>
				<td align="<%=align_var%>"><input type="text" class="Form_R" size=40 name="peopleEMAIL" value="<%=vfix(peopleEMAIL)%>" maxlength=100></td>
				<td align="<%=align_var%>" nowrap class="form_title">E-mail</td>
			</tr>
			<%End if%>
			<%If Len(peopleOFFICE) > 0 Then%>
			<tr>
				<td align="<%=align_var%>"><input dir="<%=dir_obj_var%>" type="text" class="Form_R" size=20 name="peopleOFFICE" value="<%=vfix(peopleOFFICE)%>"></td>
				<td align="<%=align_var%>" nowrap class="form_title"><span id="word10" name=word10><!--תפקיד--><%=arrTitles(10)%></span></td>
			</tr>
			<%End if%>
			<%If Len(peopleCELL) > 0 Then%>
			<tr>
				<td align="<%=align_var%>"><input type="text" class="Form_R" size=20 name="peopleCell" value="<%=vfix(peopleCELL)%>"></td>
				<td align="<%=align_var%>" nowrap class="form_title"><span id="word11" name=word11><!--טלפון נייד--><%=arrTitles(11)%></span></td>
			</tr>
			<%End if%>
			<%If Len(peoplePHONE) > 0 Then%>	
			<tr>
				<td align="<%=align_var%>"><input type="text" class="Form_R" size=20 name="peoplePHONE" value="<%=vfix(peoplePHONE)%>"></td>
				<td align="<%=align_var%>" nowrap class="form_title"><span id="word12" name=word12><!--טלפון--><%=arrTitles(12)%></span></td>
			</tr>
			<%End if%>
			<%If Len(peopleFAX) > 0 Then%>	
			<tr>
				<td align="<%=align_var%>"><input type="text" class="Form_R" size=20 name="peopleFAX" value="<%=vfix(peopleFAX)%>"></td>
				<td align="<%=align_var%>" nowrap class="form_title"><span id="word13" name=word13><!--פקס--><%=arrTitles(13)%></span></td>
			</tr>
			<%End if%>
			<%If Len(trim(companyName)) > 0 Then%>	
			<tr>
			 <td align="<%=align_var%>" width=100%>
			 <%If trim(COMPANIES) = "1" Then%>
			 <A href="../companies/company.asp?companyID=<%=companyId%>" target=_self class="link1" dir=rtl>
			 <%End If%>
			 <%=companyName%>
			 <%If trim(COMPANIES) = "1" Then%>
			 </a>
			 <%End If%>
			 </td>
			 <td align="<%=align_var%>" width=120 nowrap class="form_title"><span id=word7 name=word7><!--קישור ל--><%=arrTitles(7)%></span><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>   
			</tr>
			<%End If%>
			<%If trim(contactID) <> "" Then%>
			<tr>
			 <td align="<%=align_var%>" width=100%>
			 <%If trim(COMPANIES) = "1" And trim(private_flag) = "0" Then%>
			 <A href="../companies/contact.asp?contactID=<%=contactID%>" target=_self class="link1" dir=rtl>
			 <%End If%>
			 <%=contactName%>
			 <%If trim(COMPANIES) = "1" And trim(private_flag) = "0" Then%>
			 </a>
			 <%End If%>
			 </td>
			 <td align="<%=align_var%>" width=120 nowrap class="form_title"><span id=word8 name=word8><!--קישור ל--><%=arrTitles(8)%></span><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td>   
			</tr>
			<%End If%>
			<%End If%>					
			<tr><td height=10 nowrap colspan=2></td></tr>
			</table></td></tr>
			<tr>
		    <td width="100%" align="<%=align_var%>">
			<TABLE WIDTH=100% BORDER=0 CELLSPACING=1 CELLPADDING=3 bgcolor=White>	
			<!-- start fields dynamics -->	
			<%If trim(appId) = "" Then%>			
			<!--#INCLUDE FILE="default_fields.asp"-->
			<%else%>
			<!--#INCLUDE FILE="upd_fields.asp"-->
			<%End If%>
			<!-- end fields dynamics -->
			</table>
	</td></tr>
	<%if is_must_fields then %>
	<tr><td width="100%" height="10" colspan=4></td></tr>
	<tr><td colspan="4" height="10" align="<%=td_align%>" dir="<%=dir_obj_var%>">&nbsp;<font color=red><span id=word19 name=word19><!--שדות חובה--><%=arrTitles(19)%>&nbsp;-&nbsp;*</span></font>&nbsp;</td></tr>
	<%end if%>				
	<tr><td width="100%" height="10" colspan=4></td></tr>
	<tr><td align=center colspan=4>
			<TABLE ALIGN=center BORDER=0 CELLSPACING=1 CELLPADDING=1>
				<TR>
				    <td align="center" nowrap width=45% class="title_table_admin"><INPUT class="but_menu" type="button" value="<%=arrButtons(2)%>"  style="width:90px" onclick="document.location.href='feedbacks.asp'" id=button2 name=button2></td>
				    <td width=10% nowrap></td>
					<TD align="center" width=45%><INPUT style="width:90px" type="submit" class="but_menu" value="<%=arrButtons(1)%>" id="button1" name="button1" onclick="return click_fun();"></TD>
				</TR>
			</TABLE>
		</td>
	</tr>
	<tr><td width="100%" height="10" colspan=4></td></tr>
	</form>    
</table>
<%End If%>
<% If (isNumeric(quest_id) = true) Then%>
<script>
<!--
  <%
			If trim(lang_id) = "1" Then
				str_alert = "!!!נא למלא את השדה"
			Else
				str_alert = "Please provide the answer!!!"
			End If	
   %> 		 		
		function click_fun()
		{ 
        <% sqlStr = "SELECT Field_Id,Field_Must,Field_Type FROM FORM_FIELD Where product_id=" & quest_id & " Order by Field_Order"
		'Response.Write sqlStr
		'Response.End
		set fields=con.GetRecordSet(sqlStr)
		 do while not fields.EOF 
		  Field_Id = fields("Field_Id")				
		  Field_Must = trim(fields("Field_Must"))
		  Field_Type = trim(fields("Field_Type"))
		  'Response.Write  Field_Type
		  If Field_Must = "True"  Then		  
		 %> 
		
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
	    <%
	     End If	 
	   End If
       fields.moveNext()
	   loop
       set fields=nothing
    %>   
       return true;	}
//-->
</SCRIPT> 
<%End If%>
<!-- end code -->                
<%set con=nothing%>
</body>
</html>

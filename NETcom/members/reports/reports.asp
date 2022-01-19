<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
	type_reports = trim(Request("type"))
	If type_reports = "cl" Then
		where_reports = " AND FORM_CLIENT = '1'"
	    numOftab = 1
        numOfLink = 3        
		topLevel2 = 17 'current bar ID in top submenu - added 03/10/2019
	Else
	    where_reports = " AND FORM_MAIL = '1'"
	    numOftab = 2
        numOfLink = 4
		topLevel2 = 63 'current bar ID in top submenu - added 03/10/2019
	End If
	start_date = "1/" & Month(Date()) & "/" & Year(date())
	end_date = DateAdd("d",30,start_date)
	companyID = trim(Request("company_ID"))
	User_ID = trim(Request("user_id"))
	quest_id = trim(Request("quest_id"))
	projectId = trim(Request("project_Id"))
	If Request.Form("dateStart") <> nil Then
		start_date = Request.Form("dateStart")
	End If
	If Request.Form("dateEnd") <> nil Then
		end_date = Request.Form("dateEnd")
	End If
%>
<%   
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 25 Order By word_id"				
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
<script language="javascript" type="text/javascript" src="../CalendarPopup.js"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function report_open(fname){
	if(isNaN(eval(window.document.all("quest_id").value)) == false)
	{
		formsearch.action=fname;
		formsearch.target = '_blank';
		formsearch.submit();
		
		start_tr.style.display='none';

	}else{
   <%
     If trim(lang_id) = "1" Then
        str_alert = "! נא לבחור טופס"
     Else
		str_alert = "Please choose a form !"
     End If   
    %>
	window.alert("<%=str_alert%>");
	window.document.all("quest_id").focus();
	}
}
function report_openedText(fname){
	if(isNaN(eval(window.document.all("quest_id").value)) == false)
	{
		formsearch.action=fname;
		formsearch.target = '_self';
		
		if (start_tr.style.display=='none')
		{	
			formsearch.submit();			
		}else{
			start_tr.style.display='none';
		}
				
	}else{
   <%
     If trim(lang_id) = "1" Then
        str_alert = "! נא לבחור טופס"
     Else
		str_alert = "Please choose a form !"
     End If   
    %>
	window.alert("<%=str_alert%>");	
	window.document.all("quest_id").focus();
	}
}
function report_openfr(fname){
	formsearch.action=fname;
	formsearch.target = '_blank';
	formsearch.submit();
}
function report_open_self(fname){
	if (window.document.all("quest_id").value != "0")
	{
		formsearch.action=fname;
		formsearch.target = '_self';
		formsearch.submit();
	}
}

var oPopup = window.createPopup();
function ProductDropDown(obj)
{
	oPopup.document.body.innerHTML = Product_Popup.innerHTML;
	var popupBodyObjProduct = oPopup.document.body;
	popupBodyObjProduct.style.border = "1px black solid"; 	    	 
	var dHeight = popupBodyObjProduct.scrollHeight;
	var dWidth = popupBodyObjProduct.scrollWidth;			
	oPopup.show(0, 21, 320, 108, obj); 
	popupBodyObjProduct = null;
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
		window.document.focus;   
}
function callCalendar(pf,pid)
	{

	cal1xx.select(pf,pid,'dd/MM/yyyy')
	}

function report_openedQuestion(fid,rType)
{	
	if(rType == 'cl')
		formsearch.action = "report_text_form.asp";
	else
		formsearch.action = "report_text.asp";
	formsearch.target = '_blank';
	formsearch.FieldId.value = fid;
	formsearch.submit();
	return false;
}
function report_openedQuestion(fid,rType)
{	
	if(rType == 'cl')
		formsearch.action = "report_text_form.asp";
	else
		formsearch.action = "report_text.asp";
	formsearch.target = '_blank';
	formsearch.FieldId.value = fid;
	formsearch.submit();
	return false;
}
//-->
</SCRIPT>
</head>
<body topmargin=0 bottommargin=0 leftmargin=0 rightmargin=0 bgcolor="#E5E5E5">
<!--#include file="../../logo_top.asp"-->
<table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
<tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="0" >
<tr><td width=100% align="<%=align_var%>"><!--#include file="../../top_in.asp"--></td></tr>
</table>
</td></tr>
<tr><td>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<tr><td class="page_title" dir="<%=dir_obj_var%>">&nbsp;</td></tr>

<tr><td height=15 nowrap></td></tr>
 <tr>    
    <td width="100%" valign="top" align="center">
    <table width="100%" cellspacing="0" cellpadding="3" align=center border="0" bgcolor="#ffffff" >
      <tr>
        <td width="100%" align=center valign=top >
<!-- start code -->
	<FORM action="" method=POST id=formsearch name=formsearch target=_self>
	<input type=hidden name="type" id="type" value="<%=type_reports%>">
	<table border="0" cellspacing="1" width=450 align=center bgcolor="#646E77" cellpadding="4" dir="<%=dir_var%>">
		<tr>
			<td bgcolor="#dbdbdb" align="center" colspan="2" height="15" nowrap>
				<input type="hidden" name="report" value="yes">
			</td>
		</tr>
		<tr>
			<td width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_obj_var%>">				
			<select dir="<%=dir_obj_var%>" name="quest_id" class="app" ID="quest_id" style="width:320">
			<option value="" id=word3 name=word3><%=arrTitles(3)%></option>
			<%
			 If type_reports = "cl" Then
				If is_groups = 0  Then
				sqlstr = "Select product_id, product_name from Products Where "&_
				" product_number = '0' and ORGANIZATION_ID=" & OrgID & " order by product_name"
				' משתמש אשר שייך לקבוצה אבל אינו אחראי באף קבוצה
				Else
				sqlstr = "Execute get_products_list '" & OrgID & "','" & UserID & "'"
				End If
				'Response.Write sqlstr
				'Response.End
				set rs_products = con.GetRecordSet(sqlstr)
				if not rs_products.eof then 
					ProductsList = rs_products.getRows()		
				end if
				set rs_products=nothing				
										
				Else
					sqlstr = "Select product_id, product_name from Products Where FORM_MAIL = '1' And "&_
					" product_number = '0' AND ORGANIZATION_ID=" & OrgID & " order by product_name"
					'Response.Write sqlstr
					'Response.End
					set rs_products = con.GetRecordSet(sqlstr)
					if not rs_products.eof then 
						ProductsList = rs_products.getRows()		
					end if
					set rs_products=nothing			
				End If
				
				If IsArray(ProductsList) Then
					For i=0 To Ubound(ProductsList,2)
						prod_Id = ProductsList(0,i)   	
						product_name = ProductsList(1,i)
				%>
					<OPTION value="<%=prod_Id%>" <%If trim(quest_id) = trim(prod_Id) Then%> selected <%End If%>> <%=product_name%> </OPTION>
				<%	Next	
				End If	
				%>
				</select>
			</td>
			<td width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id=word4 name=word4><!--טופס--><%=arrTitles(4)%></span>&nbsp;</td>
		</tr>
		<%If type_reports = "cl" Then%>
		<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">		
		<select name="user_id" id="user_id" class="norm" dir="<%=dir_obj_var%>" style="width:320" onchange="document.formsearch.action='reports.asp';document.formsearch.target='_self';document.formsearch.submit();">		
		<option value=""><%=arrTitles(5)%></option>	
		<%  sqlstr = "Select USER_ID, FIRSTNAME + ' ' + LASTNAME FROM USERS WHERE ORGANIZATION_ID = " & OrgID			
		    sqlstr = sqlstr & " Order BY FIRSTNAME + ' ' + LASTNAME "
			set rs_user=con.GetRecordSet(sqlstr)
			If Not rs_user.eof Then
				arr_user = rs_user.getRows()
			End If
			Set rs_user = Nothing
			If isArray(arr_user) Then
			For uu=0 To Ubound(arr_user,2)%>
			<option value="<%=arr_user(0,uu)%>" <%If trim(User_ID) = trim(arr_user(0,uu)) Then%> selected <%End if%>><%=arr_user(1,uu)%></option>			
		<%Next
		End If%>
		</select>
		</TD>
		<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<span id=word6 name=word6><!--עובד--><%=arrTitles(6)%></span>&nbsp;</td>
	</TR>
	<%Else%>
	<input type=hidden name="user_id" id="user_id" value="">
	<%End If%>
	<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">
		<select name="company_id" id="company_id" class="norm" dir="<%=dir_obj_var%>" style="width:320" onchange="document.formsearch.action='reports.asp';document.formsearch.target='_self';document.formsearch.submit();">
		<%If trim(lang_id) = "1" Then%>
		<option value=""><%=String(28,"-")%> כל ה<%=trim(Request.Cookies("bizpegasus")("CompaniesMulti"))%> <%=String(28,"-")%></option>	
		<%Else%>
		<option value=""><%=String(26,"-")%> All <%=trim(Request.Cookies("bizpegasus")("CompaniesMulti"))%> <%=String(26,"-")%></option>	
		<%End If%>
		<%sqlstr = "Select DISTINCT COMPANIES.Company_ID, COMPANIES.Company_NAME FROM COMPANIES INNER JOIN APPEALS "&_
		" ON COMPANIES.Company_ID = APPEALS.Company_ID " &_
		" WHERE COMPANIES.ORGANIZATION_ID = " & OrgID
		If trim(projectId) <> "" Then
		sqlstr = sqlstr & " AND APPEALS.project_Id = " & projectId
		End if	
		If trim(User_ID) <> "" Then
			sqlstr = sqlstr & " AND APPEALS.User_ID = " & User_ID
		End if	    
		sqlstr = sqlstr & " Order BY Company_NAME "
		set rs_comp=con.GetRecordSet(sqlstr)
		If Not rs_comp.eof Then
			arr_comp = rs_comp.getRows()
		End If
		Set rs_comp = Nothing
		If isArray(arr_comp) Then
		For cc=0 To Ubound(arr_comp,2)%>		
		<option value="<%=arr_comp(0,cc)%>" <%If trim(companyID) = trim(arr_comp(0,cc)) Then%> selected <%End if%>><%=arr_comp(1,cc)%></option>			
		<%  Next
		   End If%>		
		</select>
		</TD>
		<td width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;</td>
	</TR>
	<%If type_reports = "cl" Then%>
	<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">
		<select name="project_id" id="project_id" class="norm" dir="<%=dir_obj_var%>" style="width:320" onchange="document.formsearch.action='reports.asp';document.formsearch.target='_self';document.formsearch.submit();">
		<%If trim(lang_id) = "1" Then%>
		<option value=""><%=String(27,"-")%> כל ה<%=trim(Request.Cookies("bizpegasus")("ProjectMulti"))%> <%=String(27,"-")%></option>		
		<%Else%>
		<option value=""><%=String(28,"-")%> All <%=trim(Request.Cookies("bizpegasus")("ProjectMulti"))%> <%=String(28,"-")%></option>		
		<%End If%>
		<% 	sqlstr = "Select DISTINCT PROJECTS.PROJECT_ID, PROJECTS.PROJECT_NAME FROM APPEALS INNER JOIN PROJECTS "&_
			" ON APPEALS.PROJECT_ID = PROJECTS.PROJECT_ID WHERE APPEALS.ORGANIZATION_ID = " & OrgID
			If trim(companyID) <> "" Then
			sqlstr = sqlstr & " AND APPEALS.COMPANY_ID IN (0," & companyID  & ")"			
			End if	
			If trim(User_ID) <> "" Then
				sqlstr = sqlstr & " AND APPEALS.User_ID = " & User_ID
			End if			
			sqlstr = sqlstr & " Order BY PROJECT_NAME"
			Set rs_proj = con.getRecordSet(sqlstr)
			If Not rs_proj.eof Then
				arr_proj = rs_proj.getRows()
			End If
			Set rs_proj = Nothing
			If isArray(arr_proj) Then
			For pp=0 To Ubound(arr_proj,2)%>
			<option value="<%=arr_proj(0,pp)%>"  <%If trim(projectId) = trim(arr_proj(0,pp)) Then%> selected <%End if%>><%=arr_proj(1,pp)%></option>		
		<% Next
		  End If%>		
		</select>
		</TD>
			<TD width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<%=trim(Request.Cookies("bizpegasus")("Projectone"))	%>&nbsp;</TD>
		</TR>
		
	<TR>
		<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>">
		<select name="mechanism_id" id="mechanism_id" class="norm" dir="<%=dir_obj_var%>" style="width:320">
		<%If trim(lang_id) = "1" Then%>
		<option value=""><%=String(27,"-")%> כל המנגנונים <%=String(27,"-")%></option>		
		<%Else%>
		<option value=""><%=String(25,"-")%> All Sub-Projects <%=String(25,"-")%></option>		
		<%End If%>
		<% 	sqlstr = "Select DISTINCT Mechanism.Mechanism_ID, Mechanism.Mechanism_Name, PROJECTS.PROJECT_NAME "&_
			" FROM APPEALS INNER JOIN Mechanism ON APPEALS.Mechanism_ID = Mechanism.Mechanism_ID "&_
			" Inner JOIN PROJECTS ON PROJECTS.PROJECT_ID = Mechanism.PROJECT_ID WHERE APPEALS.ORGANIZATION_ID = " & OrgID
			If trim(ProjectID) <> "" Then
			sqlstr = sqlstr & " AND APPEALS.Project_ID =" & ProjectID
			End if	
			If trim(User_ID) <> "" Then
				sqlstr = sqlstr & " AND APPEALS.User_ID = " & User_ID
			End if			
			sqlstr = sqlstr & " Order BY Mechanism_Name"
			set rs_mech = con.getRecordSet(sqlstr)
			If Not rs_mech.eof Then
				arr_mech = rs_mech.getRows()
			End If
			Set rs_mech = Nothing
			If isArray(arr_mech) Then
			For mm=0 To Ubound(arr_mech,2)%>
			<option value="<%=arr_mech(0,mm)%>"  <%If trim(mechanismId) = trim(arr_mech(0,mm)) Then%> selected <%End if%>><%=arr_mech(2,mm)%> - <%=arr_mech(1,mm)%></option>		
		<% Next
		   End If%>		
		</select>
		</TD>
			<TD width="20%" class="subject_form" align="<%=align_var%>">&nbsp;<!--מנגנון--><%=arrTitles(14)%>&nbsp;</TD>
		</TR>		
		<%Else%>
		<input type=hidden name="project_id" id="project_id" value="">
		<%End If%>		
		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>">
	<a href='' onclick='callCalendar(document.formsearch.dateStart,"AsdateStart");return false;' id='AsdateStart'>
					<image src='../../images/calend.gif' border=0 ></a>&nbsp;<input dir='ltr' type='text' class='passw' size=10 id="dateStart" name='dateStart' value='<%=start_date%>' maxlength=10 readonly>&nbsp;
			
			</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id=word7 name=word7><!--- תאריך מ--><%=arrTitles(7)%></span>&nbsp;</TD>
		</TR>
		<TR>						
			<TD width="80%" bgcolor="#DBDBDB" align="<%=align_var%>" dir="<%=dir_var%>">	
				<a href='' onclick='callCalendar(document.formsearch.dateEnd,"AsdateEnd");return false;' id="AsdateEnd">
						<image src='../../images/calend.gif' border=0 ></a>&nbsp;<input dir="ltr" type="text" class="passw" size=10 id="dateEnd" name="dateEnd" value="<%=end_date%>" maxlength=10 readonly>&nbsp;
</TD>
			<TD width="20%" class="subject_form" nowrap align="<%=align_var%>">&nbsp;<span id="word8" name=word8><!--- תאריך עד--><%=arrTitles(8)%></span>&nbsp;</TD>
		</TR>	
		<%If type_reports = "cl" Then%>
		<%If trim(EMAILS) = "1" Then%>
		<!--<TR>					
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('report_stat.asp');"  style="width:250"><--התפלגות מכירות עפ"י מקורות פרסום-><%=arrTitles(15)%></a></TD>
		</TR>-->
		<%End If%>
		<TR>					
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('report_forms.asp');"  style="width:250"><!--דוח מסכם - טפסים מלאים--><%=arrTitles(9)%></a></TD>
		</TR>				
		<%Else%>
		<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('report_answers.asp');"  style="width:250"><!--דוח מסכם - משובים מדיוור--><%=arrTitles(10)%></a></TD>
		</TR>
		<%End If%>				
		<TR>						
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_openedText('reports.asp');" style="width:250"><!--דוח תשובות לשאלה חופשית--><%=arrTitles(11)%></a></TD>						
		</TR>					
		<TR id=start_tr style="display:<%if trim(quest_id) <> "" then%>block<%else%>none<%end if%>;">
			<TD style="height:100%" width=600 colspan="2" bgcolor="#DBDBDB">
				<input type="hidden" name="FieldId" id="FieldId" value="">		
				<TABLE width="100%" cellpadding="0" cellspacing="0" border=0 align=center>
				<tr><td style="width:600px">&nbsp;</td></tr>	
			<%
			if trim(quest_id) <> "" then
				sqlStr = "Select Langu from products where product_id=" & quest_id
				set prod = con.GetRecordSet(sqlStr)
				if not prod.eof then
					if prod("Langu") = "eng" then
						td_align = "left"
					else
						td_align = "right"
					end if
				end if
				set prod = nothing					
				
				sqlStr = "SELECT Field_Id,FIELD_TITLE FROM FORM_FIELD Where product_id = " & quest_id & " and (Field_Type =1 or Field_Type=2 or Field_Type=7) Order by Field_Order"
				set fields=con.GetRecordSet(sqlStr)
				If not fields.eof Then
					arr_fields = fields.getRows()
				End If
				Set fields = Nothing
				If isArray(arr_fields) Then
				For ff=0 To Ubound(arr_fields,2)
				Field_Id = trim(arr_fields(0,ff))
				FIELD_TITLE = trim(arr_fields(1,ff))					
				
				If FIELD_TITLE<>"" Then					
				%>
				<TR>
					<TD width="100%" bgcolor="#DBDBDB" align=<%=td_align%> valign="top"  bgcolor=#DADADA style="padding-right:3px;" dir="<%=dir_obj_var%>">&nbsp;&nbsp;<a class=linkFaq href='' onclick="return report_openedQuestion('<%=Field_Id%>','<%=type_reports%>')"><%=FIELD_TITLE%></a></td>
				</TR>
				<TR><TD height="5"></TD></TR>						
				<%
				else
				%>
				<TR>
					<TD width="100%" bgcolor="#DBDBDB" align=<%=td_align%> valign="top"  bgcolor=#DADADA nowrap>&nbsp;&nbsp;<a class=linkFaq href='' onclick='return report_openedQuestion(<%=Field_Id%>)'>הבושת</a>&nbsp;&nbsp;</td>
				</TR>	
				<TR><TD height="5"></TD></TR>																		
				<%
				end if
				Next		
			  end if	
			End If
			%>	
			</TABLE>
		  </TD>				
		</TR>									
		<TR>
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('report_graph.asp?type_reports=<%=type_reports%>');" style="width:250"><span id="word12" name=word12><!--(דוח התפלגות שאלות (עמודות--><%=arrTitles(12)%></span></a></TD>
		</TR>						
		<!--<TR>						
			<TD align=center bgcolor="#DBDBDB" colspan=2><A class="but_menu" href="#" onclick="report_open('report_graph_pie.asp?type_reports=<%=type_reports%>');" style="width:250"><span id="word13" name=word13>-(דוח התפלגות שאלות (עוגה-><%=arrTitles(13)%></span></a></TD>
		</TR>	-->
	</TABLE>			
</FORM>
  <!-- End main content -->
</td></tr>
</table>	 
  	<SCRIPT LANGUAGE="JavaScript">document.write(getCalendarStyles());</SCRIPT>
			<SCRIPT LANGUAGE="JavaScript" type="text/javascript">
            <!--
            var cal1xx = new CalendarPopup('CalendarDiv');
                cal1xx.showNavigationDropdowns();
                cal1xx.yearSelectStart
                cal1xx.offsetX = -50;
                cal1xx.offsetY = 0;
            //-->
			</SCRIPT>
					<DIV ID='CalendarDiv' name='CalendarDiv' STYLE='VISIBILITY:hidden;POSITION:absolute;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>
	
</body>
<%set con=nothing%>
</html>

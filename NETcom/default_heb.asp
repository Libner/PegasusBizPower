<!-- #include file="connect.asp" -->
<!-- #include file="reverse.asp" -->
<!-- #include file="members/checkWorker.asp" -->
<html>
<head>
<title>Bizpower</title>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="IE4.css" rel="STYLESHEET" company_type="text/css">
<link href="../dynamic_style.css" rel="STYLESHEET" type="text/css">
<script src="javascript/script.js"></script>
<script language=javascript>
<!--
		function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus; 		  
}

	var oPopup = window.createPopup();
	function StatusDropDown(obj)
	{
		oPopup.document.body.innerHTML = Status_Popup.innerHTML; 
		oPopup.document.charset = "windows-1255";
		oPopup.show(-65, 17, 80, 82, obj);    
	}
	
	var oPopupOut = window.createPopup();
	function StatusDropDownOut(obj)
	{
		oPopupOut.document.body.innerHTML = Status_Popup_OUT.innerHTML; 
		oPopupOut.document.charset = "windows-1255";
		oPopupOut.show(-65, 17, 80, 82, obj);    
	}
  
	function closeTask(contactID,companyID,taskID)
	{
			h = parseInt(470);
			w = parseInt(405);
			window.open("members/tasks/closetask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskId=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=50,left=100,width="+w+",height="+h+",align=center,resizable=0");
	}
	
	function addtask(contactID,companyID,taskID)
	{
		h = parseInt(470);
		w = parseInt(405);
		window.open("members/tasks/addtask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskID=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=50,left=120,width="+w+",height="+h+",align=center,resizable=0");		
	}
	
	function task_typeDropDown(obj)
	{
	    oPopup.document.body.innerHTML = task_type_Popup.innerHTML;
	    oPopup.document.charset = "windows-1255"; 
	    oPopup.show(-115, 17, 130, 82, obj);    
	}
	
	function task_typeDropDownOUT(obj)
	{
	    oPopupOut.document.body.innerHTML = task_type_Popup_OUT.innerHTML;
	    oPopupOut.document.charset = "windows-1255"; 
	    oPopupOut.show(-115, 17, 130, 82, obj);    
	}

//-->
</script>
</head>
<%
	'Session.Abandon()
	Response.Cookies("bizpegasus")("wizardId") = ""
	 
	If Request("wizard_id") <> nil Then	   
		Response.Cookies("bizpegasus")("wizardId") = Request("wizard_id")
		Response.Redirect "members/wizard/wizard_" & trim(Request("wizard_id")) & "_1.asp"				
	End If	
	
%>
<body marginwidth="10" marginheight="0" hspace="10" vspace="0" topmargin="0"
leftmargin="10" rightmargin="10" bgcolor="white">
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">        
<tr><td width=100% colspan=3>
<!--#include file="logo_top.asp"-->
</td></tr>

     <tr>				
			<!--td align=right valign="top" width="255" nowrap bgcolor="#DEDFF7">
			<include file="ashafim_inc.asp">				
			</td>							
			<td width="2" nowrap bgcolor="#FFFFFF"></td-->
			<td width=100% valign=top>
			<table cellpadding=0 cellspacing=0 width=100% border=0>
			<tr><td>
			<!--#include file="top.asp"-->
			</td></tr>			
			<tr><td height="2" bgcolor="#FFFFFF"></td></tr>			
			<tr>
				<td>
				<table border="0" width="100%" cellpadding="0" cellspacing="0">								
					<tr>
					    <%If trim(SURVEYS) = "1" Then%>
						<%If trim(SURVEYS) = "1" And trim(COMPANIES) = "1" Then%>
						<td align="left" width="50%">			
						<table border="0" width="98%" cellspacing="0" cellpadding="0">
						<%ElseIf trim(SURVEYS) = "1" And trim(COMPANIES) = "0" Then%>
						<td align="right" width="100%" colspan=2 valign=top>
						<table border="0" width="50%" cellspacing="0" cellpadding="0">						
						<%End If%>																
						<tr>
							<td bgcolor="#FFFFFF"></td>
							<td bgcolor="#FFFFFF" width="100%" align="center">
							<table border="0" width="100%" cellspacing="0" cellpadding="0">
								<tr>
									<td align="right">					
									<table border="0" width="100%" cellspacing="0" cellpadding="0">
										<tr>
											<td align="center" dir="rtl" class="title_form">
											הצג טפסים וסקרים
											</td>
										</tr>																		
										<tr>
							<td align="center">
								<table border="0" width=100% cellpadding="2" cellspacing="0" ID="Table8">																
							<%'If trim(COMPANIES) = "1" then%>
								<FORM action="members/appeals/appeal.asp" method=POST id="form_tofes" name=form_tofes target="_self">   
								<tr>
								<td align=right width=20% nowrap class=card></td>
								<td align=right width=50% nowrap class=card>
								<SELECT dir=rtl ID="quest_id" name=quest_id class="form_text" style="width: 95%" onchange="form_tofes.submit();">	
								<OPTION value="">בחר טופס</OPTION>
									<%
										set rs_products = con.GetRecordSet("Select product_id, product_name from products where product_number = '0' and ORGANIZATION_ID=" & trim(Request.Cookies("bizpegasus")("OrgID")) & " order by product_id")
										do while not rs_products.eof
										prod_Id = rs_products("product_id")
										product_name = rs_products("product_name")%>
										<OPTION value="<%=prod_Id%>" <%If trim(prodId) = trim(prod_Id) Then%> selected <%End If%>> <%=product_name%> </OPTION>
									<%	rs_products.MoveNext
										loop
									set rs_products=nothing%>	
								</SELECT>
								</td>
								<td align=right width=30% nowrap class=card style="padding-right:10px">מלא טופס</td>
								</tr>
								</FORM>
							
								<tr><td height=1 nowrap colspan=3></td></tr>
								<FORM action="members/appeals/appeals.asp" method=POST id="form_search_tofes1" name=form_search_tofes1 target="_self">   
								<tr>
								<td align=right width=20% nowrap class=card></td>
								<td align=right width=50% nowrap class=card>
								<SELECT dir=rtl ID="prodId" name=prodId class="form_text" style="width: 95%" onchange="form_search_tofes1.submit();" >	
								<OPTION value="">בחר טופס</OPTION>
									<%
										set rs_products = con.GetRecordSet("Select product_id, product_name from products where product_number = '0' AND FORM_CLIENT = '1' and ORGANIZATION_ID=" & trim(Request.Cookies("bizpegasus")("OrgID")) & " order by product_id")
										do while not rs_products.eof
										prod_Id = rs_products("product_id")
										product_name = rs_products("product_name")%>
										<OPTION value="<%=prod_Id%>" <%If trim(prodId) = trim(prod_Id) Then%> selected <%End If%>> <%=product_name%> </OPTION>
									<%	rs_products.MoveNext
										loop
									set rs_products=nothing%>	
								</SELECT>
								</td>
								<td align=right width=30% nowrap class=card style="padding-right:10px">טפסים מלאים</td>
								</tr>
								</FORM>
							
								<%'end if%>
								<tr><td height=1 nowrap colspan=3></td></tr>
								<FORM action="members/appeals/feedbacks.asp" method=POST id="form_search_tofes" name=form_search_tofes target="_self">   
								<tr>
								<td align=right width=20% nowrap class=card></td>
								<td align=right width=50% nowrap class=card>
								<SELECT dir=rtl ID=prodId name=prodId class="form_text" style="width: 95%" onchange="form_search_tofes.submit();">	
								<OPTION value="">בחר טופס</OPTION>
									<%
										set rs_products = con.GetRecordSet("Select product_id, product_name from products where product_number = '0' AND FORM_MAIL = '1' AND ORGANIZATION_ID=" & trim(Request.Cookies("bizpegasus")("OrgID")) & " order by product_id")
										do while not rs_products.eof
										prod_Id = rs_products("product_id")
										product_name = rs_products("product_name")%>
										<OPTION value="<%=prod_Id%>" <%If trim(prodId) = trim(prod_Id) Then%> selected <%End If%>> <%=product_name%> </OPTION>
									<%	rs_products.MoveNext
										loop
									set rs_products=nothing%>	
								</SELECT>
								</td>
								<td align=right width=30% nowrap class=card style="padding-right:10px">משובים מדיוור</td>
								</tr>
								</FORM>
								<tr><td height=5 nowrap colspan=3></td></tr>										
								</table>
							</td>
						</tr>
						
					</table>
					</td>
				</tr>
		</table>
		</td>
		<td bgcolor="#FFFFFF"></td>
		</tr>		
		</table>																				
	</td>
	<%End If%>	
	<%If trim(COMPANIES) = "1" Then%>
	<td width="20" nowrap></td>
	<%If trim(COMPANIES) = "1" And trim(SURVEYS) = "1" Then%>
	<td align="right" width="50%" valign=top>	
	<table border="0" width="98%" cellspacing="0" cellpadding="0">		
	<%ElseIf trim(COMPANIES) = "1" And trim(SURVEYS) = "0" Then%>
	<td align="right" width="100%" colspan=2 valign=top>
	<table border="0" width="50%" cellspacing="0" cellpadding="0">		
	<%End If%>
	  	
<tr>
	<td bgcolor="#FFFFFF"></td>
	<td bgcolor="#FFFFFF" width="100%" align="center">
	<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td align="right">					
		<table border="0" width="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="center" dir="rtl" class="title_form">
				הצג תיקי <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>
				</td>
			</tr>																								
			<tr>
				<td align="right">
					<table width="100%" border="0" cellpadding="2" cellspacing="0">							
					<FORM action="members/companies/companies.asp" method=POST id="Form1" name=form_search target="_self">   
					<tr>
					<td align=right width=25% nowrap class=card><input type="image" onclick="form_search.submit();" src="images/search.gif" ID="Image4" NAME="Image1"></td>
					<td align=right width=45% nowrap class=card><input type="text" class="form_text" dir="rtl" style="width:95%;" value="<%=vFix(search_company)%>" name="search_company" ID="Text1"></td>
					<td align=right width=30% nowrap class=card style="padding-right:10px">שם <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
					</tr>
					</FORM>
					<tr><td height=1 nowrap colspan=3></td></tr>
					<FORM action="members/projects/default.asp" method=POST id="Form2" name="form_search_project" target="_self">   
					<tr>
					<td align=right width=25% nowrap class=card><input type="image" onclick="form_search.submit();" src="images/search.gif" ID="Image5" NAME="Image2"></td>
					<td align=right width=45% nowrap class=card><input type="text" class="form_text" dir="rtl" style="width:95%;" value="<%=vFix(search_project)%>" name="search_project" ID="Text2"></td>
					<td align=right width=30% nowrap class=card style="padding-right:10px">שם <%=trim(Request.Cookies("bizpegasus")("Projectone"))%></td>
					</tr>
					</FORM>
					<tr><td height=1 nowrap colspan=3></td></tr>
					<FORM action="members/companies/contacts.asp" method=POST id="Form3" name="form_search_contact" target="_self">   
					<tr>
					<td align=right width=25% nowrap class=card><input type="image" onclick="form_search.submit();" src="images/search.gif" ID="Image6" NAME="Image2"></td>
					<td align=right width=45% nowrap class=card><input type="text" class="form_text" dir="rtl" style="width:95%;" value="<%=vFix(search_contact)%>" name="search_contact" ID="Text3"></td>
					<td align=right width=30% nowrap class=card style="padding-right:10px">שם <%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td>
					</tr>
					</FORM>												
				</table>
		</td></tr>
	</table></td></tr>
	</table></td>
	<td bgcolor="#FFFFFF"></td></tr>
	</table></td>
	<%End If%>	
	</tr>
	<tr><td height="20"></td></tr>									
	<tr>
	    <%If trim(CASH_FLOW) = "1" Then%>
	    <%If trim(CASH_FLOW) = "1" And trim(WORK_PRICING) = "1" Then%>
		<td align="left" width="50%" valign=top>	
		<table border="0" width="98%" cellspacing="0" cellpadding="0">	
		<%ElseIf trim(CASH_FLOW) = "1" And trim(WORK_PRICING) = "0" Then%>
		<td align="left" width="100%" colspan=2 valign=top>
		<table border="0" width="50%" cellspacing="0" cellpadding="0">	
		<%End If%>					
			<tr>
				<td bgcolor="#FFFFFF"></td>
				<td bgcolor="#FFFFFF" width="100%" align="center">
				<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<tr>
						<td align="left">
						
						<table border="0" width="100%" cellspacing="0" cellpadding="0">
							<tr>
								<td align="center" dir="rtl" class="title_form">
								הצג תזרים
								</td>
							</tr>																		
							<tr>
								<td align="right">
									<table width="100%" border="0" cellpadding="2" cellspacing="0">
									
								<%
								start_date = DateValue("1/" & Month(Date()) & "/" & Year(date()))
								end_date = DateAdd("d",30,start_date)
								%>	
								<FORM action="members/projects/movements_activate.asp" method=POST id="form_search_tazrim" name=form_search_tazrim target="_self">   
								<tr>								
								<td align=right width=20% nowrap class=card></td>
								<td align=right width=50% nowrap class=card><input dir="ltr" class="form_text" type="text" id="start_date" name="start_date" value="<%=start_date%>" style="width:75" onclick="return popupcal(this);" readonly></td>
								<td align=right width=30% nowrap class=card style="padding-right:10px">מתאריך</td>
								</tr>
								<tr><td height=1 nowrap colspan=3></td></tr>																							
								<tr>
								<td align=right width=20% nowrap class=card></td>
								<td align=right width=50% nowrap class=card><input dir="ltr" class="form_text" type="text" id="end_date" name="end_date" value="<%=end_date%>" style="width:75" onclick="return popupcal(this);" readonly></td>
								<td align=right width=30% nowrap class=card style="padding-right:10px">עד תאריך</td>
								</tr>
								<tr><td height=1 nowrap colspan=3></td></tr>
								<tr>
 							    <td align=right width=20% nowrap class=card></td>
								<td align=right width=50% nowrap class=card>
								<select dir="rtl" name="bank_id" style="width:95%" class="form_text" ID="bank_id" onchange="form_search_tazrim.submit();">
								<OPTION value="">בחר חשבון</OPTION>		
									<%  sqlstr = "Select bank_id, bank_name FROM banks WHERE ORGANIZATION_ID = " & OrgID & " Order BY bank_name"
										set rs_bank = con.getRecordSet(sqlstr)
										While not rs_bank.eof
									%>
									<option value="<%=rs_bank(0)%>" <%If trim(bank_id) = trim(rs_bank(0)) Then%> selected <%End If%>><%=rs_bank(1)%></option>
									<%
										rs_bank.moveNext
										Wend
										set rs_bank = Nothing
									%>
									</select>
								</td>
								<td align=right width=30% nowrap class=card style="padding-right:10px">חשבון בנק</td>
								</tr>	
								</FORM>															
								</table>					
								</td>
							</tr>
						</table>
						</td>
					</tr>
			</table>
				</td>
				<td bgcolor="#FFFFFF"></td>
			</tr>
	    </table>	 																					
	</td>
	<%End If%>			
	<%If trim(WORK_PRICING) = "1" Then%>
	<td width="20" nowrap></td>
	   <%If trim(WORK_PRICING) = "1" And trim(CASH_FLOW) = "1" Then%>
		<td align="right" width="50%" valign=top>		
		<table border="0" width="98%" cellspacing="0" cellpadding="0">
		<%ElseIf trim(WORK_PRICING) = "1" And trim(CASH_FLOW) = "0" Then%>
		<td align="right" width="100%" colspan=2 valign=top>
		<table border="0" width="50%" cellspacing="0" cellpadding="0">
		<%End If%>									
		<tr>
			<td bgcolor="#FFFFFF"></td>
			<td bgcolor="#FFFFFF" width="100%" align="center">
			<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table2">
			<tr>
				<td align="left">												
				<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table3">
				<tr>
					<td align="center" dir="rtl" class="title_form">
					עלות העבודה
					</td>
				</tr>
																																	
				<tr>
			<td align="right">
			<table width="100%" border="0" cellpadding="2" cellspacing="0" ID="Table4">								
			<FORM action="members/projects/projects_list.asp" method=POST id="Form4" name=form_search_alut target="_self">   
			<tr>
			<td align=right width=25% nowrap class=card><input type="image" onclick="form_search_alut.submit();" src="images/search.gif" ID="Image12" NAME="Image12"></td>
			<td align=right width=45% nowrap class=card><input type="text" class="form_text" dir="rtl" style="width:95%;" value="<%=vFix(search_company)%>" name="search_company" ID="Text4"></td>
			<td align=right width=30% nowrap class=card style="padding-right:10px">ל<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
			</tr>
			</FORM>
			<tr><td height=1 nowrap colspan=3></td></tr>
			<FORM action="members/projects/projects_list.asp" method=POST id="Form5" name="form_search_project_alut" target="_self">   
			<tr>
			<td align=right width=25% nowrap class=card><input type="image" onclick="form_search_project_alut.submit();" src="images/search.gif" ID="Image13" NAME="Image13"></td>
			<td align=right width=45% nowrap class=card><input type="text" class="form_text" dir="rtl" style="width:95%;" value="<%=vFix(search_project)%>" name="search_project" ID="Text5"></td>
			<td align=right width=30% nowrap class=card style="padding-right:10px">ל<%=trim(Request.Cookies("bizpegasus")("Projectone"))%></td>
			</tr>
			</FORM>
			<tr><td height=1 nowrap colspan=3></td></tr>
			<FORM action="members/projects/workers.asp" method=POST id="Form6" name="form_search_worker_alut" target="_self">   
			<tr>
			<td align=right width=25% nowrap class=card><input type="image" onclick="form_search_worker_alut.submit();" src="images/search.gif" ID="Image14" NAME="Image14"></td>
			<td align=right width=45% nowrap class=card><input type="text" class="form_text" dir="rtl" style="width:95%;" value="<%=vFix(search_worker)%>" name="search_worker" ID="Text6"></td>
			<td align=right width=30% nowrap class=card style="padding-right:10px">לעובד</td>
			</tr>
			</FORM>												
		</table></td>	
		</tr></table>
		</td></tr></table>			
		</td>
		<td bgcolor="#FFFFFF"></td>
		</tr>
		</table></td>
		<%End If%>
	</tr>															
	</table></td>
	</tr>
 <%	If trim(TASKS) = "1" Then %>
	<tr><td height=2 nowrap bgcolor="#FFFFFF"></td></tr>
 <%
    reciver_id = UserID 
    If trim(UserID) = trim(sender_id)  Then
		class_ = "4"
	ElseIf  trim(UserID) = trim(reciver_id) Then
		class_ = "7"
	End if	 
   
	if trim(reciver_id) <> "" then
		where_reciver = " AND reciver_id = " & reciver_id
		where_sender = ""
		title = "נכנסות" 	
	end if
	if trim(Request.QueryString("task_status"))<>"" then
		task_status=Request.QueryString("task_status")    
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

	task_sort = Request.QueryString("task_sort")	
	if trim(task_sort)="" then  task_sort=0 end if 

	If trim(task_status) <>""  Then		
		if trim(task_status) = "all" Then
			status = ""
			where_status = ""
		else	
			where_status = " And (task_status in (" & sFix(task_status) & "))"
			status = task_status
		end If	
	ElseIf sender_id = "" Then
		where_status = " And (task_status IN ('1','2')) " 
 		status = "1,2"	 	   
	Else
		status = "1,2"
		where_status = " And (task_status IN ('1','2')) "  
	End If
 
	If trim(Request("taskTypeID")) <> nil Then
		taskTypeID = trim(Request("taskTypeID"))
	Else taskTypeID = ""	
	End If	  
		
	Select Case(status)
		Case "1" : no_search = "פתוחות"
		Case "2" : no_search = "בטיפול" 
		Case "1,2" : no_search = "פתוחות או בטיפול"
		Case "3" : no_search = "סגורות" 
		Case "all" : no_search = ""
	End Select

	dim arr_Status(3)
	arr_Status(1)="פתוחה"
	arr_Status(2)="בטיפול"
	arr_Status(3)="סגורה"

	dim sortby(12)	
	sortby(0) = "task_status, task_date DESC"
	sortby(1) = "rtrim(ltrim(company_name))"
	sortby(2) = "rtrim(ltrim(company_name)) DESC"
	sortby(3) = "task_date"
	sortby(4) = "task_date DESC"
	sortby(5) = "contact_name"
	sortby(6) = "contact_name DESC"
	sortby(7) = "sender_name"
	sortby(8) = "sender_name DESC"
	sortby(9) = "reciver_name"
	sortby(10) = "reciver_name DESC"
	sortby(11) = "project_name"
	sortby(12) = "project_name DESC"

	urlSort="default.asp?task_status=" & task_status & "&task_status_OUT=" & task_status_OUT & "&task_sort_OUT=" & task_sort_OUT
	UrlStatus="default.asp?task_sort = " & task_sort & "&task_status_OUT=" & task_status_OUT & "&task_sort_OUT=" & task_sort_OUT
	urlType="default.asp?task_status=" & task_status & "&task_sort=" &task_sort & "&task_status_OUT=" & task_status_OUT & "&task_sort_OUT=" & task_sort_OUT
%>
<tr>
<td class="title_form" width=100% align=right dir=rtl><a class=normalB href="members/tasks/default.asp?T=IN" target=_self><font color="#ffffff">&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%> נכנסות&nbsp;<%=no_search%>&nbsp;</font><font color="#E6E6E6">(<%=user_name%>)</font>&nbsp;</a></td>
</tr>		   
<tr><td width=100%> 
<a name="table_tasks"></a>
   <table width="100%" cellspacing="1" cellpadding="0" border=0>      
     <tr style="line-height:18px"> 	  
          <td align=center class="title_sort" width=26 nowrap>&nbsp;</td>     
		  <td align=right class="title_sort" width=100%>תוכן&nbsp;</td>              
          <td width="160" nowrap align="right" class="title_sort<%if trim(task_sort)="1" OR trim(task_sort)="2" then%>_act<%end if%>"><%if trim(task_sort)="1" then%><a class="title_sort" href="<%=urlSort%>&task_sort=<%=task_sort+1%>#task_status_OUT" title="למיון בסדר יורד"><%elseif trim(task_sort)="2" then%><a class="title_sort" href="<%=urlSort%>&task_sort=<%=task_sort-1%>#task_status_OUT" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&task_sort=1#task_status_OUT" title="למיון בסדר עולה"><%end if%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<img src="images/arrow_<%if trim(task_sort)="1" then%>bot<%elseif trim(task_sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	        
          <td width="85" align="right" nowrap class="title_sort<%if trim(task_sort)="9" OR trim(task_sort)="10" then%>_act<%end if%>"><%if trim(task_sort)="9" then%><a class="title_sort" href="<%=urlSort%>&task_sort=<%=task_sort+1%>#task_status_OUT" title="למיון בסדר יורד"><%elseif trim(task_sort)="10" then%><a class="title_sort" href="<%=urlSort%>&task_sort=<%=task_sort-1%>#task_status_OUT" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&task_sort=9#task_status_OUT" title="למיון בסדר עולה"><%end if%>אל&nbsp;<img src="images/arrow_<%if trim(task_sort)="9" then%>bot<%elseif trim(task_sort)="10" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
          <td width="85" align="right" nowrap class="title_sort<%if trim(task_sort)="7" OR trim(task_sort)="8" then%>_act<%end if%>"><%if trim(task_sort)="7" then%><a class="title_sort" href="<%=urlSort%>&task_sort=<%=task_sort+1%>#task_status_OUT" title="למיון בסדר יורד"><%elseif trim(task_sort)="8" then%><a class="title_sort" href="<%=urlSort%>&task_sort=<%=task_sort-1%>#task_status_OUT" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&task_sort=7#task_status_OUT" title="למיון בסדר עולה"><%end if%>מאת&nbsp;<img src="images/arrow_<%if trim(task_sort)="7" then%>bot<%elseif trim(task_sort)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	      <td width="55" align="right" nowrap class="title_sort<%if trim(task_sort)="3" OR trim(task_sort)="4" then%>_act<%end if%>"><%if trim(task_sort)="3" then%><a class="title_sort" href="<%=urlSort%>&task_sort=<%=task_sort+1%>#task_status_OUT" title="למיון בסדר יורד"><%elseif trim(task_sort)="4" then%><a class="title_sort" href="<%=urlSort%>&task_sort=<%=task_sort-1%>#task_status_OUT" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&task_sort=3#task_status_OUT" title="למיון בסדר עולה"><%end if%>תאריך יעד<img src="images/arrow_<%if trim(task_sort)="3" then%>bot<%elseif trim(task_sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	      <td width="45" align="right" nowrap class="title_sort">'סט&nbsp;<IMG style="cursor:hand;" src="images/icon_find.gif" BORDER=0 ALT="בחר מרשימה" align=absmiddle onmousedown="StatusDropDown(this)"></td>
	    
    </tr>   
<%     
   PageSize = 5      
   Set tasksList = Server.CreateObject("ADODB.RECORDSET")     
   sqlstr = "EXECUTE get_tasks '" & search_company & "','" & search_contact & "','" & search_project & "','" & status & "','" & OrgID & "','" & taskTypeID & "','" & reciver_id & "','" & sender_id & "','" & sortby(task_sort) & "','',''"
   'Response.Write sqlstr
   'Response.End
   set  tasksList = con.getRecordSet(sqlstr)
   If not tasksList.EOF then		
		tasksList.PageSize=PageSize
		tasksList.AbsolutePage=Page
		recCount=tasksList.RecordCount 
		NumberOfPages=tasksList.PageCount
		i=0	
		arrtasks = tasksList.getRows(tasksList.PageSize)
		tasksList.Close()
		set tasksList = Nothing	
       
  do while i <= uBound(arrtasks,2)            
       companyName =  trim(arrtasks(0,i))
       'Response.Write companyName
       'Response.End
       contactName =  trim(arrtasks(1,i))
       projectName = trim(arrtasks(2,i))  
       reciver_name = trim(arrtasks(3,i))
	   sender_name = trim(arrtasks(4,i))
       task_types = trim(arrtasks(6,i))
       companyID =  trim(arrtasks(11,i))
       contactID = trim(arrtasks(12,i))
       taskId = trim(arrtasks(13,i))
       taskStatus = trim(arrtasks(17,i))
       task_content = trim(arrtasks(8,i))
       private_flag = trim(arrtasks(20,i))
       parentID =trim(arrtasks(15,i))
       
       If trim(private_flag) Then
		   url_ = "_private"
	   Else
		   url_ = ""	   
       End If   
       
       If IsDate(trim(arrtasks(7,i))) Then
		  task_date=Day(trim(arrtasks(7,i))) & "/" & Month(trim(arrtasks(7,i))) & "/" & Right(Year(trim(arrtasks(7,i))),2)
	   Else
		   task_date=""
	   End If     				
	  
	   task_types_n = ""          
	   short_task_types = ""
	   If Len(task_types) > 0 Then				
			sqlstr="Select activity_type_name From activity_types Where activity_type_id IN (" & task_types & ") Order By activity_type_id"
			'Response.Write sqlstr
			'Response.End
			set rssub = con.getRecordSet(sqlstr)				
			While not rssub.eof
				task_types_n = task_types_n & rssub(0) & "<BR>"
				rssub.moveNext
			Wend		    
			set rssub=Nothing
			If Len(task_types_n) > 0 Then
				task_types_n = Left(task_types_n,(Len(task_types_n)-4))
			End If				
		End If
       
       If trim(reciver_id) <> "" Then
       href = "href=""javascript:closeTask('" & contactID & "','" & companyID & "','" & taskID & "')"""     
       ElseIf trim(taskStatus) = "1" Then
       href = "href=""javascript:addtask('" & contactID & "','" & companyID & "','" & taskID & "')"""   
       End If
       If trim(companyID) <> "" And trim(COMPANIES) = "1" Then
       href_company = "href=""members/companies/company"&url_&".asp?companyID=" & companyID & """"
       Else
       href_company = ""
       End If
       
       If trim(taskID) <> "" Then
			sqlstr = "Select TOP 1 task_id from tasks WHERE parent_id = " & taskID
			set rs_par = con.getRecordSet(sqlstr)
			if not rs_par.eof then
				is_parent = true
			else
				is_parent = false	
			end if
			set rs_par = nothing
		End If 
     %>
        <tr height=18>
          <td align=center class="card<%=class_%>" valign=middle>
		  <%If trim(taskID) <> "" And is_parent Then%>
		  <input type=image src="images/hets4.gif" border=0 hspace=0 vspace=0 title="<%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%> המשך" onclick='window.open("members/tasks/task_children.asp?parentID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' ID="Image3" NAME="Image1">
		  <%End If%>
		  <%If trim(parentID) <> "" And trim(taskID) <> "" Then%>
		  <input type=image src="images/hets4a.gif" border=0 hspace=0 vspace=0 title="היסטורית משימה" onclick='window.open("members/tasks/task_parents.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' ID="Image9" NAME="Image4">
		  <%End If%>
		</td>        
	      <td class="card<%=class_%>" dir=rtl valign=top align="right" dir=rtl><a class="link_categ" <%=href%>><%=trim(task_content)%></a></td>	      	      	      
          <td class="card<%=class_%>" valign=top align="right" dir=rtl><a class="link_categ" <%=href%>><%=companyName%></a></td>
          <td class="card<%=class_%>" valign=top align="right" dir=rtl><a class="link_categ" <%=href%>>ל<%=reciver_name%></a></td>
          <td class="card<%=class_%>" valign=top align="right" dir=rtl><a class="link_categ" <%=href%>>מ<%=sender_name%></a></td>
          <td class="card<%=class_%>" valign=top align="right" dir=rtl><a class="link_categ" <%=href%>><%If isDate(task_date) Then%><%=task_date%><%End If%></a></td>         
	      <td class="card<%=class_%>" valign=top align="center" dir=rtl><a class="task_status_num<%=taskStatus%>" <%=href%>><%=arr_Status(taskStatus)%></a></td>	  
      </tr> 
	  <%
	  i=i+1	
	  loop
	  if NumberOfPages > 1 then
	  urlSort = urlSort & "&task_sort=" & task_sort
	  %>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap class="card<%=class_%>">
			<table border="0" cellspacing="0" cellpadding="2">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center" align="right"><A class=pageCounter title="לדפים הקודמים" href="<%=urlSort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>#table_tasks_OUT" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle" align="right"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  align="right"><A class=pageCounter href="<%=urlSort%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>#table_tasks_OUT" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td align="right"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center" align="right"><A class=pageCounter title="לדפים הבאים" href="<%=urlSort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>#table_tasks_OUT" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>	
	<%End If%> 
	<tr>
	   <td colspan="10" height=18 class="card<%=class_%>" align=center dir="ltr" style="color:#6E6DA6;font-weight:600">נמצאו <%=recCount%>&nbsp; <%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%> &nbsp;<%=no_search%></td>
	</tr>
	<%Else%>
	<tr><td colspan=10 class="card<%=class_%>" align=center>&nbsp;לא נמצאו <%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%>&nbsp;<%=no_search%></td></tr>
<% End If%>
</table>
<%'//end of left text%>
	</td></tr>	<tr><td height=2 nowrap bgcolor="#FFFFFF"></td></tr>
 <%
 sender_id = UserID 
 reciver_id = ""
 If trim(UserID) = trim(sender_id)  Then
	class_ = "4"
ElseIf  trim(UserID) = trim(reciver_id) Then
	class_ = "7"
End if	
	   
if trim(sender_id) <> "" then
	where_sender = " AND user_id = " & sender_id
	where_reciver = ""	
	title = "יוצאות"		
 end if
 if trim(Request.QueryString("task_status_OUT"))<>"" then
    task_status_OUT=Request.QueryString("task_status_OUT")    
 end if  
 
 if trim(Request.QueryString("pageOUT"))<>"" then
    pageOUT=Request.QueryString("pageOUT")
 else
    pageOUT=1
 end if  
 
 if trim(Request.QueryString("numOfRowOut"))<>"" then
    numOfRowOut=Request.QueryString("numOfRowOut")
 else    
    numOfRowOut = 1
 end if  

 task_sort_OUT = Request.QueryString("task_sort_OUT")	
 if trim(task_sort_OUT)="" then  task_sort_OUT=0 end if 

 If trim(task_status_OUT) <>""  Then		
	if trim(task_status_OUT) = "all" Or trim(task_status_OUT) = "" Then
		status_OUT = ""
		where_status = ""
	else	
		where_status = " And (task_status in (" & sFix(task_status_OUT) & "))"
		status_OUT = task_status_OUT
	end If	 
 End If
 
If trim(Request("tasktypeIDOUT")) <> nil Then
	 tasktypeIDOUT = trim(Request("tasktypeIDOUT"))
Else tasktypeIDOUT = ""	
End If		
		
Select Case(trim(status_Out))
 Case "1" : no_search_OUT = "פתוחות"
 Case "2" : no_search_OUT = "בטיפול" 
 Case "3" : no_search_OUT = "סגורות" 
 Case "" : no_search_OUT = ""
End Select

arr_Status(1)="פתוחה"
arr_Status(2)="בטיפול"
arr_Status(3)="סגורה"

'dim sortby(12)	
sortby(0) = "task_date DESC"
sortby(1) = "rtrim(ltrim(company_name))"
sortby(2) = "rtrim(ltrim(company_name)) DESC"
sortby(3) = "task_date"
sortby(4) = "task_date DESC"
sortby(5) = "contact_name"
sortby(6) = "contact_name DESC"
sortby(7) = "sender_name"
sortby(8) = "sender_name DESC"
sortby(9) = "reciver_name"
sortby(10) = "reciver_name DESC"
sortby(11) = "project_name"
sortby(12) = "project_name DESC"

urlSort_Out="default.asp?task_status_OUT=" & task_status_OUT & "&task_sort=" & task_sort & "&task_status=" & task_status
UrlStatus_Out="default.asp?task_sort_OUT= " & task_sort_OUT & "&task_sort=" & task_sort & "&task_status=" & task_status
urlType_Out="default.asp?task_status_OUT=" & task_status_OUT & "&task_sort_OUT=" &task_sort_OUT & "&task_sort=" & task_sort & "&task_status=" & task_status
%>
<tr>
<td class="title_form" width=100% align=right dir=rtl><a class=normalB href="members/tasks/default.asp?T=OUT" target=_self><font color="#ffffff">&nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%> יוצאות&nbsp;<%=no_search_Out%>&nbsp;</font><font color="#E6E6E6">(<%=user_name%>)</font>&nbsp;</a></td>
</tr>	
<tr><td width=100%> 
<a name="table_tasks_OUT"></a>
   <table width="100%" cellspacing="1" cellpadding="0" border=0>      
     <tr style="line-height:18px"> 	  
        <td align=center class="title_sort" width=26 nowrap>&nbsp;</td>     
		<td align=right class="title_sort" width=100%>תוכן&nbsp;</td>
		<td width="160" nowrap align="right" class="title_sort<%if trim(task_sort_OUT)="1" OR trim(task_sort_OUT)="2" then%>_act<%end if%>"><%if trim(task_sort_OUT)="1" then%><a class="title_sort" href="<%=urlSort_Out%>&task_sort_OUT=<%=task_sort_OUT+1%>#table_tasks" title="למיון בסדר יורד"><%elseif trim(task_sort_OUT)="2" then%><a class="title_sort" href="<%=urlSort_Out%>&task_sort_OUT=<%=task_sort_OUT-1%>#table_tasks" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort_Out%>&task_sort_OUT=1#table_tasks" title="למיון בסדר עולה"><%end if%><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;<img src="images/arrow_<%if trim(task_sort_OUT)="1" then%>bot<%elseif trim(task_sort_OUT)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	        
        <td width="85" align="right" nowrap class="title_sort<%if trim(task_sort_OUT)="9" OR trim(task_sort_OUT)="10" then%>_act<%end if%>"><%if trim(task_sort_OUT)="9" then%><a class="title_sort" href="<%=urlSort_Out%>&task_sort_OUT=<%=task_sort_OUT+1%>#table_tasks" title="למיון בסדר יורד"><%elseif trim(task_sort_OUT)="10" then%><a class="title_sort" href="<%=urlSort_Out%>&task_sort_OUT=<%=task_sort_OUT-1%>#table_tasks" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort_Out%>&task_sort_OUT=9#table_tasks" title="למיון בסדר עולה"><%end if%>אל&nbsp;<img src="images/arrow_<%if trim(task_sort_OUT)="9" then%>bot<%elseif trim(task_sort_OUT)="10" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
        <td width="85" align="right" nowrap class="title_sort<%if trim(task_sort_OUT)="7" OR trim(task_sort_OUT)="8" then%>_act<%end if%>"><%if trim(task_sort_OUT)="7" then%><a class="title_sort" href="<%=urlSort_Out%>&task_sort_OUT=<%=task_sort_OUT+1%>#table_tasks" title="למיון בסדר יורד"><%elseif trim(task_sort_OUT)="8" then%><a class="title_sort" href="<%=urlSort_Out%>&task_sort_OUT=<%=task_sort_OUT-1%>#table_tasks" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort_Out%>&task_sort_OUT=7#table_tasks" title="למיון בסדר עולה"><%end if%>מאת&nbsp;<img src="images/arrow_<%if trim(task_sort_OUT)="7" then%>bot<%elseif trim(task_sort_OUT)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>          
	    <td width="55" align="right" nowrap class="title_sort<%if trim(task_sort_OUT)="3" OR trim(task_sort_OUT)="4" then%>_act<%end if%>"><%if trim(task_sort_OUT)="3" then%><a class="title_sort" href="<%=urlSort_Out%>&task_sort_OUT=<%=task_sort_OUT+1%>#table_tasks" title="למיון בסדר יורד"><%elseif trim(task_sort_OUT)="4" then%><a class="title_sort" href="<%=urlSort_Out%>&task_sort_OUT=<%=task_sort_OUT-1%>#table_tasks" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort_Out%>&task_sort_OUT=3#table_tasks" title="למיון בסדר עולה"><%end if%>תאריך יעד<img src="images/arrow_<%if trim(task_sort_OUT)="3" then%>bot<%elseif trim(task_sort_OUT)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	    <td width="45" align="right" nowrap class="title_sort">'סט&nbsp;<IMG style="cursor:hand;" src="images/icon_find.gif" BORDER=0 ALT="בחר מרשימה" align=absmiddle onmousedown="StatusDropDownOut(this)"></td>	    
    </tr>   
<%     
   PageSize = 5      
   Set tasksList = Server.CreateObject("ADODB.RECORDSET")     
   sqlstr = "EXECUTE get_tasks '" & search_company & "','" & search_contact & "','" & search_project & "','" & status_Out & "','" & OrgID & "','" & tasktypeIDOUT & "','" & reciver_id & "','" & sender_id & "','" & sortby(task_sort_OUT) & "','',''"
   'Response.Write sqlstr
   'Response.End
   set  tasksList = con.getRecordSet(sqlstr)
   If not tasksList.EOF then		
		tasksList.PageSize=PageSize
		tasksList.AbsolutePage=pageOUT
		recCount=tasksList.RecordCount 
		NumberOfPages=tasksList.PageCount
		i=0	
		arrtasks = tasksList.getRows(tasksList.PageSize)
		tasksList.Close()
		set tasksList = Nothing	
       
  do while i <= uBound(arrtasks,2)            
       companyName =  trim(arrtasks(0,i))
       'Response.Write companyName
       'Response.End
       contactName =  trim(arrtasks(1,i))
       projectName = trim(arrtasks(2,i))  
       reciver_name = trim(arrtasks(3,i))
	   sender_name = trim(arrtasks(4,i))
       task_types = trim(arrtasks(6,i))
       companyID =  trim(arrtasks(11,i))
       contactID = trim(arrtasks(12,i))
       taskId = trim(arrtasks(13,i))
       taskStatus = trim(arrtasks(17,i))
       task_content = trim(arrtasks(8,i)) 
       private_flag = trim(arrtasks(20,i))
       parentID = trim(arrtasks(15,i))
       
       If trim(private_flag) Then
		   url_ = "_private"
	   Else
		   url_ = ""	   
       End If     
       
       If IsDate(trim(arrtasks(7,i))) Then
		  task_date=Day(trim(arrtasks(7,i))) & "/" & Month(trim(arrtasks(7,i))) & "/" & Right(Year(trim(arrtasks(7,i))),2)
	   Else
		   task_date=""
	   End If     				
	  
	   task_types_n = ""          
	   short_task_types = ""
	   If Len(task_types) > 0 Then				
			sqlstr="Select activity_type_name From activity_types Where activity_type_id IN (" & task_types & ") Order By activity_type_id"
			'Response.Write sqlstr
			'Response.End
			set rssub = con.getRecordSet(sqlstr)				
			While not rssub.eof
				task_types_n = task_types_n & rssub(0) & "<BR>"
				rssub.moveNext
			Wend		    
			set rssub=Nothing
			If Len(task_types_n) > 0 Then
				task_types_n = Left(task_types_n,(Len(task_types_n)-4))
			End If				
		End If
       
       If trim(taskStatus) = "1" Then
       href = "href=""javascript:addtask('" & contactID & "','" & companyID & "','" & taskID & "')"""   
       ElseIf trim(sender_id) <> "" Then
       href = "href=""javascript:closeTask('" & contactID & "','" & companyID & "','" & taskID & "')"""     
       End If
       If trim(companyID) <> "" And trim(COMPANIES) = "1" Then
       href_company = "href=""members/companies/company"&url_&".asp?companyID=" & companyID & """"
       Else
       href_company = ""
       End If      
       
       If trim(taskID) <> "" Then
			sqlstr = "Select TOP 1 task_id from tasks WHERE parent_id = " & taskID
			set rs_par = con.getRecordSet(sqlstr)
			if not rs_par.eof then
				is_parent = true
			else
				is_parent = false	
			end if
			set rs_par = nothing
		End If 

      %>
        <tr height=18>
            <td align=center class="card<%=class_%>" valign=middle>
			<%If trim(taskID) <> "" And is_parent Then%>
			<input type=image src="images/hets4.gif" border=0 hspace=0 vspace=0 title="<%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%> המשך" onclick='window.open("members/tasks/task_children.asp?parentID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' ID="Image10" NAME="Image1">
			<%End If%>
			<%If trim(parentID) <> "" And trim(taskID) <> "" Then%>
			<input type=image src="images/hets4a.gif" border=0 hspace=0 vspace=0 title="היסטורית משימה" onclick='window.open("members/tasks/task_parents.asp?taskID=<%=taskID%>", "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=<%=cInt(strScreenWidth*0.97)%>,height=480,align=center,resizable=0");' ID="Image11" NAME="Image4">
			<%End If%>
			</td>
	      <td class="card<%=class_%>" dir=rtl valign=top align="right" dir=rtl><a class="link_categ" <%=href%>><%=trim(task_content)%></a></td-->	      	      	      
          <td class="card<%=class_%>" valign=top align="right" dir=rtl><a class="link_categ" <%=href%>><%=companyName%></a></td>
          <td class="card<%=class_%>" valign=top align="right" dir=rtl><a class="link_categ" <%=href%>>ל<%=reciver_name%></a></td>
          <td class="card<%=class_%>" valign=top align="right" dir=rtl><a class="link_categ" <%=href%>>מ<%=sender_name%></a></td>
          <td class="card<%=class_%>" valign=top align="right" dir=rtl><a class="link_categ" <%=href%>><%If isDate(task_date) Then%><%=task_date%><%End If%></a></td>         
	      <td class="card<%=class_%>" valign=top align="center" dir=rtl><a class="task_status_num<%=taskStatus%>" <%=href%>><%=arr_Status(taskStatus)%></a></td>	  
      </tr> 
	  <%
	  i=i+1	
	  loop
	  if NumberOfPages > 1 then
	  urlSort_Out = urlSort_Out & "&task_sort_OUT=" & task_sort_OUT
	  %>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap class="card<%=class_%>">
			<table border="0" cellspacing="0" cellpadding="2">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>	         
	         <tr>
	         <%if numOfRowOut <> 1 then%> 
			 <td valign="center" align="right"><A class=pageCounter title="לדפים הקודמים" href="<%=urlSort_Out%>&pageOUT=<%=10*(numOfRowOut-1)-9%>&numOfRowOut=<%=numOfRowOut-1%>#table_tasks" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRowOut-1)) <= NumberOfPages Then
	                  if CInt(pageOUT)=CInt(i+10*(numOfRowOut-1)) then %>
		                 <td align="middle" align="right"><span class="pageCounter"><%=i+10*(numOfRowOut-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  align="right"><A class=pageCounter href="<%=urlSort_Out%>&pageOUT=<%=i+10*(numOfRowOut-1)%>&amp;numOfRowOut=<%=numOfRowOut%>#table_tasks" ><%=i+10*(numOfRowOut-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td align="right"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRowOut) then%>  
					<td valign="center" align="right"><A class=pageCounter title="לדפים הבאים" href="<%=urlSort_Out%>&pageOUT=<%=10*(numOfRowOut) + 1%>&numOfRowOut=<%=numOfRowOut+1%>#table_tasks" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>
	<%End If%> 
	<tr>
	   <td colspan="10" height=18 class="card<%=class_%>" align=center dir="ltr" style="color:#6E6DA6;font-weight:600">נמצאו <%=recCount%> &nbsp;<%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%> <%=no_search_OUT%></td>
	</tr>
	<%Else%>
	<tr><td colspan=10 class="card<%=class_%>" align=center>&nbsp;לא נמצאו <%=trim(Request.Cookies("bizpegasus")("TasksMulti"))%>&nbsp;<%=no_search_OUT%></td></tr>	
<% End If%>
</table></td></tr>
<%End If%>
</table></td></tr>
</table>

<DIV ID="Status_Popup" STYLE="display:none;">
<div dir="rtl" style="position:absolute; top:0; left:0; width:80; height:82; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;" >
<%For i=1 To uBound(arr_Status)	%>
	<DIV onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand; border-bottom:1px solid black"
	ONCLICK="parent.location.href='<%=UrlStatus%>&task_status=<%=i%>#table_tasks'">
    &nbsp;<%=arr_Status(i)%>&nbsp;
    </DIV>
<%Next%>    
    <DIV onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand;"
	ONCLICK="parent.location.href='<%=UrlStatus%>&task_status=all#table_tasks'">
    &nbsp;כל הסטטוסים&nbsp;
    </DIV>
</div>
</DIV>
<DIV ID="Status_Popup_OUT" STYLE="display:none;">
<div dir="rtl" style="position:absolute; top:0; left:0; width:80; height:82; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;" >
<%For i=1 To uBound(arr_Status)	%>
	<DIV onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand; border-bottom:1px solid black"
	ONCLICK="parent.location.href='<%=UrlStatus_Out%>&task_status_OUT=<%=i%>#table_tasks_OUT'">
    &nbsp;<%=arr_Status(i)%>&nbsp;
    </DIV>
<%Next%>    
    <DIV onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand;"
	ONCLICK="parent.location.href='<%=UrlStatus_Out%>&task_status_OUT=all#table_tasks_OUT'">
    &nbsp;כל הסטטוסים&nbsp;
    </DIV>
</div>
</DIV>

<DIV ID="task_type_Popup" STYLE="display:none;">
<div dir=rtl style="overflow: scroll; position:absolute; top:0; left:0; width:130; height:82; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#D3D3D3;SCROLLBAR-HIGHLIGHT-COLOR: #D3D3D3; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%  sqlstr = "Select activity_type_id, activity_type_name from activity_types WHERE ORGANIZATION_ID = " & OrgID & " Order By activity_type_id"
	set rstask_type = con.getRecordSet(sqlstr)
	while not rstask_type.eof %>
	<DIV dir=ltr onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; text-align:right; height:20px; background:#D3D3D3; border-bottom:1px solid black; cursor:hand;"
	ONCLICK="parent.location.href='<%=urlType%>&tasktypeID=<%=rstask_type(0)%>'">
    &nbsp;<%=rstask_type(1)%>&nbsp;
    </DIV>
	<%
		rstask_type.moveNext
		Wend
		set rstask_type=Nothing
	%>
	<DIV dir="rtl"  onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand;"
	ONCLICK="parent.location.href='<%=urlType%>'">
    &nbsp;כל הרשימה&nbsp;
    </DIV>
</div>
</DIV>

<DIV ID="task_type_Popup_OUT" STYLE="display:none;">
<div dir=rtl style="overflow: scroll; position:absolute; top:0; left:0; width:130; height:82; overflow-x:hidden; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;SCROLLBAR-FACE-COLOR:#D3D3D3;SCROLLBAR-HIGHLIGHT-COLOR: #D3D3D3; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #e5e5e5;" >
<%  sqlstr = "Select activity_type_id, activity_type_name from activity_types WHERE ORGANIZATION_ID = " & OrgID & " Order By activity_type_id"
	set rstask_type = con.getRecordSet(sqlstr)
	while not rstask_type.eof %>
	<DIV dir=ltr onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; text-align:right; height:20px; background:#D3D3D3; border-bottom:1px solid black; cursor:hand;"
	ONCLICK="parent.location.href='<%=urlType_Out%>&tasktypeIDOUT=<%=rstask_type(0)%>'">
    &nbsp;<%=rstask_type(1)%>&nbsp;
    </DIV>
	<%
		rstask_type.moveNext
		Wend
		set rstask_type=Nothing
	%>
	<DIV dir="rtl"  onmouseover="this.style.background='#6E6DA6'" onmouseout="this.style.background='#D3D3D3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#D3D3D3; cursor:hand;"
	ONCLICK="parent.location.href='<%=urlType_Out%>'">
    &nbsp;כל הרשימה&nbsp;
    </DIV>
</div>
</DIV>
</body>
</html>

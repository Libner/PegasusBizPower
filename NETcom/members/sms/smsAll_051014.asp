<%@ Language=VBScript%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<!-- #include file="../../title_meta_inc.asp" -->
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="stylesheet" type="text/css">
<script LANGUAGE="javascript" type="text/javascript">
<!--
var oPopup = window.createPopup();


		function checkSMS(){
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
				str_confirm = "? האם ברצונך לשלוח הודעה גורפת עבור נמענים המסומנים"
			Else
				str_confirm = "Are you sure want to send SMS for the selected recipients?"
			End If   
		    %>
			if (fl && confirm("<%=str_confirm%>")){
				var txtnew = new String(document.form1.trapp.value);
				document.form1.trapp.value = txtnew.substr(0,txtnew.length - 1);
				window.document.all("sms_flag").value = "1";
				h = parseInt(520);
				w = parseInt(520);
				//alert(document.form1.trapp.value)
			window.open("../ContactSMS/addSMSgroupContacts.asp?SmsId=" + document.form1.trapp.value, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
				return true;
			}
			else if (fl) return false;
		}
		  <%
			If trim(lang_id) = "1" Then
				str_confirm = "! נא לסמן נמענים על מנת לשלוח הודעה גורפת"
			Else
				str_confirm = "Please select recipients to send Sms !"
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
	//-->
</script>
	</head>
	<%sqlPer="SELECT  bar_id  FROM bar_users WHERE (user_id =" & UserID &") and bar_id=47 and is_visible=1"
 set rs_Per = con.getRecordSet(sqlPer)
			if not rs_Per.eof then	
			SMS="1"
			else
			SMS="0"
		end if
		set rs_Per= nothing   %>
		<%
		arr_StatusT = Array("","נשלחה","לא נשלחה"," מספר שגוי")	
		
		%>
		
	<%
	app_status = trim(Request("app_status"))
	if trim(Request("orgname")) <> "" then
		orgname = trim(Request("orgname")) 
		where_orgname = " and LOWER(COMPANY_NAME) LIKE '%"& LCase(trim(Request.Form("orgname"))) &"%'"		
	else
		where_orgname = ""
	end if
	if trim(Request("phone")) <> "" then
	phone=trim(Request("phone"))
	else
	phone=""
	end if
	if trim(Request("textcont")) <> "" then
		textcont=trim(Request("textcont"))
	else
		textcont=""
	end if
	if trim(Request("textTitle")) <> "" then
		textTitle=trim(Request("textTitle"))
	else
	textTitle=""
	end if
		
	if trim(Request("clname")) <> "" then
		clname = trim(Request("clname"))
		where_clname = " and LOWER(CONTACT_NAME) LIKE '%"& LCase(trim(Request.Form("clname"))) &"%'"		
	else
		where_clname = ""
	end if
	
	if trim(Request("usname")) <> "" then
		usname = trim(Request("usname"))		
		where_usname = " and LOWER(sender_NAME) LIKE '%"& LCase(trim(Request.Form("usname"))) &"%'"		

	else
		where_usname = ""
	end if
	
 urlSort = "smsAll.asp?status=" & app_status & "&clname=" & Server.URLEncode(clname) & "&orgname=" & Server.URLEncode(orgname) & _
  "&usname=" & Server.URLEncode(usname) & "&orderOwner=" & Server.URLEncode(orderOwner) 


	dim sortby(10)	
	sortby(1) = " date_send, SMS_ID DESC"
	sortby(2) = " date_send DESC, SMS_ID DESC"
	
	sortby(5) = " CONTACT_NAME , SMS_ID DESC"
	sortby(6) = "CONTACT_NAME DESC, SMS_ID DESC"

'	sortby(3) = " StatusName,SMS_ID"
'	sortby(4) = " StatusName DESC,SMS_ID DESC"
'sortby(3) = " U1.FIRSTNAME, U1.LASTNAME, SMS_ID DESC"
	
'	sortby(4) = " U1.FIRSTNAME DESC, U1.LASTNAME DESC, SMS_ID DESC"
	
	sortby(3) = " COMPANY_NAME, SMS_ID DESC"
	sortby(4) = " COMPANY_NAME DESC, SMS_ID DESC"
 
	sort = Request("sort")	
	if sort = nil then
		sort = 2
	end if





' ------ delete appeals --------

	
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
	<body>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>" ID="Table1">
			<tr>
				<td width="100%">
					<!-- #include file="../../logo_top.asp" -->
				</td>
			</tr>
			<%numOfTab = 46%>
			<%numOfLink = 4%>
			<tr>
				<td width="100%">
					<!--#include file="../../top_in.asp"-->
				</td>
			</tr>
			<tr><td class="page_title" style="border-right:none;height:25px" ></td></tr>
			<tr>
				<td valign="top">
					<FORM action="smsAll.asp?sort=<%=sort%>" method=POST id="form1" name="form1" target="_self">
						<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" ID="Table3">
							<tr>
								<td width="100%" valign="top" align="center">
									<table width="100%" cellspacing="0" cellpadding="0" align="center" border="0" bgcolor="#ffffff"
										ID="Table4">
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
'response.Write sortby(sort) 

'sqlstr = "exec dbo.get_SMS_All"
 ' response.Write app_status
  sqlstr = "exec dbo.get_SMS_All_padding @Page=" & Page & ", @RecsPerPage=" & PageSize & ", @company_Name ='" & sFix(orgname) & _
  "', @contact_name='" & sFix(clname) & "', @SMS_STATUS='" & app_status & "', @phone='" & phone & _
 "', @sender_Name='" & sFix(usname) & _
  "', @sort='" & sortby(sort) &"',@SmsType='"& SmsType &"',@textcont='"& textcont &"',@textTitle='" &textTitle &"'"
  '  "', @company_id='" & CompanyId & "', @contact_id='" & ContactId & "', @project_id='" & ProjectId &_
   ' "', @appeal_id='" & search_id & "', @product_id='" & prodId & "', @UserId='" & UserID & "', @archive='" & arch & _
   ' "', @user_name='" & sFix(usname) & "', @orderOwner='" & sFix(orderOwner) & "', @IsGroups='" & is_groups & _
   ' "', @str_where_values='" & sFix(str_where_values) & "'"	
   'Response.Write sqlStr
    'Response.End
	set app=con.GetRecordSet(sqlStr)	%>

												<!-- start search row -->
												<input type="hidden" name="trapp" value="" ID="trapp">

												<input type="hidden" name="sms_flag" value="0" ID="sms_flag"> 
												<input type="hidden" name="change_status_flag" value="0" ID="change_status_flag">
												<table border="0" cellspacing="1" cellpadding="0" width="100%" ID="Table5">
													<tr style="background-color: #8A8A8A">
														<td dir="<%=dir_obj_var%>" align="left"  style="padding-left: 5px;"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:280px;" value="<%=vFix(textcont)%>" name="textcont" ID="textcont" maxlength="200">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input type="submit" value="חפש" class="but_menu" style="width: 50px" ID="btnSubmit" NAME="btnSubmit">&nbsp;&nbsp;<input type="button" 
		value="הצג הכל" class="but_menu" style="width: 60px" onclick="document.location.href='smsAll.asp?prodId=<%=prodId%>&sort=<%=sort%>&arch=<%=arch%>'" ID="btnShowAll" NAME="btnShowAll"</td>
														<td width="85"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:90px;" value="<%=vFix(usname)%>" name="usname" ID="usname" maxlength="50"></td>
														<td width="100" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:120px;" value="<%=vFix(textTitle)%>" name="textTitle" ID="textTitle" maxlength="50"></td>
														<td width="100" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:90px;" value="<%=vFix(phone)%>" name="phone" ID="phone" maxlength="50"></td>
														<td width="250" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:90px;" value="<%=vFix(clname)%>" name="clname" ID="clname" maxlength="50"></td>
														<%if false then%><td  width="250"  align="<%=align_var%>" dir="<%=dir_obj_var%>"><input type="text" class="search" dir="<%=dir_obj_var%>" style="width: 150px;" value="<%=vFix(orgname)%>" name="orgname" ID="orgname" maxlength="50"></td><%end if%>
														<td></td>
														<td ><select ID="app_status" dir="<%=dir_obj_var%>" class="search" style="width: 65px"
		onchange="document.location.href='smsAll.asp?sort=<%=sort%>&app_status=' + this.value;" NAME="app_status">	
		<option dir="<%=dir_obj_var%>" value="" ><!--הכל--><%=arrTitles(19)%></option>
		<%For i=1 To Ubound(arr_StatusT)%>
		<option value="<%=i%>" <%If trim(arr_StatusT(i)) = trim(app_status) Then%> selected <%End If%> ><%=arr_StatusT(i)%></option>
		<%Next%>    
		</select></td><td></td>
													</tr>
													<!-- end search row -->
												<tr bgcolor="<%=tr_bgcolor%>" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='<%=tr_bgcolor%>';">
			<td width="250" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="right" style="padding: 0px 3px;">תוכן</td>
														<td width="55" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="right" style="padding: 0px 3px;">מעת</td>
														<td width="55" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="right"  style="padding: 0px 3px;">שדה 
															כותרת</td>
														<td width="65" nowrap class="title_sort" dir="<%=dir_obj_var%>" align="right" style="padding: 0px 3px;">טלפון 
															נייד</td>
														<td width="85" nowrap align="r" dir="<%=dir_obj_var%>" class="title_sort<%if sort=5 or sort=6 then%>_act<%end if%>"><a class="title_sort" title="<%=arrTitles(25)%>" HREF="<%=urlSort%>&sort=<%if sort=6 then%>5<%elseif sort=5 then%>6<%else%>6<%end if%>" target="_self">איש קשר<img src="../../images/arrow_<%if trim(sort)="5" then%>down<%elseif trim(sort)="6" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
														<%if false then%><td width="85"  nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>"><a class="title_sort" title="מיון" HREF="<%=urlSort%>&sort=<%if sort=4 then%>3<%elseif sort=3 then%>4<%else%>4<%end if%>" target="_self">&nbsp;שם חברה<img src="../../images/arrow_<%if trim(sort)="3" then%>down<%elseif trim(sort)="4" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td><%end if%>
														<td width="50" nowrap align="right" dir="rtl" style="padding: 0px 3px;" class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" title="<%=arrTitles(25)%>" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self">תאריך<img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
														<td width="48" nowrap  class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>" >&nbsp;סטטוס&nbsp;</td>
														<td width="20" align="right" dir="rtl" valign="top" class="title_sort" style="padding: 0px 3px;"><INPUT type="checkbox" LANGUAGE="javascript" onclick="return cball_onclick()" title="<%=arrTitles(20)%>" id="cb_all" name="cb_all"></td>
<%	if not app.eof then %>
													</tr>
<%
	
recCount = app("CountRecords")
 aa = 0
	 do while not app.eof
		If aa Mod 2 = 0 Then
			tr_bgcolor = "#E6E6E6"
		Else
			tr_bgcolor = "#C9C9C9"
		End If	
		Smsid=app("Sms_id")	
	ids = ids & Smsid 		
		
		%>
	<tr bgcolor="<%=tr_bgcolor%>" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='<%=tr_bgcolor%>';">
													
														<td align="right" dir="rtl" valign="top" style="padding: 0px 3px;"><a class="title_sort" href="../companies/contact.asp?companyId=<%=app("company_id")%>&contactId=<%=app("Contact_id")%>"><%=app("sms_Content")%></a></td>
														<td align="right" valign="top" style="padding: 0px 3px;"><a class="title_sort" href="../companies/contact.asp?companyId=<%=app("company_id")%>&contactId=<%=app("Contact_id")%>"><%=app("SENDER_NAME")%></a></td>
														<td align="<%=align_var%>" valign="top" style="padding: 0px 3px;"><a class="title_sort" href="../companies/contact.asp?companyId=<%=app("company_id")%>&contactId=<%=app("Contact_id")%>"><%=app("TitleName")%></a></td>
														<td align="right" valign="top" style="padding: 0px 3px;"><a class="title_sort" href="../companies/contact.asp?companyId=<%=app("company_id")%>&contactId=<%=app("Contact_id")%>"><%=app("contact_CELL")%></a></td>
														<td width =250  align="right" dir="rtl" valign="top" style="padding: 0px 3px;"><a class="title_sort" href="../companies/contact.asp?companyId=<%=app("company_id")%>&contactId=<%=app("Contact_id")%>"><%=app("CONTACT_NAME")%> </a></td>
														<%if false then%><td width=250  align="right" dir="rtl" valign="top" style="padding: 0px 3px;"><a class="title_sort" href="../companies/contact.asp?companyId=<%=app("company_id")%>&contactId=<%=app("Contact_id")%>"><%=app("COMPANY_NAME")%></a></td><%end if%>
														<td width=60  align="right" dir="rtl" valign="top" style="padding: 0px 3px;"><a class="title_sort" href="../companies/contact.asp?companyId=<%=app("company_id")%>&contactId=<%=app("Contact_id")%>"><%=day(app("date_send"))%>/<%=month(app("date_send"))%>/<%=mid(year(app("date_send")),3,2)%></a></td>
														<td align="center"  align="right" dir="rtl" valign="top"  <%if app("StatusName")="נשלחה" then%>style="padding: 0px 3px;background-color:#66C936" <%elseif app("StatusName")="לא נשלחה" then%> style="padding: 0px 3px;background-color:#FF0000" <%else%>style="padding: 0px 3px;"<%end if%>><%=app("StatusName")%></td>
														<td align="center"  align="right" dir="rtl" valign="top" style="padding: 0px 3px;"><INPUT type="checkbox" id="cb<%=app("Sms_Id")%>" name="cb<%=app("Sms_Id")%>"></td>
</tr>
		<%	app.movenext
		aa = aa + 1
			if not app.eof then
		ids = ids & ","
		end if
		loop%>
												</table>
												<input type="hidden" name="ids" value="<%=ids%>" ID="ids"></td>
										</tr>
										<%	NumberOfPages = Fix((recCount / PageSize)+0.9)	%>
										<%	'End If 	%>
										<% if NumberOfPages > 1 then%>
										<tr class="card">
											<td width="100%" align="center" nowrap class="card" dir="ltr">
												<table border="0" cellspacing="0" cellpadding="2" ID="Table6">
													<% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRow") <> nil Then
	               numOfRow = Request.QueryString("numOfRow")
	           Else numOfRow = 1
	           End If %>
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
										<%End If
										end if

    set app = nothing%>
									</table>
								</td>
								<td width="100" nowrap align="<%=align_var%>" valign="top" class="td_menu">
									<table cellpadding="2" cellspacing="0" width="100%" border="0" ID="Table7">
										<tr>
											<td colspan="2" height="50" nowrap></td>
										</tr>
										<%if trim(SMS)="1" then%>
										<tr>
											<td colspan="2" align="right"><a class="button_edit_2" style="width:90px; line-height:110%; padding:3px" href='javascript:void(0)'
													onclick="javascript:checkSMS()">SMS שלח</a></td>
										</tr>
								</td>
							</tr>
							<%end if%>
							<tr>
								<td colspan="2" height="10" nowrap></td>
							</tr>
						</table>
				</td>
			</tr>
		</table>
		</form></td></tr> </table> </center></div>
	</body>
</html>
<%set con=nothing%>

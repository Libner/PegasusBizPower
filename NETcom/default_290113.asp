<!-- מסך משתמש ראשי, INCLUDE עליון מסך סטאטוס הזמנות
נתונים: תאריכון פגישות, חיפוש אנשי קשר, חיפוש טפסים, משימות נכנסות, משימות יוצאות -->
<!-- #include file="connect.asp" -->
<!-- #include file="reverse.asp" -->
<!-- #include file="members/checkWorker.asp" -->
<%
	If trim(lang_id) = "1" Then
		Session.LCID = 1037 
	Else
		Session.LCID = 2057
	End If
  
	'משתמש מורשה להוסיף פגישות לאחרים
	sqlstr = "Select IsNULL(Add_Meetings,0) From Users Where User_ID = " & UserID
	set rs_check = con.getRecordSet(sqlstr)
	if not rs_check.eof Then
		AddMeetings = trim(rs_check.Fields(0))
	else
		AddMeetings = 0
	end if			

	If AddMeetings = "1" Then
		participant_id = ""
	Else
		participant_id = UserID	
	End If
	
' ---------- Page Variables ----------
    Const intCharToShow = 19		' The number of characters shown in each day
    Const bolEditable   = True		' If the calendar is editable or not (Can be tied into password verification)

    Dim dtToday 			' Today's meeting_date
    Dim dtCurrentDate			' The current meeting_date
    Dim aCalendarDays(42)		' Array of possible calendar dates
    Dim iFirstDayOfMonth		' The first day of the month
    Dim iDaysInMonth	 		' The number of days in the month
    Dim iColumns, iRows	, iDay, iWeek	' The numer of columns and rows in the table, and counters to print them
    Dim objConn, strConn, strSQL, objRS ' Database Variables
    Dim counter 			' Loop counter
    Dim strNextMonth, strPrevMonth	' The next and previous month dates
    Dim dailyMsg			' The message for the day
    Dim current_date				' The current day being displayed by the loops
    Dim strPage				' The link that each day takes you too		
	
    dtToday = Date()
    date_ = Date()
    If Request("date_") <> nil Then
		date_ = Request("date_")
	End If
	
	If Request("currentMonth") = "" And Request("currentYear") = ""  Then	 
		currentMonth = Month(date_)
		currentYear = Year(date_)
	Else
		currentMonth = Request("currentMonth")
		currentYear = Request("currentYear")	
	End If

    If currentMonth <> "" And currentYear <> ""  Then    
		If Request("date_") = nil Then
			If trim(currentMonth) <> cStr(Month(Date())) Then
				dtCurrentDate = "1" & "/" & currentMonth & "/" &  currentYear
			Else
				dtCurrentDate = Day(Date()) & "/" & currentMonth & "/" &  currentYear
			End If	
		Else
			dtCurrentDate = date_
		End If
    Else
         dtCurrentDate = dtToday
    End If    

    iFirstDayOfMonth = DatePart("w", DateSerial(Year(dtCurrentDate), Month(dtCurrentDate), 1))
    iDaysInMonth = DatePart("d", DateSerial(Year(dtCurrentDate), Month(dtCurrentDate)+1, 1-1))

    For counter = 1 to iDaysInMonth
      aCalendarDays(counter + iFirstDayOfMonth - 1) = counter
    Next

    iColumns = 7
    iRows    = 6 - Int((43 - (iFirstDayOfMonth + iDaysInMonth)) / 7)    
    
    'Response.Write (43 - (iFirstDayOfMonth + iDaysInMonth)) / 7

    strPrevMonth = Server.URLEncode(DateAdd("m", -1, dtCurrentDate))
    strNextMonth = Server.URLEncode(DateAdd("m",  1, dtCurrentDate))
    
    If trim(lang_id) = "1" Then
    daysname = array(" 'א"," 'ב"," 'ג"," 'ד"," 'ה"," 'ו","שבת")
    Else
    daysname = array("S","M","T","W","T","F","S")
    End If    	

	  
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 1 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
	arrTitlesD = rstitle.getRows()
	redim arrTitles(Ubound(arrTitlesD,2)+2)
	For i=0 To Ubound(arrTitlesD,2)		
		arrTitles(arrTitlesD(0,i)) = arrTitlesD(1,i)		
	Next
	End If
	set rstitle = Nothing	
%> 	
<html>
<head>
<!-- #include file="title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="IE4.css" rel="STYLESHEET" type="text/css">
<STYLE TYPE="text/css">
    <!--   
    .blackBacking   {background-color: #000000;}
    .names 	    {background-color: #CCCCCC; font-size: 13px; color: #000000; text-decoration: none; text-align:  center;  font-weight: bold; border-top : 1px solid black; border-right : 1px solid black; line-height: 150%}
    .calendarBody   {background-color: #F0F0F0; font-size: 12px; color: #000000; text-decoration: none; text-align:  center; border : 1px solid black; border-right : none;}
    .calCurrentDay  {background-color: #F0F0F0; font-size: 12px; color: #000000; text-align: center; font-weight:bold; border: solid #FF5959 2px}
    .calOtherDay    {background-color: #F0F0F0; font-size: 12px; color: #000000; text-align:center; line-height : 140%; font-weight:bold; border-top : 1px solid black; border-right : 1px solid black}
    .calNotDay	    {background-color: #F0F0F0; font-size: 12px; color: #000000; line-height : 140%; text-align:  center; border-top : 1px solid black; border-right : 1px solid black}
    .calFormMenu    {background-color: #4C5D87; font-size: 13px; color: #FFFFFF; text-decoration: none; text-align:  center;  font-weight: bold; border : 1px solid black}
    -->
</STYLE>
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
		oPopup.show(0, 17, 70, 82, obj);    
	}
	
	var oPopupOut = window.createPopup();
	function StatusDropDownOut(obj)
	{
		oPopupOut.document.body.innerHTML = Status_Popup_OUT.innerHTML; 
		oPopupOut.document.charset = "windows-1255";
		oPopupOut.show(0, 17, 70, 82, obj);    
	}
  
	function closeTask(contactID,companyID,taskID)
	{
		h = parseInt(490);
		w = parseInt(470);
		window.open("members/tasks/closetask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskId=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=50,left=100,width="+w+",height="+h+",align=center,resizable=0");
	}
	
	function addtask(contactID,companyID,taskID)
	{
		h = parseInt(530);
		w = parseInt(470);
		window.open("members/tasks/addtask.asp?contactID=" + contactID + "&companyId=" + companyID + "&taskID=" + taskID, "T_Wind" ,"scrollbars=1,toolbar=0,top=20,left=120,width="+w+",height="+h+",align=center,resizable=0");		
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
<body>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">        
	<tr><td width=100% colspan=3>
	<!--#include file="logo_top.asp"-->
	</td></tr>
	<tr>				
			<!--td align="<%=align_var%>" valign="top" width="255" nowrap bgcolor="#DEDFF7">
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
				<table border="0" width="100%" cellpadding="0" cellspacing="0" dir=<%=dir_var%>>												
				<tr>		
				<%If trim(SURVEYS) = "1" Then%>					
					<td bgcolor="#FFFFFF" width="<%If trim(COMPANIES) = "1" Then%>50%<%Else%>100%<%End If%>" align="<%=align_var%>" valign=top>
					<table border="0" width="<%If trim(COMPANIES) = "1" Then%>100%<%Else%>50%<%End If%>" cellspacing="0" cellpadding="0" align="<%=align_var%>">							
						<tr>
							<td align="right">					
							<table border="0" width="100%" cellspacing="0" cellpadding="0">
							<tr>
								<td align="center" dir="<%=dir_var%>" class="title_form">
								<!--הצג טפסים וסקרים--><%=arrTitles(3)%>
								</td>
							</tr>																		
							<tr>
								<td align="center" valign=top>
								<table border="0" width=100% cellpadding="4" cellspacing="0">																
								<FORM action="members/appeals/appeal.asp" method=POST id="form_tofes" name=form_tofes target="_self">   
								<tr>
								<td align="<%=align_var%>" width=5% nowrap class=card></td>
								<td align="<%=align_var%>" width=70% nowrap class=card>
								<SELECT dir="<%=dir_obj_var%>" ID="quest_id" name=quest_id class="form_text" style="width: 99%" onchange="form_tofes.submit();">	
								<OPTION value=""><!--בחר טופס--><%=arrTitles(4)%></OPTION>
								<%
									If is_groups = 0 Then
									sqlstr = "SELECT product_id, product_name FROM Products WHERE (product_number = '0')  "&_
									" AND (SHOW_INSITE = 1) AND (ORGANIZATION_ID=" & OrgID & ") order by product_name"
									Else
									'sqlstr = "Select DISTINCT Products.product_id, Products.product_name from Products Inner Join Users_To_Products "&_
									'" ON Products.Product_ID = Users_To_Products.Product_ID WHERE Users_To_Products.User_ID = " & UserID &_
									'" And Products.product_number = '0' and Products.ORGANIZATION_ID=" & OrgID & " order by Products.product_name"
									sqlstr = "Execute get_products_list_enter '" & OrgID & "','" & UserID & "'"
									End If
									'Response.Write sqlstr
									'Response.End
									set rs_products = con.GetRecordSet(sqlstr)										
									if not rs_products.eof then 
										ProductsList = rs_products.getRows()		
									end if
									set rs_products=nothing				
									If IsArray(ProductsList) Then
									For i=0 To Ubound(ProductsList,2)
										prod_Id = ProductsList(0,i)   	
										product_name = ProductsList(1,i)										
								%>
								<OPTION value="<%=prod_Id%>" <%If trim(prodId) = trim(prod_Id) Then%> selected <%End If%>> <%=product_name%> </OPTION>
								<%	Next	
									End If	
								%>
								</SELECT>
								</td>
								<td align="<%=align_var%>" width=25% nowrap class=card style="padding-right:10px;padding-left:10px"><span id=word5 name=word5><!--מלא טופס--><%=arrTitles(5)%></span></td>
								</tr>
								</FORM>							
								<tr><td height=1 nowrap colspan=3></td></tr>
								<FORM action="members/appeals/appeals.asp" method=POST id="form_search_tofes1" name=form_search_tofes1 target="_self">   
								<tr>
								<td align="<%=align_var%>" width=5% nowrap class=card></td>
								<td align="<%=align_var%>" width=70% nowrap class=card>
								<SELECT dir="<%=dir_obj_var%>" ID="prodId" name=prodId class="form_text" style="width: 99%" onchange="form_search_tofes1.submit();" >	
								<OPTION value="" id=word6 name=word6><!--בחר טופס--><%=arrTitles(6)%></OPTION>
								<%
								If is_groups = 0 Then
								sqlstr = "Select product_id, product_name from Products Where "&_
								" product_number = '0' and ORGANIZATION_ID=" & OrgID & " order by product_name"
								Else
								sqlstr = "Execute get_products_list '" & OrgID & "','" & UserID & "'"
								End If																						
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
									End If	%>
								</SELECT>
								</td>
								<td align="<%=align_var%>" width=25% nowrap class=card style="padding-right:10px;padding-left:10px"><span id=word7 name=word7><!--טפסים מלאים--><%=arrTitles(7)%></span></td>
								</tr>
								</FORM>							
								<%'end if%>
								<%If trim(EMAILS) = "1" Then%>
								<tr><td height=1 nowrap colspan=3></td></tr>
								<FORM action="members/appeals/feedbacks.asp" method=POST id="form_search_tofes" name=form_search_tofes target="_self">   
								<tr>
								<td align="<%=align_var%>" width=5% nowrap class=card></td>
								<td align="<%=align_var%>" width=70% nowrap class=card>
								<SELECT dir="<%=dir_obj_var%>" ID="Select1" name=prodId class="form_text" style="width: 99%" onchange="form_search_tofes.submit();">	
								<OPTION value=""><!--בחר טופס--><%=arrTitles(8)%></OPTION>
								<%
									sqlstr = "Select product_id, product_name from Products Where FORM_MAIL = '1' And "&_
									" product_number = '0' AND ORGANIZATION_ID=" & OrgID & " order by product_name"
									'Response.Write sqlstr
									'Response.End
									set rs_products = con.GetRecordSet(sqlstr)
									if not rs_products.eof then 
										FeedBackProductsList = rs_products.getRows()		
									end if
									set rs_products=nothing				
									If IsArray(FeedBackProductsList) Then
									For i=0 To Ubound(FeedBackProductsList,2)
										prod_Id = FeedBackProductsList(0,i)   	
										product_name = FeedBackProductsList(1,i)											
									%>
									<OPTION value="<%=prod_Id%>" <%If trim(prodId) = trim(prod_Id) Then%> selected <%End If%>> <%=product_name%> </OPTION>
									<%	Next	
									End If%>
								</SELECT>
								</td>
								<td align="<%=align_var%>" width=25% nowrap class=card style="padding-right:10px;padding-left:10px"><span id=word9 name=word9><!--משובים מדיוור--><%=arrTitles(9)%></span></td>
								</tr>
								</FORM>
								<%End If%>
							<%If trim(is_meetings) = "1" Then%>	
							<tr><td height=1 nowrap colspan=3></td></tr>	
							<tr><td class="card" height=77 colspan=3></td></tr>																				
							<%End If%>
						</table></td></tr>						
						</table></td></tr>						 								
					</table></td>	
					<%End If%>						
					<td width=10 nowrap></td>							
					<%If trim(COMPANIES) = "1" Then%>												
					<td align="<%=align_var%>" width="<%If trim(SURVEYS) = "1" Then%>50%<%Else%>100%<%End If%>" nowrap valign=top>															
						<table border="0" width="<%If trim(SURVEYS) = "1" Then%>100%<%Else%>50%<%End If%>" cellspacing="0" cellpadding="0" dir=<%=dir_var%>>	
						<tr>
						<td bgcolor=white width="100%" align="center" valign=top>
						<table border="0" width="100%" cellspacing="0" cellpadding="0">
						<tr>
							<td align="right" valign=top>					
							<table border="0" width="100%" cellspacing="0" cellpadding="0">
								<tr>
									<td align="center" dir="<%=dir_var%>" class="title_form">
									<!--הצג תיקי--><%=arrTitles(10)%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesMulti"))%>
									</td>
								</tr>																								
								<tr>
									<td align="right">
										<table width="100%" border="0" cellpadding="4" cellspacing="0">							
										<FORM action="members/companies/default.asp" method=POST id="form_search" name=form_search target="_self">   
										<tr>
										<td align="<%=align_var%>" width=15% nowrap class=card><input type="image" onclick="form_search.submit();" src="images/search.gif" ID="Image1" NAME="Image1"></td>
										<td align="<%=align_var%>" width=60% nowrap class=card><input type="text" class="form_text" dir="<%=dir_obj_var%>" style="width:100%;" value="<%=vFix(search_company)%>" name="search_company" ID="search_company"></td>
										<td align="<%=align_var%>" width=25% nowrap class=card style="padding-right:10px;padding-left:10px"><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></td>
										</tr>
										</FORM>
										<tr><td height=1 nowrap colspan=3></td></tr>
										<FORM action="members/projects/default.asp" method=POST id="form_search_project" name="form_search_project" target="_self">   
										<tr>
										<td align="<%=align_var%>" width=15% nowrap class=card><input type="image" onclick="form_search.submit();" src="images/search.gif" ID="Image2" NAME="Image2"></td>
										<td align="<%=align_var%>" width=60% nowrap class=card><input type="text" class="form_text" dir="<%=dir_obj_var%>" style="width:100%;" value="<%=vFix(search_project)%>" name="search_project" ID="search_project"></td>
										<td align="<%=align_var%>" width=25% nowrap class=card style="padding-right:10px;padding-left:10px"><%=trim(Request.Cookies("bizpegasus")("Projectone"))%></td>
										</tr>
										</FORM>
										<tr><td height=1 nowrap colspan=3></td></tr>
										<FORM action="members/companies/default.asp" method=POST id="form_search_contact" name="form_search_contact" target="_self">   
										<tr>
										<td align="<%=align_var%>" width=15% nowrap class=card><input type="image" onclick="form_search.submit();" src="images/search.gif" ID="Image6" NAME="Image2"></td>
										<td align="<%=align_var%>" width=60% nowrap class=card><input type="text" class="form_text" dir="<%=dir_obj_var%>" style="width:100%;" value="<%=vFix(search_contact)%>" name="search_contact" ID="search_contact"></td>
										<td align="<%=align_var%>" width=25% nowrap class=card style="padding-right:10px;padding-left:10px"><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td>
										</tr>
										</FORM>	
							<%If trim(is_meetings) = "1" Then%>			
							<tr><td height=1 nowrap colspan=3></td></tr>	
							<tr><td class="card" height=77 colspan=3></td></tr>																						
							<%End If%>
							</table></td></tr>
					</table></td></tr></table></td></tr></table></td>
					    <%End If%>
						<%If trim(is_meetings) = "1" Then%>	
						<td bgcolor="#FFFFFF" width=10 nowrap></td>
						<td align="right" valign=top>							
						<table align="right" border="0" width="170" cellspacing="0" cellpadding="0" dir=<%=dir_var%> ID="Table1">
						<tr>
						<td align="center" dir="<%=dir_var%>" class="title_form">
						<!--הצג פגישות--><%=arrTitles(43)%>&nbsp;
						</td>
						</tr>					
						<tr>    
							<td width="100%" valign="top" align="<%=align_var%>" class="card">
							<table cellpadding=3 cellspacing=2 width=170 align=center border=0  dir="<%=dir_obj_var%>" ID="Table2">
							<FORM NAME="form_calendar" ACTION="default.asp" METHOD="POST" ID="form_calendar">
							<tr>		
								<td align=center nowrap>
								<SELECT NAME="currentMonth" CLASS="form_text" ID="currentMonth" dir="<%=dir_var%>" style="width:95" onchange="form_calendar.submit();">
								<% For counter = 1 to 12 %>
									<OPTION VALUE="<%=counter%>" <% If (DatePart("m", dtCurrentDate) = counter) Then Response.Write "SELECTED"%>><%=MonthName(counter)%></OPTION>
								<% Next %>
								</SELECT>
								</td>			   	
								<td align="center" nowrap>
								<SELECT NAME="currentYear" CLASS="form_text" ID="currentYear" dir="<%=dir_var%>" style="width:60" onchange="form_calendar.submit();">
								<% For counter = -2 to 2 %>
									<OPTION VALUE="<%=Year(dtCurrentDate)+counter%>" <% If (DatePart("yyyy", dtCurrentDate) = Year(dtCurrentDate)+counter) Then Response.Write "SELECTED"%>><%=Year(dtCurrentDate)+counter%></OPTION>
								<% Next %>
								</SELECT>
								</td>		
							</TR>		
							</table>
							</td>
						</tr>					  	
						<tr>
							<td class="card" width="100%" align="center">
							<TABLE CELLSPACING=0 CELLPADDING=0 BORDER=0 WIDTH=160 dir="<%=dir_obj_var%>" class="calendarBody" ID="Table3">
  							<tr><th colspan=7 align=center style="border-right : 1px solid black"><%=MonthName(Month(dtCurrentDate))%>&nbsp;<%=Year(dtCurrentDate)%></th></tr>
							<!-- Writring the days of the week for headers -->
								<TR VALIGN="TOP" ALIGN="CENTER" dir=ltr>
								<% For iDay = vbSunday To vbSaturday%>
									<TD WIDTH="14%" class="names"><%=daysname(iDay-1)%></TD>
								<% Next %>
								</TR>
							<%   
								For iWeek = 1 To iRows
								Response.Write "<TR VALIGN=TOP>"
								For iDay = 1 To iColumns
								' Checks to see if there is a day this month on the meeting_date being written
								If aCalendarDays((iWeek-1)*7 + iDay) > 0 then
									current_date = DateSerial(Year(dtCurrentDate), Month(dtCurrentDate), aCalendarDays((iWeek-1)*7 + iDay))	  							
									current_date = Day(current_date) & "/" & Month(current_date) & "/" & Year(current_date)
									strPage = "members/meetings/default.asp?date_=" & current_date & "&currentMonth=" & currentMonth & "&currentYear=" & currentYear & "&participant_id=" & participant_id	  
								 
									meeting_count = 0
									style = "style=""background-color:#F0F0F0; """ 		
									meetings_str = "<span style='border-top:1px dashed #000000;width:100%;height:50%;font-weight:500'>0"
									meetings_str = meetings_str & "<span style='font-size:11px'> "&arrTitles(1)&"</span></A>"		
								  										
									sqlStr = "SET DATEFORMAT DMY; Select meeting_date,Count(meeting_id) from meetings_view where ORGANIZATION_ID = "& OrgID
									If trim(participant_id) <> "" Then
										sqlStr = sqlStr & " AND participant_id = " & participant_id
									End If
									sqlStr = sqlStr & " And meeting_date = '" & current_date & "' Group BY meeting_date"
									'Response.Write sqlStr
									'Response.End
									set rs_meetings = con.GetRecordSet(sqlStr)
									if not rs_meetings.EOF then									
										meeting_count = rs_meetings(1)
										date_pr = rs_meetings(0) 			
									
										meetings_str = "<span style='border-top:1px dashed #000000;width:100%;height:50%;font-weight:500'>" & meeting_count & "<span style='font-size:11px'> "&arrTitles(1)&"</span></span>"
										if meeting_count < 1 then
											style = "style=""background-color:#F0F0F0; """ 			
										else
											style = "style=""background-color:#53FE01; """ 
										end if
								end if
								set rs_meetings = nothing	
								   
								' Checks to see if the day being printed is today
								If DateDiff("d",current_date,dtCurrentDate) = 0 Then
									Response.Write "<TD CLASS='calCurrentDay'" 
 								Else
   									Response.Write "<TD CLASS='calOtherDay'" 
 								End If	  
							 	 
 								Response.Write " style=""cursor:hand"" onClick=""document.location.href='" & strPage & "'""" 	  
 								Response.Write style & ">" 	  
 								Response.Write ("&nbsp;" & aCalendarDays((iWeek-1)*7 + iDay))
							 	   	  
								Else 
								Response.Write ("<TD CLASS='calNotDay'>&nbsp;")
								End IF

								Response.Write "</TD>"
								Next
								Response.Write "</TR>"
								Next  
							%>							
							</TABLE></td>
							<%If trim(AddMeetings) = "1" Then%>  
							<%Else%>
							<input type=hidden name="participant_id" id="participant_id" value="<%=participant_id%>">	
							<%End If%>   
							</form>
							</tr>					
							</table>					
							</td>	
							<%End If%>							
							</tr>							
						
  </table></td></tr>	
 <%	If trim(TASKS) = "1" Then %>
	<tr><td height=2 nowrap bgcolor="#FFFFFF"></td></tr>
 <%
    dim arr_Status(3)
	If trim(lang_id) = "1" Then
	arr_Status(1)="פתוחה"
	arr_Status(2)="בטיפול"
	arr_Status(3)="סגורה"
	Else
	arr_Status(1)="new"
	arr_Status(2)="active"
	arr_Status(3)="close"	
	End If
    
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

	If lang_id = 1 Then
	    Select Case(status)	    
			Case "1" : no_search = "פתוחות"
			Case "2" : no_search = "בטיפול" 
			Case "1,2" : no_search = "פתוחות או בטיפול"
			Case "3" : no_search = "סגורות" 
			Case "all" : no_search = ""		
	    End Select
	Else
	   Select Case(status)	    
			Case "1" : no_search = "new"
			Case "2" : no_search = "active" 
			Case "1,2" : no_search = "new or active"
			Case "3" : no_search = "close" 
			Case "all" : no_search = ""
		End Select	
	End If

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
   
	<tr><td height=2 nowrap bgcolor="#FFFFFF"></td></tr>


<%End If%>
</table></td></tr>
</table>
<%	If trim(TASKS) = "1" Then %>
<DIV ID="Status_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; width:70; height:82; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;" >
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
    &nbsp;<span id=word34 name=word34><!--הכל--><%=arrTitles(34)%></span>&nbsp;
    </DIV>
</div>
</DIV>
<DIV ID="Status_Popup_OUT" STYLE="display:none;">
<div dir=<%=dir_obj_var%> style="position:absolute; top:0; left:0; width:70; height:82; border-bottom:1px solid black; border-top:1px solid black; border-left:1px solid black; border-right:1px solid black;" >
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
     &nbsp;<span id="word35" name=word35><!--הכל--><%=arrTitles(35)%></span>&nbsp;
    </DIV>
</div>
</DIV>
<%End If%>
</body>
</html>

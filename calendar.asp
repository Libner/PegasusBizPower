<%
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
    
    daysname = array(" 'א"," 'ב"," 'ג"," 'ד"," 'ה"," 'ו","שבת")
%>
<html>
<head>
<title>calendar</title>
<meta name=vs_defaultClientScript content="JavaScript">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
<meta name=ProgId content=VisualStudio.HTML>
<meta name=Originator content="Microsoft Visual Studio .NET 7.1">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
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
</head>
<body MS_POSITIONING="GridLayout">
<table cellpadding=3 cellspacing=2 width=170 align=center border=0  dir="rtl">
							<FORM NAME="form_calendar" ACTION="default.asp" METHOD="POST" ID="form_calendar">
							<tr>		
								<td align=center nowrap>
								<SELECT NAME="currentMonth" CLASS="form_text" ID="currentMonth" dir="ltr" style="width:95" onchange="form_calendar.submit();">
								<% For counter = 1 to 12 %>
									<OPTION VALUE="<%=counter%>" <% If (DatePart("m", dtCurrentDate) = counter) Then Response.Write "SELECTED"%>><%=MonthName(counter)%></OPTION>
								<% Next %>
								</SELECT>
								</td>			   	
								<td align="center" nowrap>
								<SELECT NAME="currentYear" CLASS="form_text" ID="currentYear" dir="ltr" style="width:60" onchange="form_calendar.submit();">
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
							<TABLE CELLSPACING=0 CELLPADDING=0 BORDER=0 WIDTH=160 dir="rtl" class="calendarBody" align=center>
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
									meetings_str = meetings_str & "<span style='font-size:11px'>יום</span></A>"		
								  										
									'sqlStr = "SET DATEFORMAT DMY; Select meeting_date,Count(meeting_id) from meetings_view where ORGANIZATION_ID = "& OrgID
									'If trim(participant_id) <> "" Then
									'	sqlStr = sqlStr & " AND participant_id = " & participant_id
									'End If
									'sqlStr = sqlStr & " And meeting_date = '" & current_date & "' Group BY meeting_date"
									'Response.Write sqlStr
									'Response.End
									'set rs_meetings = con.GetRecordSet(sqlStr)
									'if not rs_meetings.EOF then									
										'meeting_count = rs_meetings(1)
										'date_pr = rs_meetings(0) 			
									
										meetings_str = "<span style='border-top:1px dashed #000000;width:100%;height:50%;font-weight:500'>" & meeting_count & "<span style='font-size:11px'> פגישות</span></span>"
										if meeting_count < 1 then
											style = "style=""background-color:#F0F0F0; """ 			
										else
											style = "style=""background-color:#53FE01; """ 
										end if
								'end if
								'set rs_meetings = nothing	
								   
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
							</form>
							</tr>					
							</table>								
														

</body>
</html>

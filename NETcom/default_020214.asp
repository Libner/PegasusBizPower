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

</head>
<body>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">        
	<tr><td width=100% colspan=3>
	<!--#include file="logo_top.asp"-->
	</td></tr>
	<tr>				
			<td width=100% valign=top>
			<table cellpadding=0 cellspacing=0 width=100% border=0>
			<tr><td>
			<!--#include file="top.asp"-->
			</td></tr>			
			<tr><td height="42" bgcolor="#FFFFFF"></td></tr>			
			<tr><td align=right><!--include -->
			 <iframe src="<%=Application("VirDir")%>/netCom/info_new.aspx" 
          width="766" height="62" frameborder="0" ></iframe>
   
			<!--include end -->
			
			</td></tr>
			
   




</table></td></tr>
</table>


</body>
</html>

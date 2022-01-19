<!--#include file="../../connect.asp"-->
<%  

    	
	If Request.form("screenId") <> nil Then
		if isnumeric(trim(Request.form("screenId"))) then
			screenId = trim(Request.form("screenId"))
		end if
	end if
	If Request.form("tourId") <> nil Then
		if isnumeric(trim(Request.form("tourId"))) then
			tourId = trim(Request.form("tourId"))
		end if
	end if
	If Request.form("competitorId") <> nil Then
		if isnumeric(trim(Request.form("competitorId"))) then
			competitorId = trim(Request.form("competitorId"))
		end if
	end if
	If Request.form("startDate") <> nil Then
		if isdate(trim(Request.form("startDate"))) then
			startDate = trim(Request.form("startDate"))
		end if
	end if
	If Request.form("endDate") <> nil Then
		if isdate(trim(Request.form("endDate"))) then
			endDate = trim(Request.form("endDate"))
		end if
	end if
'response.Write(	"<br>selTour=" & Request.form("selTour"))
'response.Write(	"<br>selCompetitor=" & Request.form("selCompetitor"))
'response.Write(	"<br>Start_Date=" & Request.form("Start_Date"))
'response.Write(	"<br>End_Date=" & Request.form("End_Date"))


	
'response.Write(	"<br>tourId=" & tourId)
'response.Write(	"<br>competitoId=" & competitoId)
'response.Write(	"<br>startDate=" & startDate)
'response.Write(	"<br>endDate=" & endDate)

	
	If Request.querystring("screenId") <> nil Then
		if isnumeric(trim(Request.querystring("screenId"))) then
			screenId = trim(Request.querystring("screenId"))
		end if
	end if
	If Request.querystring("tourId") <> nil Then
		if isnumeric(trim(Request.querystring("tourId"))) then
			tourId = trim(Request.querystring("tourId"))
		end if
	end if
	If Request.querystring("competitorId") <> nil Then
		if isnumeric(trim(Request.querystring("competitorId"))) then
			competitorId = trim(Request.querystring("competitorId"))
		end if
	end if
	If Request.querystring("startDate") <> nil Then
		if isdate(trim(Request.querystring("startDate"))) then
			startDate = trim(Request.querystring("startDate"))
		end if
	end if
	If Request.querystring("endDate") <> nil Then
		if isdate(trim(Request.querystring("endDate"))) then
			endDate = trim(Request.querystring("endDate"))
		end if
	end if

	
'response.Write(	"<br>screenId=" & Request.querystring("screenId"))
'response.Write(	"<br>tourId=" & Request.querystring("tourId"))
'response.Write(	"<br>competitorId=" & Request.querystring("competitorId"))
'response.Write(	"<br>startDate=" & Request.querystring("startDate"))
'response.Write(	"<br>endDate=" & Request.querystring("endDate"))


	
'response.Write(	"<br>screenId=" & screenId)
'response.Write(	"<br>tourId=" & tourId)
'response.Write(	"<br>competitorId=" & competitorId)
'response.Write(	"<br>startDate=" & startDate)
'response.Write(	"<br>endDate=" & endDate)


	isValid=1
	
	if tourId<>"" and competitorId<>"" and startDate<>"" and endDate<>"" then 'only joined to tour and Competitor
	
		'check if exists screen for the same tour and the same Competitor in the same interval:  the new and old dates are overlaped
		sqlstr= "set dateformat dmy;select Screen_Id from Compare_Screens where Tour_Id=" & TourId & " and Competitor_Id=" & competitorId & " and "
		sqlstr= sqlstr & " not( datediff(dd,End_Date,'" & StartDate & "')>0 or datediff(dd,Start_Date,'" & EndDate & "')<0)"
		if screenID<>"" then
			sqlstr= sqlstr & " and Screen_Id<>" & screenId
		end if
		'response.write sqlstr & "<br>"
		'23/08/2017','23/07/2017'<0
		set rs_tmp = con.getRecordSet(sqlstr)
		if not rs_tmp.eof then
				isValid=0
		end if
		set rs_tmp = Nothing
	End If
	response.write isValid
%>
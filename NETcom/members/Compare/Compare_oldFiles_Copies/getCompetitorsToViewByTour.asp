<!--#include file="../../connect.asp"-->
<%
	If Request.form("tourId") <> nil Then
		if isnumeric(trim(Request.form("tourId"))) then
			if trim(Request.form("tourId"))<>"0" then
				tourId = trim(Request.form("tourId"))
			end if
		end if
	end if
	
	competitorId=0
	If Request.form("competitorId") <> nil Then
		if isnumeric(trim(Request.form("competitorId"))) then		
			if trim(Request.form("competitorId"))<>"0" then
				competitorId = cint(Request.form("competitorId"))
			end if
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

	If Request.querystring("tourId") <> nil Then
		if isnumeric(trim(Request.querystring("tourId"))) then
			tourId = trim(Request.querystring("tourId"))
		end if
	end if
	If Request.querystring("competitorId") <> nil Then
		if isnumeric(trim(Request.querystring("competitorId"))) then
			competitorId = cint(Request.querystring("competitorId"))
		end if
	end if
			dir_var = "ltr"  
		  align_var = "right"  
		    dir_obj_var = "rtl"
		%>    
		<option value="0"></option>
	<%if tourId<>"" then
															
														'get Competitor' names 
															sqlstr="SELECT Compare_Screens.Competitor_Id,Competitor_Name from Compare_Screens "
										sqlstr=sqlstr & " inner join Compare_Competitors on Compare_Competitors.Competitor_Id=Compare_Screens.Competitor_Id where Tour_Id=" & tourId 
										sqlstr=sqlstr & "  and datediff(dd,Compare_Screens.[Start_Date],getdate())>=0 and datediff(dd,Compare_Screens.[End_Date],getdate())<=0 "
										sqlstr=sqlstr & "  order by Competitor_Name,End_Date desc, Start_Date desc"
															
															'sqlstr=sqlstr & "  Where Tour_Id = " & Tour_Id
															set rssub = con.getRecordSet(sqlstr)
															do while not rssub.eof															
																selCompetitor_Id= cint(rssub("Competitor_Id"))																
																selCompetitor_Name = trim(rssub("Competitor_Name"))	
																%>
																<option value="<%=selCompetitor_Id%>"  <%if selCompetitor_Id =competitorId then%>selected<%end if%>><%=selCompetitor_Name%></option>
																<%
																rssub.movenext
															loop
															set rssub=Nothing%>
<%end if%>

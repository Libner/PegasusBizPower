<!--#include file="../../connect.asp"-->
<%  

    	
	If Request.form("competitorId") <> nil Then
		if isnumeric(trim(Request.form("competitorId"))) then
			competitorId = trim(Request.form("competitorId"))
		end if
	end if
	strValues=""
	
	if competitorId<>""  then 
		sqlstr= "select Competitor_SearchTerms FROM Compare_Competitors where Competitor_Id=" & CompetitorId 
		'23/08/2017','23/07/2017'<0
		set rs_tmp = con.getRecordSet(sqlstr)
		if not rs_tmp.eof then
			strValues=strValues & rs_tmp("Competitor_SearchTerms") 
		end if
		set rs_tmp = Nothing
	End If
	response.write strValues
%>
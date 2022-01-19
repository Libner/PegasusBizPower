<!--#include file="../../connect.asp"-->
<%  

    	
	If Request.form("competitorId") <> nil Then
		if isnumeric(trim(Request.form("competitorId"))) then
			competitorId = trim(Request.form("competitorId"))
		end if
	end if
	If Request.querystring("competitorId") <> nil Then
		if isnumeric(trim(Request.querystring("competitorId"))) then
			competitorId = trim(Request.querystring("competitorId"))
		end if
	end if
	strimgTaglogo=""
		
	if competitorId<>""  then 
		'23/08/2017','23/07/2017'<0
		set fso=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
		set pr=con.getRecordSet("SELECT Competitor_Logo FROM Compare_Competitors where Competitor_Id="& competitorId)
		if not pr.eof then
			fileName1=pr("Competitor_Logo")
			if fileName1<>"" then
				fileString= Server.MapPath("../../../download/competitors/"& fileName1 ) 'deleting of existing file
				if fso.FileExists(fileString) then
					strimgTaglogo="<img src='../../../download/competitors/"& fileName1& "' border=0>" 
				end if
			end if	
		end if
	End If
	response.write strimgTaglogo
%>
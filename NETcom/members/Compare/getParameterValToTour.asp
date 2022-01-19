<!--#include file="../../connect.asp"-->
<%  

    	
	If Request.form("tourId") <> nil Then
		if isnumeric(trim(Request.form("tourId"))) then
			tourId = trim(Request.form("tourId"))
		end if
	end if
	If Request.form("paramId") <> nil Then
		if isnumeric(trim(Request.form("paramId"))) then
			paramId = trim(Request.form("paramId"))
		end if
	end if
	strValues=""
	if tourId<>"" and paramId<>""  then 
		sqlstr= "select Parameter_Value,Parameter_Name_Variabled FROM Compare_Parameters_ToTour where Tour_Id=" & TourId & " and Parameter_Id=" & paramId 
		'23/08/2017','23/07/2017'<0
		set rs_tmp = con.getRecordSet(sqlstr)
		if not rs_tmp.eof then
				strValues=rs_tmp("Parameter_Value") & "~|~" & rs_tmp("Parameter_Name_Variabled")
		end if
		set rs_tmp = Nothing
	End If
	response.write strValues
%>
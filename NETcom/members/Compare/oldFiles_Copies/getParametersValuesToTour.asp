<!--#include file="../../connect.asp"-->
<%  

    	
	If Request.form("tourId") <> nil Then
		if isnumeric(trim(Request.form("tourId"))) then
			tourId = trim(Request.form("tourId"))
		end if
	end if
	strValues=""
	
	if tourId<>""  then 
		sqlstr= "select Parameter_Id,Parameter_Value,Parameter_Name_Variabled FROM Compare_Parameters_ToTour where Tour_Id=" & TourId & " order by Parameter_Id"
		'23/08/2017','23/07/2017'<0
		set rs_tmp = con.getRecordSet(sqlstr)
		do while not rs_tmp.eof
			if strValues<>"" then
				strValues=strValues & "#|#" 
			end if
				strValues=strValues & rs_tmp("Parameter_Id")  & "~|~" & rs_tmp("Parameter_Value") & "~|~" & rs_tmp("Parameter_Name_Variabled")
			rs_tmp.movenext
		loop
		set rs_tmp = Nothing
	End If
	response.write strValues
%>
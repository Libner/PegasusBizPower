<!--#include file="../../reverse.asp"-->
<!--#include file="../../connect.asp"-->
<%
    OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
    set xmldoc = Server.CreateObject("Microsoft.XMLDOM")
	xmldoc.async=false
	xmldoc.load(request)

	for each x in xmldoc.documentElement.childNodes
		if x.NodeName = "meeting_date" then meeting_date=x.text  
		if x.NodeName = "start_time"   then start_time=x.text  
		if x.NodeName = "end_time"     then end_time=x.text  
		if x.NodeName = "meeting_participants" then meeting_participants=x.text  
		if x.NodeName = "OrgID" then OrgID=x.text 	
		if x.NodeName = "meetingID" then meetingID=x.text 				
	next
	
	meetings_list = ""
	
	If trim(OrgID) <> "" Then    
		sqlstr = "SET DATEFORMAT DMY; Select DISTINCT meeting_id from meetings_view WHERE Organization_ID = " & OrgID &_
		" And participant_id IN (" & meeting_participants & ") And meeting_date = '" & meeting_date & "'"
		If trim(meetingID) <> "" Then
			sqlstr = sqlstr & " And Meeting_ID <> " & meetingID
		End If	
		set rs_meetings = con.getRecordSet(sqlstr)   
		if rs_meetings.eof Then
			meetings_list = ""
		else  
		while not rs_meetings.eof      
			meeting_id = rs_meetings(0)    
			
			sqlStr1 = "SET DATEFORMAT DMY; Select meeting_id from meetings where ORGANIZATION_ID = "& OrgID &_
			" And meeting_date = '" & meeting_date & "' And meeting_id = " & meeting_id &_									
			" And " &_
			" (( " &_			
			" datediff(minute, start_time, '"&end_time&"') > 0 And " &_
			" datediff(minute, end_time, '"&end_time&"') <= 0 And " &_
			" datediff(minute, start_time,'"&start_time&"') <= 0 And " &_
			" datediff(minute, end_time,'"&start_time&"') < 0 ) " &_
			" Or " &_
			" ( " &_
			" datediff(minute, start_time, '"&end_time&"') > 0 And " &_
			" datediff(minute, end_time, '"&end_time&"') >= 0 And " &_
			" datediff(minute, start_time,'"&start_time&"') >= 0 And " &_
			" datediff(minute, end_time,'"&start_time&"') < 0 )) "
 			
			set rs_check = con.getRecordSet(sqlStr1)
			If not rs_check.eof Then
			    meeting_id = rs_check(0)
				meetings_list = meetings_list & meeting_id & ","
			End If
			set rs_check = Nothing
			
			rs_meetings.moveNext
		wend	
		end if    
    Else
		meetings_list = ""
    End If
    
    set rs_meetings = Nothing
    Set con = Nothing
    set xmldoc = Nothing
    
    Response.Write meetings_list
    
%>
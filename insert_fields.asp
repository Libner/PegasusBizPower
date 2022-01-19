g<!--#include file="netcom/connect.asp" -->
<!--#include file="netcom/reverse.asp" -->
<%   count = 0
     sqlstr  = "SELECT APPEAL_ID FROM APPEALS WHERE (QUESTIONS_ID = 870) AND "&_
     " (APPEAL_ID NOT IN (SELECT Appeal_ID FROM Form_Value WHERE FIELD_ID = 6244))"
     set rs_appeals = con.getRecordSet(sqlstr)
     while not rs_appeals.eof 
		appID = rs_appeals(0)
		sqlstr = "Insert Into Form_Value (APPEAL_ID,FIELD_ID,FIELD_VALUE) Values (" & appID & ", 6244, NULL)"	
		con.ExecuteQuery(sqlstr)
		count = count + 1
		rs_appeals.moveNext
     wend
     set rs_appeals = Nothing
     set con = Nothing 
     
     Response.CharSet = "windows-1255"    
     Response.Write " נוספו " & count & " שורות "
%>

<!--#include file="../NETcom/connect.asp"-->
<!--#include file="../NETcom/reverse.asp"-->
<%
	sqlstr =  "select Appeal_Id,dbo.get_DateDiff(RelationId,Appeal_date) as NDays from appeals where RelationId is not null and RelationId>0"

		'Response.Write sqlstr
		'Response.End
		set rsq = con.getRecordSet(sqlstr)
		'If not rsq.eof Then	
		while not rsq.eof 
		appeal_id=rsq("APPEAL_ID")
		
		NDays=rsq("NDays")
		    sqlstr="UPDATE APPEALS SET NumberDaysBetweenForms = '" & NDays &  "' WHERE (appeal_id = " & appeal_id & ")"
		'Response.Write(sqlstr &"<BR>")
		'Response.End 
		con.ExecuteQuery(sqlstr)
		rsq.moveNext
		wend
		'end if
response.Write "done"
response.end

%>



  SET DATEFORMAT DMY

SELECT    *
FROM         APPEALS APP 
WHERE     (app.QUESTIONS_ID = 16504) and appeal_status=3
and 
exists
(select * from APPEALS app2
where APP.APPEAL_ID=app2.RelationId )

select * FROM         APPEALS 
where relationId=243276

and app.QUESTIONS_ID = 16735


select * from appeals where relationid>0
select * from appeals where appeal_id=1164605
select * from appeals where relationid=1164605

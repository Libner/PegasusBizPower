<!--#include file="../NETcom/connect.asp"-->
<!--#include file="../NETcom/reverse.asp"-->
<%
	''sqlstr =  "SELECT     APPEAL_ID, FIELD_ID, FIELD_VALUE, dbo.get_Country_CRM(FIELD_VALUE) AS Country_Id FROM FORM_VALUE WHERE (FIELD_ID = 40107) AND (dbo.get_Country_CRM(FIELD_VALUE) IS NOT NULL)	and APPEAL_ID>1616407"
'sqlstr="SELECT     FORM_VALUE.APPEAL_ID, FORM_VALUE.FIELD_ID, FORM_VALUE.FIELD_VALUE, APPEALS.APPEAL_DATE, APPEALS.appeal_CountryId " & _
'" FROM         FORM_VALUE LEFT OUTER JOIN APPEALS ON APPEALS.APPEAL_ID = FORM_VALUE.APPEAL_ID  " & _
'" WHERE     (FORM_VALUE.FIELD_ID = 40107) AND (FORM_VALUE.FIELD_VALUE LIKE '%עם שייט תענוגות למקסיקו%') order by field_value desc"


sqlstr =  "SELECT     FORM_VALUE.APPEAL_ID, FORM_VALUE.FIELD_ID, FORM_VALUE.FIELD_VALUE, APPEALS.APPEAL_DATE, APPEALS.appeal_CountryId " & _
" FROM         FORM_VALUE LEFT OUTER JOIN " & _
" APPEALS ON APPEALS.APPEAL_ID = FORM_VALUE.APPEAL_ID " & _
" WHERE     (FORM_VALUE.FIELD_ID = 40107) AND (FORM_VALUE.FIELD_VALUE LIKE '%ארה%')  and  (appeal_CountryId IS NULL) order by field_value desc"

		'Response.Write sqlstr
		'Response.End
		set rsq = con.getRecordSet(sqlstr)
		'If not rsq.eof Then	
		while not rsq.eof 
		appeal_id=rsq("APPEAL_ID")
		'country_id=rsq("Country_Id")
		'country_id=26
		country_id=54
		    sqlstr="UPDATE APPEALS SET appeal_CountryId = " & country_id &  " WHERE (appeal_id = " & appeal_id & ")"
		'Response.Write(sqlstr &"<BR>")
		'Response.End 
		con.ExecuteQuery(sqlstr)
		rsq.moveNext
		wend
		'end if
response.Write "done"
response.end

%>

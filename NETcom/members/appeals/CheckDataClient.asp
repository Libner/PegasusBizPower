	
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
	<%
	'if false then
	'recCount=10
	trapp=request.QueryString("appIds")
	 sqlContact="select CONTACT_ID,Company_id from Appeals where CONTACT_ID is not null and   Appeal_id in (" & trapp & ")"
     set  c = con.getRecordSet(sqlContact)
	do while not c.eof 
	if ContactId="" then
	 ContactId = c("CONTACT_ID")
	 else
	 ContactId =ContactId &"," & c("CONTACT_ID")
	 end if
	 
	 c.Movenext
	 Loop
	set c=Nothing
	if ContactId<>"" then
	 sqlstr = "SELECT count(*) as CountRecords  FROM dbo.tasks where Contact_Id in ("& ContactId &") and (task_status=1 or task_status=2)" 
	 set  tasksCount = con.getRecordSet(sqlstr)
	 If not tasksCount.EOF then		
        recCount = tasksCount("CountRecords")	
      else
       recCount=0
     End if
      else
       recCount=0
     End if
    ' end if
    ' recCount=15
     response.Write recCount
     %>
	 
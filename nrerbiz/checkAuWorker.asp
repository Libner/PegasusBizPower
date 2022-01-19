<%
Dim workerName

check_worker 'Call subroution - checking of user permissions

Sub check_worker

username=session("admin.username")
password=session("admin.password")
if isnull(username) or isnull(password) or trim(username)="" or trim(password)="" then
	isworker=false
else


	sqlstring="SELECT * FROM workers WHERE loginName='" & username & "' AND Password='" & password  &"'"
		
	set worker=con.Execute(sqlstring)
	if worker.EOF  then
		isworker=false
	else
		workerLogin=trim(worker("LoginName"))
		workerPassword=trim(worker("Password"))
		workerName=worker("firstName")
		session("activateType") = 0
		if workerPassword<>password or workerLogin<>username then
			isworker=false
		else
			session("admin.username")=username
			session("admin.password")=password
			isworker=true
		end if
	end if
end if
if not isworker then
	Response.redirect "../default.asp"
end if
end sub

%>


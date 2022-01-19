<%
	Session.Timeout = 999	
	Dim workerName
	if Request.Form("username")<>nil then
		username=Request.Form("username")
		password=Request.Form("password")
	elseif session("admin.username") <> nil then	
		username=session("admin.username")
		password=session("admin.password")
	else 	
		Response.Redirect(strLocal & "/nrerbiz/default.asp")
	end if

	sqlstring="SELECT LoginName,Password,firstName FROM ADMINISTARTORS WHERE loginName='" & sfix(username) &_
	"' AND Password='" & sfix(password) &"'"
	'Response.Write sqlstring
	'Response.End
	set worker=con.GetRecordSet(sqlstring)
	if worker.EOF then
		 Response.Redirect(strLocal & "/nrerbiz/default.asp")
	else
    	 workerLogin=trim(worker("LoginName"))
		 workerPassword=trim(worker("Password"))
		 workerName=worker("firstName")
		 if workerPassword<>password or workerLogin<>username then
		 	isworker=false
		 else
		 	session("admin.username")=username
		 	session("admin.password")=password
		 	isworker=true
		 end if
	end if

%>


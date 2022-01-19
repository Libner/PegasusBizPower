<!--#INCLUDE FILE="../Netcom/reverse.asp"-->
<!--#include file="../Netcom/connect.asp"-->
<%phone = trim(Request.QueryString("phone"))
	phone = Replace(phone, " ", "")
	phone = Replace(phone, "_", "")
	phone = Replace(phone, "-", "")
	phone = Replace(phone, ".", "")
	phone = Replace(phone, Chr(39), "")
	phone = Replace(phone, Chr(34), "")	
    
	workerId = trim(Request.QueryString("workerId"))	 
	If Not IsNumeric(workerId) Then
		workerId = 0
	End If
	
	workerName = trim(Request.QueryString("workerName"))	 
	If IsNULL(workerName) Then
		workerName = ""
	End If	 
	 
	 ContactId = 0
	 
	If Len(phone)>0 Then
		sqlstr = "SELECT TOP 1 C.CONTACT_ID FROM dbo.CONTACTS C WHERE (C.ORGANIZATION_ID  = 264) " & _ 
		" AND ((C.phone = '" & sFix(phone) & "') " & _
		" OR (C.cellular = '" & sFix(phone) & "'))  "			
	    set rs = con.GetRecordSet(sqlstr) 
		If Not rs.EOF then		
			ContactId = cLng(rs("CONTACT_ID"))	
		End If
		Set rs = Nothing	
		if 	ContactId=0 then
			sqlstr = "SELECT TOP 1 C.CONTACT_ID FROM dbo.CONTACTS C WHERE (C.ORGANIZATION_ID  = 264) " & _ 
		" AND ((C.phone_additional = '" & sFix(phone) & "'))  "			
	    set rs1 = con.GetRecordSet(sqlstr) 
		If Not rs1.EOF then		
			ContactId = cLng(rs1("CONTACT_ID"))	
		End If
		Set rs1 = Nothing	
		
		end if
	End If	
	
	If Len(workerName)>0 Or workerId > 0 Then'
		sqlstr = "SELECT TOP 1 U1.LOGINNAME, U1.PASSWORD " & _
		" FROM dbo.Users U1 WHERE (U1.ORGANIZATION_ID  = 264) "
		If workerId > 0 Then
			sqlstr = sqlstr & " AND (U1.USER_ID = " & workerId & ")"
		Else
			sqlstr = sqlstr & " AND (U1.LOGINNAME = '" & sFix(workerName) & "')"
		End If			
	    set rs = con.GetRecordSet(sqlstr) 
		If Not rs.EOF then		
			LOGINNAME = trim(rs("LOGINNAME"))
			PASSWORD = trim(rs("PASSWORD"))
		End If
		Set rs = Nothing		
	End If		 		 
	 
	 Set con = Nothing 	
	 'If ContactId > 0 Then %>
 <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
 <html>
<head>
	<title>Enter From Web</title>
	<meta charset="windows-1255">
	<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
</head>
<body>
	<form method="post" action="../netcom/default.asp" target="_self" ID="Form1" name="Form1">
	<table border="0" cellpadding="0" cellspacing="0" width="100%" style="display: none;" >
    <tr>
      <td><input type="text" name="username" id="username" value="<%=vFix(LOGINNAME)%>"></td>
    </tr>
    <tr>
      <td><input type="text" name="password" id="password" value="<%=vFix(PASSWORD)%>"></td>
    </tr>
    <tr>
      <td><input type="text" name="OpenContactId" id="OpenContactId" value="<%=vFix(ContactId)%>"></td>
    </tr>    
    <tr>
      <td><input type="submit" value="enter" name="bntSubmit" id="bntSubmit"></td>
    </tr>
  </table>
  </form>
<script language="javascript" type="text/javascript">
<!--
		window.document.getElementById("bntSubmit").click();
//-->
</script>
</body>
</html>
<%'End If%>
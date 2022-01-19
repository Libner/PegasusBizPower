<%@ Language=VBScript%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%prodId = trim(Request.QueryString("prodId"))
     pageID = trim(Request.QueryString("pageID"))
   	
    if Request("Page")<>"" then
		Page=Request("Page")
    else
		Page=1
    end if	
	
	if trim(Request.QueryString("numOfRow"))<>"" then
		numOfRow=Request.QueryString("numOfRow")
	else
		numOfRow = 1
	end if  
	
	found_product = false
	If isNumeric(prodId) Then
		sqlStr = "SELECT EMAIL_SUBJECT, Is_Archive FROM PRODUCTS WHERE product_id=" & prodId & " AND ORGANIZATION_ID=" & OrgID	
	    set rs_product = con.GetRecordSet(sqlStr)	
	    If not rs_product.eof Then	
			product_name = trim(rs_product("EMAIL_SUBJECT"))
			IsArchive = trim(rs_product("Is_Archive"))
			found_product = true
		End If
		set rs_product=nothing
	End if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body style="margin:0px;background-color:#e5e5e5" onload="window.focus();">
<%If found_product Then%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<tr><td class="page_title" align="center" dir=rtl>
<span style="width:80%;text-align:center">&nbsp;רשימת נמענים שלא פתחו את המייל&nbsp;<font color="#6E6DA6"><%=product_name%></font>&nbsp;</span>
<a class="button_edit_1" style="width:115;" href='#' ONCLICK="window.open('excelDownloadN.asp?prodId=<%=prodId%>',null,'alwaysRaised,left=180,top=200,height=200,width=450,status=no,toolbar=no,menubar=no,location=no');">הורדה ל-Excel</a>
</td></tr>
<tr>
<td align=center width=100% valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0>
<tr><td width=100% valign=top>
	<table cellspacing="0" cellpadding="0" align=center border="0" width=100%>
	<tr><td align=right>
		<table border=0 width=100% cellpadding=1 cellspacing=1>				
		<tr>	
			<td width=65 nowrap align=center class="card3">כרטיס לקוח</td>																
			<td width=65 nowrap align=center class="card3">תאריך הפצה</td>							
			<td width=100% align=left class="card3">&nbsp;Email&nbsp;</td>
			<td width=170 nowrap align=right class="card3">&nbsp;חברה&nbsp;</td>			
			<td width=140 nowrap align=right class="card3">&nbsp;שם נמען&nbsp;</td>
		</tr>				
	<%	sqlStr = "SELECT PEOPLE_ID,PEOPLE_NAME,PEOPLE_EMAIL,PEOPLE_COMPANY,DATE_SEND,CONTACT_ID,IsEmailValid "
			If IsArchive = "0" Then
				sqlStr = sqlStr & " From PRODUCT_CLIENT "
			Else
				sqlStr = sqlStr & " From PRODUCT_CLIENT_ARCH "
			End If
			sqlStr = sqlStr & " WHERE (ORGANIZATION_ID=" & OrgID & ") AND (PRODUCT_ID = " & prodId & ") " & _
			" AND (IsNULL(IS_OPENED, 0) = 0) Order by PEOPLE_EMAIL"				
			''Response.Write sqlStr
			set rs_clients = con.GetRecordSet(sqlStr)
			if not rs_clients.eof then
					rs_clients.PageSize = 25
					rs_clients.AbsolutePage=Page					
					recCount=rs_clients.RecordCount 
					NumberOfPages=rs_clients.PageCount
					i=1
					j=0 
					do while (not rs_clients.eof and j < rs_clients.PageSize)
						PEOPLE_ID=rs_clients.Fields(0) 
						PEOPLE_NAME = rs_clients.Fields(1)	
						EMAIL = rs_clients.Fields(2)
						COMPANY_NAME = rs_clients.Fields(3)									
						DATE_SEND = rs_clients.Fields(4)						
						CONTACT_ID = trim(rs_clients.Fields(5))
						IsEmailValid = trim(rs_clients.Fields(6))
						If IsDate(DATE_SEND) Then
							DATE_SEND = Day(DATE_SEND) & "/" & Month(DATE_SEND) & "/" & Right(Year(DATE_SEND),2)
						ElseIf trim(IsEmailValid) = "0" Then
							DATE_SEND = "מייל לא תקין"
						End If	%>
				<tr>
					<td class="card" align=center><%If Len(CONTACT_ID) > 0 Then%><a href="../companies/contact.asp?contactId=<%=CONTACT_ID%>" target="_blank"><img src="../../images/edit_icon.gif" border="0"></a><%End If%></td>
					<td align=center class="card"><%=DATE_SEND%></td>			
					<td class="card" align=left>&nbsp;<%=EMAIL%>&nbsp;</td>											
					<td class="card" align=right dir="rtl">&nbsp;<%=COMPANY_NAME%>&nbsp;</td>					
					<td class="card" align=right dir="rtl">&nbsp;<%=PEOPLE_NAME%>&nbsp;</td>
				</tr>					
				<%
					rs_clients.movenext
					j=j+1
					loop
				%>
				<%
					url = "not_opened_list.asp?prodId=" & prodId
				%>
	  <%if NumberOfPages > 1 then%>
	  <tr>							
	 	 <td align=center colspan="6" class="card">
		 <table border="0" cellspacing="0" cellpadding="2">
		 <% If NumberOfPages > 10 Then 
				num = 10 : numOfRows = cInt(NumberOfPages / 10)
			Else num = NumberOfPages : numOfRows = 1    	                      
			End If
		%>
	    <tr>
		<%if numOfRow <> 1 then%> 
		<td valign="middle" align="right">
		<A class=pageCounter title="לדפים הקודמים" href="<%=url%>&page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
		<%end if%>
		<td><font size="2" color="#001c5e">[</font></td>    
		 <%for i=1 to num
	        If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	        if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		        <td bgcolor="#e6e6e6" align="right"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	        <%else%>
	            <td bgcolor="#e6e6e6" align="right"><A class=pageCounter href="<%=url%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	        <%end if
	        end if
	    next%>
		<td bgcolor="#e6e6e6" align="right"><font size="2" color="#001c5e">]</font></td>
		<%if NumberOfPages > cint(num * numOfRow) then%>  
		<td bgcolor="#e6e6e6" align="right"><A class=pageCounter title="לדפים הבאים" href="<%=url%>&page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
		<%end if%>		               
		</tr></table></td>
	    <%end if 'recCount > res.PageSize%>				
		<% set rs_clients = nothing %>
		<tr><td colspan=11 class="card" align=center><font style="color:#6E6DA6;font-weight:600">סה"כ <%=recCount%> נמענים שלא פתחו את המייל</font></td></tr>
		<%end if%>
		</table></td></tr>											
		</table>
		<%end if%>				
		</td>		
	</tr>
	</table></td></tr>
</table>
</body>
</html>
<%set con=nothing%>
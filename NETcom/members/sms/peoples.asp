<!--#include file="../../connect.asp"-->
<%ScriptTimeout=6000%>
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%smsId = Request.QueryString("smsId")   	
   if Request("Page")<>"" And isNumeric(Request("Page")) then
		Page = cInt(Request("Page"))
   Else
		Page=1
   End if	
	
	if trim(Request.QueryString("numOfRow"))<>"" then
		numOfRow=Request.QueryString("numOfRow")
	else
		numOfRow = 1
	end if  
	
	found_product = false
	If isNumeric(smsId) Then
		sqlStr = "SELECT sms_desc FROM sms_sends WHERE (sms_id=" & smsId & ") AND (ORGANIZATION_ID=" & OrgID & ")"
	    set rs_product = con.GetRecordSet(sqlStr)	
	    If not rs_product.eof Then	
			product_name = rs_product("sms_desc")		
			found_product = true
		Else
			found_product = false
		End If
		set rs_product=nothing
	Else
		found_product = false	
	End if%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body style="margin:0px;background-color:#e5e5e5" onload="window.focus();">
<%If found_product Then%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<tr><td class="page_title" align="center" dir=rtl>
<span style="width:80%;text-align:center">&nbsp;רשימת נמענים בהפצה&nbsp;<font color="#6E6DA6"><%=product_name%></font>&nbsp;</span>
<a class="button_edit_1" style="width:115;" href='#' ONCLICK="window.open('excelDownloadS.asp?smsId=<%=smsId%>',null,'alwaysRaised,left=180,top=200,height=200,width=450,status=no,toolbar=no,menubar=no,location=no');">הורדה ל-Excel</a>
</td></tr>
<tr>
<td align=center width=100% valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0>
<tr><td width="100%" valign="top">
	<table cellspacing="0" cellpadding="0" align="center" border="0" width="100%">
	<tr><td align="right">
		<table border="0" width="100%" cellpadding="1" cellspacing="1">				
		<tr>
			<td width="65" nowrap align="center" class="card3">תאריך שליחה</td>							
			<td width="100%" align="left" class="card3">&nbsp;טלפון נייד&nbsp;</td>
			<td width="170" nowrap align="right" class="card3">&nbsp;חברה&nbsp;</td>			
			<td width="140" nowrap align="right" class="card3">&nbsp;שם נמען&nbsp;</td>
		</tr>				
<%		sqlStr = "SELECT PEOPLE_ID, PEOPLE_NAME, PEOPLE_CELL, PEOPLE_COMPANY, DATE_SEND " &_
			" FROM sms_to_people WHERE (ORGANIZATION_ID=" & OrgID & ") AND (sms_id = " & smsId & ") " &_
			" AND DATE_SEND Is Not NULL Order by PEOPLE_CELL"				
			''Response.Write sqlStr
			set rs_tmp = con.GetRecordSet(sqlStr)
			recCount = 0
			if not rs_tmp.eof then
					rs_tmp.PageSize = 25
					rs_tmp.AbsolutePage=Page					
					recCount=rs_tmp.RecordCount 
					NumberOfPages=rs_tmp.PageCount
					i=1
					j=0 
					do while (not rs_tmp.eof and j < rs_tmp.PageSize)
					PEOPLE_ID=rs_tmp("PEOPLE_ID") 
					PeopleCELL = rs_tmp("PEOPLE_CELL")
					COMPANY_NAME = rs_tmp("PEOPLE_COMPANY")
					PEOPLE_NAME = rs_tmp("PEOPLE_NAME")				
					DATE_SEND = rs_tmp("DATE_SEND")
					If IsDate(DATE_SEND) Then
						DATE_SEND = Day(DATE_SEND) & "/" & Month(DATE_SEND) & "/" & Right(Year(DATE_SEND),2)
					End If%>
				<tr>	
					<td align=center class="card"><%=DATE_SEND%></td>			
					<td class="card" align=left>&nbsp;<%=PeopleCELL%>&nbsp;</td>											
					<td class="card" align=right>&nbsp;<%=COMPANY_NAME%>&nbsp;</td>					
					<td class="card" align=right>&nbsp;<%=PEOPLE_NAME%>&nbsp;</td>
				</tr>					
				<%
					rs_tmp.movenext
					j=j+1
					loop
				%>
				<%
					url = "peoples.asp?smsId=" & smsId
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
		<%	
		  	 set rs_tmp = nothing				
			 end if				
		%>
		<tr><td colspan=11 class="card" align=center><font style="color:#6E6DA6;font-weight:600">סה"כ <%=recCount%> נמענים בהפצה</font></td></tr>
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
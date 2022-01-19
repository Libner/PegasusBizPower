<!--#include file="../../netcom/connect.asp"-->
<!--#include file="../../netcom/reverse.asp"-->
<%
	OrgID = trim(Request.QueryString("OrgID"))	
	if trim(OrgID) <> "" then
		sqlStr = "select ORGANIZATION_NAME from  ORGANIZATIONS where ORGANIZATION_ID = " & OrgID
		set rs_ORGANIZATIONS = con.GetRecordSet(sqlStr)
		if not rs_ORGANIZATIONS.eof then
			if trim(rs_ORGANIZATIONS("ORGANIZATION_NAME")) <> "" then
				org_name = rs_ORGANIZATIONS("ORGANIZATION_NAME")
			end if
		end if
		set rs_ORGANIZATIONS = nothing
	end if		
	
	if trim(Request.QueryString("delMail")) <> "" then
	    delMail = trim(Request.QueryString("delMail"))	
		slqStr = "delete from blocked_emails where blocked_id="&delMail&" and ORGANIZATION_ID = "&OrgID
		con.ExecuteQuery(slqStr)		
	end if	
	
	if trim(Request.QueryString("dellAll")) = "1" then	   
		slqStr = "delete from blocked_emails where ORGANIZATION_ID = "&OrgID
		con.ExecuteQuery(slqStr)		
	end if	

	if trim(Request.QueryString("page"))<>"" then
		page=Request.QueryString("page")
	else
		Page=1
	end if  
	 
	if trim(Request.QueryString("numOfRow"))<>"" then
		numOfRow=Request.QueryString("numOfRow")
	else
		numOfRow = 1
	end if  

	dim sortby(16)	
	sortby(1) = " blocked_email, CAST(date_add as float) DESC"
	sortby(2) = " blocked_email DESC, CAST(date_add as float) DESC"
	sortby(3) = " count_people,CAST(date_add as float) DESC"
	sortby(4) = " count_people DESC, CAST(date_add as float) DESC"
	sortby(5) = " CAST(date_add as float)"
	sortby(6) = " CAST(date_add as float) DESC"
	sort = Request("sort")	
	if sort = nil then
		sort = 6
	end if				  
%>
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
function checkDelete() 
{   
	return window.confirm("? האם ברצונך למחוק את המייל מרשימת המיילים החסומים");	
}

function checkDeleteAll()
{ 
	return window.confirm("? האם ברצונך למחוק כל המיילים מרשימת המיילים החסומים");
}
//-->
</SCRIPT>

</head>
<body style="margin:0px" onload="window.focus();">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top"  bgcolor="#c0c0c0">
<tr>
<td width="100%" class="page_title" colspan=2>
<a class="button_edit_1" style="width:115;" dir=rtl href="excelDownloadBlocked.asp?OrgID=<%=OrgID%>">הורדה ל-Excel</a>
<span style="width:80%;text-align:center">&nbsp;רשימת מיילים חסומים&nbsp;-&nbsp;<font color="#000000"><%=org_name%></font></span>
</td></tr>         
<tr><td width="100%" valign="top">
<table width="100%" align="center" bgcolor="FFFFFF" border="0" cellpadding="1" cellspacing="1">
<tr bgcolor="#B4B4B4">   
<%
    if Request("Page")<>"" then
		Page=Request("Page")
	else
		Page=1
	end if	
	urlSort = "blocked_mails.asp?OrgID=" & OrgID
	sqlStr = "select blocked_id, blocked_email, date_add "&_
	" from blocked_emails where ORGANIZATION_ID=" & OrgID & " ORDER BY " & sortby(sort) 
	set rs_blocked_emails = con.GetRecordSet(sqlStr)														
%>
   	<td width="80" nowrap align="center" class="11normalB">&nbsp;מחיקה&nbsp;</td>	
	<td width=130 align=center nowrap class="11normalB"><a class="adm" HREF="<%=urlSort%>&sort=<%if sort=6 then%>5<%elseif sort=5 then%>6<%else%>6<%end if%>" target="_self">תאריך הוספה<img src="../../netcom/images/arrow_<%if trim(sort)="5" then%>down<%elseif trim(sort)="6" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>	
	<td width=100% align="right" class="11normalB"><a class="adm" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self">מייל<img src="../../netcom/images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>
</tr>					
<%
	If not rs_blocked_emails.eof Then
		If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
			PageSize = RowsInList
		Else	
     		PageSize = 25
		End If	   
	    rs_blocked_emails.PageSize = PageSize
		rs_blocked_emails.AbsolutePage=Page
		recCount=rs_blocked_emails.RecordCount 	
		Mail_count=rs_blocked_emails.RecordCount	
		NumberOfPages = rs_blocked_emails.PageCount
		i=1
		j=0
		do while (not rs_blocked_emails.eof and j<rs_blocked_emails.PageSize)
		    blocked_id = trim(rs_blocked_emails("blocked_id"))
			blocked_email = trim(rs_blocked_emails("blocked_email"))
			date_add = trim(rs_blocked_emails("date_add"))
			If IsDate(date_add) And trim(date_add) <> "" Then
				date_add = DateValue(Day(date_add) & "/" & Month(date_add) & "/" & Year(date_add))
			End if				
%>		
	<tr bgcolor="#DDDDDD">
	    <td class="11normalB" align=center><a href="blocked_mails.asp?delMail=<%=blocked_id%>&OrgID=<%=OrgID%>" ONCLICK="return checkDelete()"><IMG SRC="../../netcom/images/delete_icon.gif" BORDER=0></a></td>
		<td class="11normalB" align=center><%=date_add%></td>
		<td class="11normalB" align="right" dir="ltr">&nbsp;<%=blocked_email%>&nbsp;</td>
	</tr>					
<%
		rs_blocked_emails.movenext
		j=j+1
		loop
		set rs_blocked_emails = nothing
%>
<% if NumberOfPages > 1 then%>
<tr class="card"  bgcolor="#DDDDDD">
  <td width="100%" align=center nowrap class="card" colspan=11>
	<table border="0" cellspacing="0" cellpadding="2">               
	<% If NumberOfPages > 10 Then 
	        num = 10 : numOfRows = cInt(NumberOfPages / 10)
	    else num = NumberOfPages : numOfRows = 1    	                      
	    End If
	    If Request.QueryString("numOfRow") <> nil Then
	        numOfRow = Request.QueryString("numOfRow")
	    Else numOfRow = 1
	    End If
	%>
	<tr>
	<%If numOfRow <> 1 Then%> 
		<td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" title="לדפים הקודמים"><b><<</b></a></td>			                
		<%end if%>
	        <td><font size="2" color="#060165">[</font></td>
	        <%for i=1 to num
	            If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	            if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		            <td align="center"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	            <%else%>
	                <td align="center"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=i+10*(numOfRow-1)%>&numOfRow=<%=numOfRow%>"><%=i+10*(numOfRow-1)%></a></td>
	            <%end if
	            end if
	        next%>	    
			<td><font size="2" color="#060165">]</font></td>
		<%if NumberOfPages > cint(num * numOfRow) then%>  
			<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" title="לדפים הבאים"><b>>></b></a></td>
		<%end if%>   
		</tr> 				  	             
	</table></td></tr>				
	<%End If%>		
	<tr><td align=center height=20px class="normalB" bgcolor="#C7C7C7" colspan=11>נמצאו&nbsp;<%=Mail_count%>&nbsp;מיילים חסומים</td></tr>								
</table></td>
<%		
Else
%>
  <tr><td  bgcolor="#DDDDDD" align=center colspan=5>לא נמצאו מיילים חסומים</td></tr>
<%
End If
%>					
	</table></td>
<%'//end of blocked_emails%>			
</td>	
</tr>
</table>
</body>
<%set con=nothing%>
</html>

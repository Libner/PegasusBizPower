<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%
	OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
	org_name = trim(Request.Cookies("bizpegasus")("ORGNAME"))
	lang_id	 = trim(Request.Cookies("bizpegasus")("LANGID"))
	If isNumeric(lang_id) = false Or IsNull(lang_id) Or trim(lang_id) = "" Then
		lang_id = 1 
	End If
	If lang_id = 2 Then
		dir_var = "rtl"
		align_var = "left"
		dir_obj_var = "ltr"
	Else
		dir_var = "ltr"
		align_var = "right"
		dir_obj_var = "rtl"
	End If	
	
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

	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 72 Order By word_id"				
	set rstitle = con.getRecordSet(sqlstr)		
	If not rstitle.eof Then
	dim arrTitles()
	arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
	arrTitlesD = Split(trim(arrTitlesD), ";")		
	redim arrTitles(Ubound(arrTitlesD))
	For i=1 To Ubound(arrTitlesD)			
		arr_title = Split(arrTitlesD(i),",")			
		If IsArray(arr_title) And Ubound(arr_title) = 1 Then
		arrTitles(arr_title(0)) = arr_title(1)
		End If
	Next
	End If
	set rstitle = Nothing
	  
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
function FailureDelete(blocked_id) {
   <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את המייל"
     Else
		str_confirm = "Are you sure want to delete the Mail ?"
     End If   
   %>		
	if (confirm("<%=str_confirm%>"))		
	{
		window.document.location.href = "blocked_mails.asp?delMail=" + blocked_id;
		return false;
	}
	return false;	
}

function checkDeleteAll()
{
   <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק כל המיילים מרשימת המיילים החסומים"
     Else
		str_confirm = "Are you sure want to delete the all emails from the blocked emails list ?"
     End If   
   %>	
	return window.confirm("<%=str_confirm%>");
}
//-->
</SCRIPT>

</head>
<body style="margin:0px" onload="window.focus();">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF dir="<%=dir_var%>">
<tr>
<td width="100%" class="page_title" dir="<%=dir_var%>" colspan=2>
<span style="width:80%;text-align:center">&nbsp;<%=arrTitles(11)%>&nbsp;<!--מיילים&nbsp;חסומים--><%=arrTitles(7)%>&nbsp;-&nbsp;<font color="#595698"><%=org_name%></font></span>
<a class="button_edit_1" style="width:115;" href="excelDownloadBlocked.asp"><!--הורדה ל-Excel--><%=arrTitles(12)%></a>
</td></tr>         
<tr><td width="100%" valign="top">
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_obj_var%>">
<tr><td width="100%" valign="top" align="<%=align_var%>">    
<%'//start of blocked_emails%>	
<table border=0 width=100% cellpadding=0 cellspacing=1>
<%
    if Request("Page")<>"" then
		Page=Request("Page")
	else
		Page=1
	end if	
	urlSort = "blocked_mails.asp?1=1"
	sqlStr = "select blocked_id, blocked_email, date_add "&_
	" from blocked_emails where ORGANIZATION_ID=" & OrgID & " ORDER BY " & sortby(sort) 
	set rs_blocked_emails = con.GetRecordSet(sqlStr)														
%>
<tr>
	<td width=100 align=center nowrap class="title_sort<%if sort=5 or sort=6 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=6 then%>5<%elseif sort=5 then%>6<%else%>6<%end if%>" target="_self"><!--תאריך הוספה--><%=arrTitles(3)%><img src="../../images/arrow_<%if trim(sort)="5" then%>down<%elseif trim(sort)="6" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>	
	<td width=100% align="<%=align_var%>" dir="<%=dir_var%>" class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self"><!--מייל--><%=arrTitles(4)%><img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>
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
	<tr>
		<td class="card" align=center><%=date_add%></td>
		<td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=blocked_email%>&nbsp;</td>
	</tr>					
<%
		rs_blocked_emails.movenext
		j=j+1
		loop
		set rs_blocked_emails = nothing
%>
<% if NumberOfPages > 1 then%>
<tr class="card">
  <td width="100%" align=center nowrap class="card" colspan=11>
	<table border="0" cellspacing="0" cellpadding="2" ID="Table4">               
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
	<tr><td align=center height=20px class="card"  colspan=11><font style="color:#6E6DA6;font-weight:600"><!--נמצאו--><%=arrTitles(5)%>&nbsp;<%=Mail_count%>&nbsp;<!--מיילים חסומים--><%=arrTitles(7)%></font></td></tr>								
</table></td>
<%		
Else
%>
  <tr><td class="title_sort1" align=center colspan=5><!--לא נמצאו מיילים חסומים--><%=arrTitles(6)%></td></tr>
<%
End If
%>					
	</table></td>
<%'//end of blocked_emails%>			
</td>	
</tr>
<tr><td height=10 nowrap></td></tr>
</table>
</body>
<%set con=nothing%>
</html>

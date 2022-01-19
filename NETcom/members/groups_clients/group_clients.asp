<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%groupId = trim(Request("groupId"))
  pageId = trim(Request("pageId"))
 
  If trim(groupId) <> "" Then
	sqlstr="Select GROUPNAME FROM GROUPS WHERE GROUPID="&groupId
	set rsn = con.getRecordSet(sqlstr)
	If not rsn.eof Then
		groupname = trim(rsn(0))
	End if
	set rsn = Nothing	
  End If    

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
  
    sort = Request.QueryString("sort")	
    if trim(sort)="" then  sort = 0  end if 
  
    dim sortby(10)	
	sortby(0) = "PEOPLE_ID DESC"
	sortby(1) = "rtrim(ltrim(PEOPLE_EMAIL))"
	sortby(2) = "rtrim(ltrim(PEOPLE_EMAIL)) DESC"
	sortby(3) = "PEOPLE_NAME"
	sortby(4) = "PEOPLE_NAME DESC"
	sortby(5) = "PEOPLE_COMPANY"
	sortby(6) = "PEOPLE_COMPANY DESC"
	sortby(7) = "PEOPLE_OFFICE"
	sortby(8) = "PEOPLE_OFFICE DESC"
	sortby(9) = "PEOPLE_CELL"
	sortby(10)= "PEOPLE_CELL DESC"
	urlSort = "group_clients.asp?groupID=" & groupId
	
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 47 Order By word_id"				
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
	set rstitle = Nothing%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body style="margin:0px">
<table border="0" bordercolor="navy" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF dir="<%=dir_var%>">
<tr><td width="100%" class="page_title" dir="<%=dir_obj_var%>">&nbsp;<!--קבוצת דיוור--><%=arrTitles(15)%>&nbsp;<font color="#6E6DA6"><%=groupName%></font></td></tr>         
<tr><td width="100%" colspan=2>
 <table border=0 width=100% cellpadding=0 cellspacing=1 bgcolor="#FFFFFF">			
	 <tr>								
	    <td width=80 nowrap align="<%=align_var%>" class="title_sort<%if trim(sort)="9" OR trim(sort)="10" then%>_act<%end if%>"><%if trim(sort)="9" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="10" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=9" title="למיון בסדר עולה"><%end if%><!--טלפון נייד--><%=arrTitles(17)%>&nbsp;<%if trim(sort)="9" OR trim(sort)="10" then%><img src="../../images/arrow_<%if trim(sort)="9" then%>bot<%elseif trim(sort)="10" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="4"><%End If%></a></td>
	    <td width=100 nowrap align="<%=align_var%>" class="title_sort<%if trim(sort)="7" OR trim(sort)="8" then%>_act<%end if%>"><%if trim(sort)="7" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="8" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=7" title="למיון בסדר עולה"><%end if%><!--תפקיד--><%=arrTitles(16)%>&nbsp;<%if trim(sort)="7" OR trim(sort)="8" then%><img src="../../images/arrow_<%if trim(sort)="7" then%>bot<%elseif trim(sort)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="4"><%End If%></a></td>
		<td width=150 nowrap align="<%=align_var%>" class="title_sort<%if trim(sort)="5" OR trim(sort)="6" then%>_act<%end if%>"><%if trim(sort)="5" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="6" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=5" title="למיון בסדר עולה"><%end if%><!--חברה--><%=arrTitles(4)%>&nbsp;<%if trim(sort)="5" OR trim(sort)="6" then%><img src="../../images/arrow_<%if trim(sort)="5" then%>bot<%elseif trim(sort)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="4"><%End If%></a></td>
		<td width=150 nowrap align="<%=align_var%>" class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=3" title="למיון בסדר עולה"><%end if%><!--שם מלא--><%=arrTitles(5)%>&nbsp;<%if trim(sort)="3" OR trim(sort)="4" then%><img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="4"><%End If%></a></td>																					
		<td width=100% style="line-height:18px" align="<%=align_var%>" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=1" title="למיון בסדר עולה"><%end if%>Email&nbsp;<%if trim(sort)="1" OR trim(sort)="2" then%><img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"><%End If%></a></td>
	 </tr>					
	 <%					
		PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
		If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
     			PageSize = 10
		End If	
		sqlStr = "Select PEOPLE_ID,PEOPLE_EMAIL,PEOPLE_COMPANY,PEOPLE_NAME,PEOPLE_OFFICE,PEOPLE_CELL FROM PEOPLES where ORGANIZATION_ID=" &_
		OrgID & " and groupId="& groupId & " order by " & sortby(sort)
		''Response.Write sqlStr
		set client = con.GetRecordSet(sqlStr)
		if not client.eof then
			client.PageSize=PageSize
			client.AbsolutePage=Page
			recCount=client.RecordCount 
			NumberOfPages=client.PageCount
			i=1		       
			j=0 
		do while (not client.EOF and i<=client.PageSize)
			PEOPLE_ID=client("PEOPLE_ID") 
			PEOPLE_EMAIL = client("PEOPLE_EMAIL")
			PEOPLE_COMPANY = client("PEOPLE_COMPANY")
			PEOPLE_NAME = client("PEOPLE_NAME")
			PEOPLE_OFFICE = client("PEOPLE_OFFICE")
			PEOPLE_CELL = client("PEOPLE_CELL")			
		%>
		<tr height=20>			
		    <td class="card" align="<%=align_var%>"><%=PEOPLE_CELL%>&nbsp;</td>
			<td class="card" align="<%=align_var%>"><%=PEOPLE_OFFICE%>&nbsp;</td>
			<td class="card" align="<%=align_var%>"><%=PEOPLE_COMPANY%>&nbsp;</td>
			<td class="card" align="<%=align_var%>"><%=PEOPLE_NAME%>&nbsp;</td>
			<td class="card" align="<%=align_var%>"><%=PEOPLE_EMAIL%>&nbsp;</td>
		</tr>					
		<%
		client.movenext
		i=i+1
	 loop
	 url = urlSort & "&sort=" & sort
	 if NumberOfPages > 1 then
	 %>
	<tr>
	   <td width="100%" align=middle colspan="5" nowrap  bgcolor="#e6e6e6">
	   <table border="0" cellspacing="0" cellpadding="2" ID="Table2">               
		<% If NumberOfPages > 10 Then 
				num = 10 : numOfRows = cInt(NumberOfPages / 10)
			else num = NumberOfPages : numOfRows = 1    	                      
			End If
		%>				    
		<tr>
			<%if numOfRow <> 1 then%> 
			<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter title="לדפים הקודמים" href="<%=url%>&page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			<%end if%>
			<td><font size="2" color="#001c5e">[</font></td>
			<%for i=1 to num
				If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
					if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
						<td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
					<%Else%>
						<td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=url%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
					<%End If
				End If
			  Next%>				    
			<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
			<%if NumberOfPages > cint(num * numOfRow) then%>  
			<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter title="לדפים הבאים" href="<%=url%>&page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
			<%end if%>   
			</tr></table></td>
		</tr>			
		<%client.close 
		set client=Nothing
		End if
		%>
		<tr>
			<td colspan="5" height=18 bgcolor="#e6e6e6" align=center class="card" dir="ltr" style="color:#6E6DA6;font-weight:600"><!--נמצאו--><%=arrTitles(6)%>&nbsp;<%=recCount%>&nbsp;<!--נמענים--><%=arrTitles(7)%></td>
		</tr>
		<% Else %>
		<tr><td colspan="5"  bgcolor="#e6e6e6" align=center style="font-size=12px;font-weight=550;color=red;" dir="<%=dir_obj_var%>"><!--לא נמצאו נמענים--><%=arrTitles(8)%></td></tr>
		<% End If%>	
		</table></td>		
		</tr></table>
</body>
<%set con=nothing%>
</html>

<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%	OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
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

	If trim(Request.QueryString("pageNum"))<>"" Then
		pageNum = cInt(trim(Request.QueryString("pageNum")))
	Else
		pageNum = 1
	End If  
	 
	if trim(Request.QueryString("numOfRow"))<>"" then
		numOfRow = trim(Request.QueryString("numOfRow"))
	else
		numOfRow = 1
	end if  

	dim sortby(2)	
	sortby(1) = " PEOPLE_EMAIL"
	sortby(2) = " PEOPLE_EMAIL DESC"	
	sort = trim(Request("sort"))
	
	if Len(sort) = 0 then
		sort = 1
	end if	
	urlSort = "notvalid_mails.asp?1=1"		

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
</head>
<body style="margin:0px" onload="window.focus();">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF 
dir="<%=dir_var%>"><tr>
<td width="100%" class="page_title" dir="<%=dir_var%>" colspan=2>
<span style="width:80%;text-align:center">&nbsp;<%=arrTitles(11)%>&nbsp;מיילים לא תקניים&nbsp;-&nbsp;<font color="#595698"><%=org_name%></font></span>
<a class="button_edit_1" style="width:115;" href="excelDownloadnotvalid.asp"><!--הורדה ל-Excel--><%=arrTitles(12)%></a>
</td></tr>         
<tr><td width="100%" valign="top">
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_obj_var%>">
<tr><td width="100%" valign="top" align="<%=align_var%>">    
<%'//start of PEOPLES%>	
<table border=0 width=100% cellpadding=0 cellspacing=1>
<%  Set cnn = Server.CreateObject("ADODB.Connection")
	cnn.ConnectionString = connString
	cnn.Open()
	Set sel_command = Server.CreateObject("ADODB.COMMAND")
	Set sel_command.ActiveConnection = cnn
    sel_command.CommandText = " SELECT PEOPLE_ID, PEOPLE_EMAIL FROM PEOPLES WHERE (ORGANIZATION_ID = ?) " & _
    " AND (IsEmailValid = 0) ORDER BY " & trim(sortby(sort))
    sel_command.CommandType= 1   
    sel_command.Parameters.Append sel_command.CreateParameter("@OrgID", 3, 1, 4, cInt(OrgID))
    Set rs_p = Server.CreateObject("ADODB.RECORDSET")
    rs_p.CursorLocation = 3
    rs_p.Open sel_command%>
<tr>
	<td width=100% align="<%=align_var%>" dir="<%=dir_var%>" class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self"><!--מייל--><%=arrTitles(4)%><img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>
</tr>					
<%
	If not rs_p.eof Then
		If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
			PageSize = RowsInList
		Else	
     		PageSize = 25
		End If	   
	    rs_p.PageSize = PageSize
		rs_p.AbsolutePage=pageNum
		recCount=rs_p.RecordCount 	
		Mail_count=rs_p.RecordCount	
		NumberOfPages = rs_p.PageCount
		i=1
		j=0
		do while (not rs_p.eof and j<rs_p.PageSize)
		    PEOPLE_ID = trim(rs_p("PEOPLE_ID"))
			PEOPLE_EMAIL = trim(rs_p("PEOPLE_EMAIL"))%>		
	<tr>
		<td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;<%=PEOPLE_EMAIL%>&nbsp;</td>
	</tr>					
<%
		rs_p.movenext
		j=j+1
		loop
		set rs_p = nothing
%>
<% if NumberOfPages > 1 then%>
<tr class="card">
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
		<td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&pageNum=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" title="לדפים הקודמים"><b><<</b></a></td>			                
		<%end if%>
	        <td><font size="2" color="#060165">[</font></td>
	        <%for i=1 to num
	            If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	            if CInt(pageNum)=CInt(i+10*(numOfRow-1)) then %>
		            <td align="center"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	            <%else%>
	                <td align="center"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&pageNum=<%=i+10*(numOfRow-1)%>&numOfRow=<%=numOfRow%>"><%=i+10*(numOfRow-1)%></a></td>
	            <%end if
	            end if
	        next%>	    
			<td><font size="2" color="#060165">]</font></td>
		<%if NumberOfPages > cint(num * numOfRow) then%>  
			<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&pageNum=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" title="לדפים הבאים"><b>>></b></a></td>
		<%end if%>   
		</tr> 				  	             
	</table></td></tr>				
	<%End If%>		
	<tr><td align=center height=20px class="card"  colspan=11><font style="color:#6E6DA6;font-weight:600"><!--נמצאו--><%=arrTitles(5)%>&nbsp;<%=Mail_count%>&nbsp;מיילים לא תקניים</font></td></tr>								
</table></td>
<%		
Else
%>
  <tr><td class="title_sort1" align=center colspan=5>לא נמצאו מיילים לא תקניים</td></tr>
<%
End If
%>					
	</table></td>
<%'//end of PEOPLES%>			
</td>	
</tr>
<tr><td height=10 nowrap></td></tr>
</table>
</body>
<%set con=nothing
cnn.Close()
Set cnn = Nothing%>
</html>

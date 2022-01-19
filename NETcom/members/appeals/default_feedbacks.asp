
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr><td width=100%>
<!-- #include file="../../logo_top.asp" -->
</td></tr>
<%numOftab = 2%>
<%numOfLink = 3%>
<%topLevel2 = 21 'current bar ID in top submenu - added 03/10/2019%>
<tr><td width=100%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title">&nbsp;</td></tr>
<tr><td height=20 nowrap></td></tr>
<tr>		
    <td width="100%" align="center" valign=top>
	<table border="0" cellspacing="1" width=60% align=center cellpadding="0" dir="<%=dir_obj_var%>">	
	<%If trim(lang_id) = "1" Then
			arr_Status = Array("","חדש","בטיפול","סגור")	
		Else
			arr_Status = Array("","new","active","close")	
		End If
		
		sqlstr = "SELECT product_id, product_name FROM Products WHERE FORM_MAIL = '1' AND "&_
		" product_number = '0' AND ORGANIZATION_ID=" & OrgID & " ORDER BY product_name"
		'Response.Write sqlstr
		'Response.End
		set rs_products = con.GetRecordSet(sqlstr)
		if not rs_products.eof then 
			arr_products = rs_products.getRows()
		end if
		set rs_products=nothing	
					
		If IsArray(arr_products) Then	%>	
	   <tr><th width="100%" align="<%=align_var%>" class="form_title" dir="<%=dir_obj_var%>" rowspan=2>
		<%If trim(lang_id) = "1" Then%>
		&nbsp;בחר טופס לקבלת המשובים:&nbsp;
		<%Else%>
		&nbsp;Choose a form to get form data:&nbsp;
		<%End If%></th>
		<th class=textp colspan="<%=Ubound(arr_Status)%>" align=center>משובים</th></tr>	
		<tr><%For i=1 To Ubound(arr_Status)%>
	    <th width=70 class="textp" nowrap align="center" dir="<%=dir_obj_var%>"><%=arr_Status(i)%></th>
		<%Next%></tr>
	<%For i=0 To  Ubound(arr_products,2) 
			prod_Id =  trim(arr_products(0,i))
			product_name =  trim(arr_products(1,i))
			app_count = Array()
			redim app_count(Ubound(arr_Status))		
			For jj=1 To Ubound(arr_Status)
			sqlstr = "exec dbo.get_feedbacks '','','','"&jj&"','"&OrgID&"','','','','','','','','"&prod_Id&"','','',''"	
			'Response.Write sqlStr
			set app=con.GetRecordSet(sqlStr)
			app_count(jj) = app.RecordCount
			set app=Nothing
			Next
	%>
		<tr height=24><td width="100%" align="<%=align_var%>" class=card dir="<%=dir_obj_var%>">
		<A class="link_categ" href="feedbacks.asp?prodId=<%=prod_Id%>" target=_self>&nbsp;<b><%=product_name%></b>&nbsp;&nbsp;&nbsp;&nbsp;<%=strCount%></A>
		</td>
		<%For jj=1 To Ubound(arr_Status)%>
		<td width=70 nowrap width="100%" align="center" class=card dir="<%=dir_obj_var%>">
		<A class="status_num<%=jj%>" href="feedbacks.asp?prodId=<%=prod_Id%>&cod=<%=jj%>" target=_self><%=trim(app_count(jj))%></A></td>
		<%Next%>
		</tr>								
<%Next%>
<% Else	%>
<tr>		
    <th width="100%" align="<%=align_var%>" class="form_title" dir="<%=dir_obj_var%>">
    <%If trim(lang_id) = "1" Then%>
	&nbsp;לא נמצאו טפסי דיוור&nbsp;
	<%Else%>
	&nbsp;Not found mailing forms&nbsp;
	<%End If%>
	</th>
</tr>
<%		
 End If	
%>	
</table></td></tr>
<tr><td height=20 nowrap></td></tr>
</table>
</body>
<%  set con = nothing %>
</html>


<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
<tr><td width="100%">
<!-- #include file="../../logo_top.asp" -->
</td></tr>
<%numOfTab = 1%>
<%numOfLink = 0%>
<%topLevel2 = 15 'current bar ID in top submenu - added 03/10/2019%>

<tr><td width="100%">
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title">&nbsp;</td></tr>
<tr><td height="20" nowrap></td></tr>
<tr>		
    <td width="100%" align="center" valign="top">
	<table border="0" cellspacing="1" width="60%" align="center" cellpadding="0" dir="<%=dir_obj_var%>">	
    <% arr_Status = session("arr_Status") 	%>
	<tr>	
    <th width="100%" align="<%=align_var%>" class="form_title" dir="<%=dir_obj_var%>" rowspan=2>
    <%If trim(lang_id) = "1" Then%>	&nbsp;בחר טופס לקבלת הנתונים:&nbsp;
	<%Else%>	&nbsp;Choose a form to get form data:&nbsp;<%End If%>
	</th>	
	<th class=textp colspan="<%=Ubound(arr_Status)%>" align="center">
	<%If trim(lang_id) = "1" Then%>
	משובים
	<%Else%>
	Appeals
	<%End if%>
	</th>	
	</tr>	
	<tr>    
	<%For i=1 To Ubound(arr_Status)%>
	<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>"><%=arr_Status(i, 1)%></th>
	<%Next%>	
    </tr>    
	<%sqlstr = "Exec dbo.GetProductsCount '" & UserID & "','" & OrgID & "','" & is_groups & "'" 	
		'Response.Write sqlstr
		'Response.End 
		set rs_products = con.GetRecordSet(sqlstr)
		If not rs_products.eof Then
		While not rs_products.eof 
			prod_Id = trim(rs_products(0))
			product_name = trim(rs_products(1))	%>
	<tr height=24>
	<td width="100%" align="<%=align_var%>" class=card dir="<%=dir_obj_var%>">
	<A class="link_categ" href="appeals.asp?prodId=<%=prod_Id%>" target=_self>&nbsp;<b><%=product_name%></b></A>
	</td>	
	<%For jj=1 To Ubound(arr_Status)%>
	<td width=70 nowrap width="100%" align="center" class=card dir="<%=dir_obj_var%>">
	<A class="status_num" style="background-color:<%=trim(arr_Status(jj, 2))%>" 
	href="appeals.asp?prodId=<%=prod_Id%>&app_status=<%=trim(arr_Status(jj, 0))%>" target="_self"><%=trim(rs_products(1 + jj))%></a></td>
	<%Next%>		
	</tr>								
<%	rs_products.moveNext
	Wend
	Set rs_products = Nothing
    Else	%>
	<tr>		
    <th width="100%" align="<%=align_var%>" class="form_title" dir="<%=dir_obj_var%>">
    <%If trim(lang_id) = "1" Then%>	&nbsp;לא קיימים טפסים במערכת עבורך, באפשרותך לבנות טופס בעזרת מחולל טפסים.&nbsp;
	<%Else%>	&nbsp;Not found forms&nbsp;<%End If%>
	</th>
    </tr>
<%		
 End If	
%>	
</table></td></tr>
<tr><td height="20" nowrap></td></tr>
</table>
</body>
</html>
<%  set con = nothing %>
<%@ Language=VBScript%>
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
<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>" ID="Table1">
<tr><td width="100%">
<!-- #include file="../../logo_top.asp" -->
</td></tr>
<%numOfTab = 5%>
<%numOfLink = 0%>
<tr><td width="100%">
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title">&nbsp;</td></tr>
<tr><td height="20" nowrap></td></tr>
<tr>		
    <td width="100%" align="center" valign="top">
	<table border="0" cellspacing="1" width="60%" align="center" cellpadding="0" dir="<%=dir_obj_var%>" ID="Table2">	
	<tr>	
    <th width="100%" align="<%=align_var%>" class="form_title" dir="<%=dir_obj_var%>" rowspan=2>
   	&nbsp;ריכוז משימות:&nbsp;</th>	
	<th class=textp colspan="2" align="center">
	משימות קישורית
	</th>	
	<th class=textp colspan="2" align="center">
	משימות מעקב 
	</th>	<th class=textp colspan="2" align="center">
	רישום ופתיחת טופס או עדכון כרטיס
	</th>	
	</tr>	
	<tr>    

	<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>">חדש</th>
		<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>">בטיפול</th>
	
	<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>">חדש</th>
	
	<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>">בטיפול</th>
	
	<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>">חדש</th>
	
	<th width=70 nowrap align="center" class="textp" dir="<%=dir_obj_var%>">בטיפול</th>
	
	
    </tr>    
	<%sqlstr = "Exec dbo.GetTasksCount '"  & OrgID  & "'" 	
		'Response.Write sqlstr
		'Response.End 
		set rs_products = con.GetRecordSet(sqlstr)
		If not rs_products.eof Then
		While not rs_products.eof 
			prod_Id = trim(rs_products(0))
			product_name = trim(rs_products(1))	%>
	<tr height=24>
	<td width="100%" align="<%=align_var%>" class=card dir="<%=dir_obj_var%>">&nbsp;<a class="link_categ" href="javascript:void(0)" onclick="window.open('taskDetailsAll.asp?pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=800, scrollbars=1');"><b><%=product_name%></b></a></td>	
	<%'For jj=1 To Ubound(arr_Status)%>
   <td class=card valign="top" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num1" href="javascript:void(0)" onclick="window.open('taskDetails.asp?status=1&Ttype=2026&pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=700, scrollbars=1');"><%=rs_products("KNew")%></a></td>	  
   <td class=card valign="top" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num2" href="javascript:void(0)" onclick="window.open('taskDetails.asp?status=2&Ttype=2026&pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=700, scrollbars=1');"><%=rs_products("KTipul")%></a></td>	  
   <td class=card valign="top" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num1" href="javascript:void(0)" onclick="window.open('taskDetails.asp?status=1&Ttype=2005&pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=700, scrollbars=1');"><%=rs_products("MaakavNew")%></a></td>	  
   <td class=card valign="top" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num2" href="javascript:void(0)" onclick="window.open('taskDetails.asp?status=2&Ttype=2005&pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=700, scrollbars=1');"><%=rs_products("MaakavTipul")%></a></td>	  
   <td class=card valign="top" align="center" dir="<%=dir_obj_var%>"><a class="task_status_num1" href="javascript:void(0)" onclick="window.open('taskDetails.asp?status=1&Ttype=2027&pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=700, scrollbars=1');"><%=rs_products("KartisNew")%></a></td>	  
   <td class=card valign="top" align="center" dir="<%=dir_obj_var%>"><a href="javascript:void(0)" class="task_status_num2" onclick="window.open('taskDetails.asp?status=2&Ttype=2027&pUserId=<%=prod_Id%>','winT','top=20, left=10, width=1300, height=700, scrollbars=1');"><%=rs_products("KartisTipul")%></a></td>	  
	<%'Next%>		
	</tr>								
<%	rs_products.moveNext
	Wend
	Set rs_products = Nothing
    Else	%>
	<tr>		
    <th  colspan=7 width="100%" align="<%=align_var%>" class="form_title" dir="<%=dir_obj_var%>">
    <%If trim(lang_id) = "1" Then%>	&nbsp;לא קיימים נתונים במערכת.&nbsp;
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
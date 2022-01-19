
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
	<head>
		<!-- #include file="../../title_meta_inc.asp" -->
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	</head>
	<body>
		<table border="0" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>">
			<tr>
				<td width="100%">
					<!-- #include file="../../logo_top.asp" -->
				</td>
			</tr>
			<%numOftab = 1%>
			<%numOfLink = 2%>
			<%topLevel2 = 16 'current bar ID in top submenu - added 03/10/2019%>
			<tr>
				<td width="100%">
					<!--#include file="../../top_in.asp"-->
				</td>
			</tr>
			<tr>
				<td class="page_title">&nbsp;</td>
			</tr>
			<tr>
				<td height="20" nowrap></td>
			</tr>
			<tr>
				<td width="100%" align="center" valign="top">
					<table border="0" cellspacing="1" width=60% align=center cellpadding="0" dir="<%=dir_obj_var%>">
						<%  	If is_groups = 0 Then
		sqlstr = "Select product_id, product_name from Products Where "&_
		" product_number = '0' AND (SHOW_INSITE = 1) AND ORGANIZATION_ID=" & OrgID & " order by product_name"
		Else
		'sqlstr = "Select DISTINCT Products.product_id, Products.product_name from Products Inner Join Users_To_Products "&_
		'" ON Products.Product_ID = Users_To_Products.Product_ID WHERE Users_To_Products.User_ID = " & UserID &_
		'" And Products.product_number = '0' and Products.ORGANIZATION_ID=" & OrgID & " order by Products.product_name"
		sqlstr = "exec dbo.get_products_list_enter '" & OrgID & "','" & UserID & "'"
		End If
		'Response.Write sqlstr
		'Response.End
		set rs_products = con.GetRecordSet(sqlstr)
		if not rs_products.eof then 
			ResProductsList = rs_products.getRows()		
		end if
		set rs_products=nothing				
		If IsArray(ResProductsList) Then		
	%>
						<tr>
							<th width="100%" align="<%=align_var%>" class="form_title" dir="<%=dir_obj_var%>">
								<%If trim(lang_id) = "1" Then%>
								&nbsp;בחר טופס להזנת הנתונים:&nbsp;
								<%Else%>
								&nbsp;Choose a form to enter form data:&nbsp;
								<%End If%>
							</th>
						</tr>
						<%For i=0 To Ubound(ResProductsList,2)		
		prod_Id = trim(ResProductsList(0,i))
		product_name = trim(ResProductsList(1,i))
	%>
						<%if prod_Id<>16504 then%>
						<tr>
							<td width="100%" align="<%=align_var%>" class=card dir="<%=dir_obj_var%>">
								<A class="link_categ" href="appeal.asp?quest_id=<%=prod_Id%>" target=_self>&nbsp;<b><%=product_name%></b></A>
							</td>
						</tr>
						<%else%>
						<tr>
							<td width="100%" align="<%=align_var%>" class=card dir="<%=dir_obj_var%>">
								<a class="link_categ" href="javascript:void(0);" onclick="javascript:window.open('../appeals/TripForm16504.aspx?quest_id=<%=prod_Id%>','','top=50,left=100,width=600,height=500,scrollbars=1,resizable=1');" >&nbsp;<b><%=product_name%></b></a>
				</td>
						</tr>
						<%end if%>
						<%	Next
    Else	%>
						<tr>
							<th width="100%" align="<%=align_var%>" class="form_title" dir="<%=dir_obj_var%>">
								<%If trim(lang_id) = "1" Then%>
								&nbsp;לא קיימים טפסים במערכת עבורך, באפשרותך לבנות טופס בעזרת מחולל 
								טפסים.&nbsp;
								<%Else%>
								&nbsp;Not found forms&nbsp;
								<%End If%>
							</th>
						</tr>
						<%		
 End If	
%>
					</table>
				</td>
			</tr>
			<tr>
				<td height="20" nowrap></td>
			</tr>
		</table>
	</body>
	<%  set con = nothing %>
</html>

<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css" />
<script LANGUAGE="JavaScript">
<!--
function CheckDelProd() {
  <%
     If trim(lang_id) = "1" Then
        str_confirm = "?האם ברצונך למחוק את הטופס רישום"
     Else
		str_confirm = "Are you sure want to delete the form?"
     End If   
  %>
  return (confirm("<%=str_confirm%>"))    
}

var prod;
function openPreview(prodId)
{
	prod = window.open("check_form.asp?prodId="+prodId,"Product","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=700, height=500, left=10, top=20");
	if((prod.document != null) && (prod.document.body != null))
	{
		prod.document.title = "Product Preview";
		prod.document.body.style.margintop  = 20;
	}	
	return false;
}

function closeProduct(quest_id,show_insite)
{
	if(quest_id != null)
	{
		if(show_insite == "True")
			msg = "<%=Space(35)%>,שים לב\n\n.מרגע זה טופס ייפתח להזנה\n\n<%=Space(10)%>? האם ברצונך להמשיך"
		else
		    msg = "<%=Space(35)%>,שים לב\n\n.מרגע זה טופס ייחסם להזנה\n\n<%=Space(10)%>? האם ברצונך להמשיך"
		
		if(window.confirm(msg))
			document.location.href = "questions.asp?prodId="+quest_id+"&show="+show_insite;
	}
	return false;

}
//-->
</script>  
</head>
<%
	reqprodid = Request.QueryString("prodId")
	if request("delProd")="1" and reqprodid <> nil then
	   set pr=con.GetRecordSet("select FILE_ATTACHMENT from products where product_id=" & reqprodid)
		if not pr.eof then
			fileName1=pr("FILE_ATTACHMENT")			
			If trim(fileName1) <> "" Then
			sqlstr = "Select TOP 1 product_id FROM products WHERE FILE_ATTACHMENT LIKE '" & sFix(fileName1) & "' AND product_id <> " & reqprodid
			set rs_check = con.getRecordSet(sqlstr)
			if rs_check.eof then
				set fs=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!   			
   				file_path="../../../download/products/" & fileName1
   				'Response.Write fs.FileExists(server.mappath(file_path))
				'Response.End
				if fs.FileExists(server.mappath(file_path)) then
					set a = fs.GetFile(server.mappath(file_path))
					a.delete			
				end if	
			end if
			set rs_check = nothing	
			End If				
		end if
		con.ExecuteQuery("Delete from Form_select where Field_Id in (Select Field_Id from Form_Field where product_id="& reqprodid & " and ORGANIZATION_ID=" & trim(OrgID) & ")")
		con.ExecuteQuery("Delete from Form_Field where product_id="&reqprodid & " and ORGANIZATION_ID=" & trim(OrgID) )
		con.ExecuteQuery("Delete from Users_To_Products where product_id="&reqprodid)
		con.ExecuteQuery("Delete from Statistic WHERE product_id="&reqprodid)		
		con.ExecuteQuery("Delete from products where product_id="&reqprodid & " and ORGANIZATION_ID=" & trim(OrgID) )
	end if
	
	if Request.QueryString("show") = "True" and reqprodid <> nil then
		con.ExecuteQuery("UPDATE products SET show_insite=1 WHERE (product_id="&reqprodid & ") AND ORGANIZATION_ID=" & trim(OrgID) )
	elseif Request.QueryString("show") = "False" and reqprodid <> nil then
		con.ExecuteQuery("UPDATE products SET show_insite=0 WHERE (product_id="&reqprodid & ") AND ORGANIZATION_ID=" & trim(OrgID) )
	end if
	
	dim sortby(16)	 
	sortby(1) = " product_name"
	sortby(2) = " product_name DESC"
	sortby(3) = " product_id"
	sortby(4) = " product_id DESC"

	sort = Request("sort")	
	if sort = nil then
		sort = 4
	end if 
	 
	urlSort = "questions.asp?1=1"
	if Request("Page")<>"" then
		Page=request("Page")
	else
		Page=1
	end if

	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 17 Order By word_id"				
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
<body>
<!--#include file="../../logo_top.asp"-->
<%numOfTab = 1%>
<%numOfLink = 1%>
<%topLevel2 = 14 'current bar ID in top submenu - added 03/10/2019%>

<!--#include file="../../top_in.asp"-->
<table width="100%" cellspacing="0" cellpadding="0" align=center border="0" bgcolor="#ffffff">
<tr><td class="page_title" width=100%>&nbsp;</td></tr>
<tr>
<td align=center width=100% valign=top>
<table cellpadding=0 cellspacing=0 width=100% dir="<%=dir_var%>">
<tr><td width=100% valign=top>
<table width=100% align=center border="0" cellpadding=1 cellspacing=1>
<tr>
	<td width="50" nowrap class="title_sort" align=center><!--מחיקה--><%=arrTitles(3)%></td>
	<td width="50" nowrap class="title_sort" align=center>הצג בהזנת טפסים</td>
	<td width="50" nowrap class="title_sort" align=center><!--שינוי פרטי הטופס--><%=arrTitles(4)%></td>	
	<td width="50" nowrap class="title_sort" align=center><!--שכפול טופס--><%=arrTitles(5)%></td>
	<td width="50" nowrap class="title_sort" align=center><!--דוח מכירות--><%=arrTitles(28)%></td>
	<td width="50" nowrap class="title_sort" align=center><!--מקורות פרסום--><%=arrTitles(26)%></td>
	<td width="50" nowrap class="title_sort" align=center><!--כתובת--><%=arrTitles(6)%></td>
	<td width="50" nowrap class="title_sort" align=center><!--קליטת נתוני לקוח למערכת--><%=arrTitles(22)%></td>
	<td width="50" nowrap class="title_sort" align=center><!--טופס קשרי לקוחות--><%=arrTitles(7)%></td>
	<td width="50" nowrap class="title_sort" align=center><!--טופס דיוור--><%=arrTitles(8)%></td>
	<td width="50" nowrap class="title_sort" align=center><!--מצבי סגירה--><%=arrTitles(27)%></td>
	<td width="50" nowrap class="title_sort" align=center><!--מענים אוטומטיים--><%=arrTitles(25)%></td>
	<%If trim(is_groups) = "1" Then%>	
	<td width="50" nowrap class="title_sort" align=center><!--הרשאות--><%=arrTitles(24)%></td>
	<%End If%>
	<td width="50" nowrap class="title_sort" align=center><!--תצוגה מקדימה--><%=arrTitles(9)%></td>
	<td width="100%" align="<%=align_var%>" class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self">&nbsp;<!--שם הטופס--><%=arrTitles(10)%>&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>
</tr>
<%  set rs_products = con.GetRecordSet("Select product_id, product_name, open_form, form_client, form_mail, add_client, show_insite from products where product_number = '0' and ORGANIZATION_ID=" & trim(OrgID) & " order by " & sortby(sort))
    if not rs_products.eof then 
		If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
			PageSize = RowsInList
		Else	
     		PageSize = 10
		End If	        
		rs_products.PageSize = PageSize
		rs_products.AbsolutePage=Page
		recCount=rs_products.RecordCount 	
		pages_count=rs_products.RecordCount	
		NumberOfPages = rs_products.PageCount
		i=1
		j=0
		do while (not rs_products.eof and j<rs_products.PageSize)
    	prod_Id = rs_products("product_id")
    	If trim(prod_Id) <> "" Then
    		sqlstr = "Select TOP 1 QUESTIONS_ID From APPEALS WHERE QUESTIONS_ID = " & prod_Id
    		set rs_check = con.getRecordSet(sqlstr)
    		If not rs_check.eof Then
    			is_send = true
    		Else
    			is_send = false
    		End If	
    		set rs_check = Nothing
    	End If
    	product_name = rs_products("product_name")
    	open_form = rs_products("open_form")
    	prodClient = rs_products("FORM_CLIENT")
		prodMail = rs_products("FORM_MAIL")
		add_client = rs_products("add_client") 
		show_insite = cBool(trim(rs_products.Fields("show_insite")))%>
<tr>
	<td align="center" class="card" nowrap>&nbsp;
	<%If is_send = false Then%>
	<a href="questions.asp?prodId=<%=prod_Id%>&delProd=1" ONCLICK="return CheckDelProd()"><IMG SRC="../../images/delete_icon.gif" BORDER=0 name=word14 title="<%=arrTitles(14)%>"></a>&nbsp;
	<%Else%>
	<input type=image SRC="../../images/delete_icon.gif" border=0 Onclick="window.alert('<%=Space(12)%>שים לב, לא ניתן למחוק את הטופס\n\nמפני שקיימים מענים של טופס זה במערכת\n\n<%=Space(26)%>על מנת למחוק את הטופס\n\n<%=Space(9)%>עליך למחוק תחילה את המענים הנל');return false;">
	<%End If%></td>
	<td align="center" class="card" nowrap><%if show_insite = false then%><IMG SRC="../../images/lamp_off.gif" BORDER=0 alt="לא מופיע בהזנת טפסים" class=hand onClick="return closeProduct('<%=prod_Id%>','<%=not show_insite%>')"><%else%><IMG SRC="../../images/lamp_on.gif" BORDER=0 alt="מופיע בהזנת טפסים" class=hand onClick="return closeProduct('<%=prod_Id%>','<%=not show_insite%>')"><%end if%></td>
	<td align="center" class="card" nowrap>&nbsp;<a href="addquestions.asp?prodId=<%=prod_Id%>"><IMG SRC="../../images/edit_icon.gif" BORDER=0 name=word15 title="<%=arrTitles(15)%>"></a>&nbsp;</td>	
	<td align="center" class="card" nowrap>&nbsp;<a href="addquestions.asp?addcopy=<%=prod_Id%>"><IMG SRC="../../images/copy_icon.gif" BORDER=0 name=word16 title="<%=arrTitles(16)%>"></a>&nbsp;</td>
	<td align="center" class="card" nowrap><a class=link1 href="reports.asp?prodId=<%=prod_Id%>" target=_self><img src="../../images/graph.gif" border=0 hspace=0 vspace=0 title="<%=arrTitles(28)%>"></a></td>
	<td align="center" class="card" nowrap><a class=link1 href="../statistic/products.asp?prodId=<%=prod_Id%>" target=_self><img src="../../images/plus_sub.gif" border=0 hspace=0 vspace=0 title="<%=arrTitles(26)%>"></a></td>
	<td align="center" class="card" nowrap><%If trim(open_form) = "1" Then%><a class=link1 href="<%=strLocal%>tophes/default.asp?<%=Encode("P=" & prod_Id)%>" target=_blank><img src="../../images/icon_explor.gif" border=0 hspace=0 vspace=0 title="<%=strLocal%>tophes/default.asp?<%=Encode("P=" & prod_Id)%>"></a><%End If%></td>
	<td align="center" class="card" nowrap><%If trim(add_client) <> "" Then%><img src="../../images/vi.gif" border=0 hspace=0 vspace=0><%End If%></td>
	<td align="center" class="card" nowrap><%If trim(prodClient) = "1" Then%><img src="../../images/vi.gif" border=0 hspace=0 vspace=0 name=word17 title="<%=arrTitles(17)%>"><%End If%></td>
	<td align="center" class="card" nowrap><%If trim(prodMail) = "1" Then%><img src="../../images/vi.gif" border=0 hspace=0 vspace=0 name=word18 title="<%=arrTitles(18)%>"><%End If%></td>
	<td align="center" class="card" nowrap>&nbsp;<a href="closings.asp?prodId=<%=prod_Id%>"><IMG SRC="../../images/notok_icon.gif" BORDER=0></a>&nbsp;</td>
	<td align="center" class="card" nowrap>&nbsp;<a href="responses.asp?prodId=<%=prod_Id%>"><IMG SRC="../../images/ok_icon.gif" BORDER=0></a>&nbsp;</td>	
	<%If trim(is_groups) = "1" Then%>
	<td align="center" class="card" nowrap>&nbsp;<a href="permitions.asp?prodId=<%=prod_Id%>"><IMG SRC="../../images/signIn_icon.gif" BORDER=0 name=word23 title="<%=arrTitles(23)%>"></a>&nbsp;</td>	
	<%End If%>
	<td align="center" class="card" nowrap>&nbsp;<a href="#" OnClick="return openPreview('<%=prod_Id%>')"><IMG SRC="../../images/preview_icon.gif" BORDER=0 name=word19 title="<%=arrTitles(19)%>"></a>&nbsp;</td>
	<td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a class="link_categ" href="addform.asp?prodId=<%=prod_Id%>"><%=product_name%>&nbsp;&nbsp;</a></td>
</tr>
<%	  rs_products.MoveNext
      j=j+1
	  loop
	set rs_products=nothing%>
<%if NumberOfPages > 1 then%>
<tr class="card">
  <td width="100%" align=center nowrap class="card" colspan="15">
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
		<td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" name=word20 title="<%=arrTitles(20)%>"><b><<</b></a></td>			                
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
			<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" name=word21 title="<%=arrTitles(21)%>"><b>>></b></a></td>
		<%end if%>   
		</tr> 				  	             
	</table></td></tr>				
	<%End If%>		
<tr><td align=center height=20px class="card"  colspan="15"><font style="color:#6E6DA6;font-weight:600"><span id=word11 name=word11><!--נמצאו--><%=arrTitles(11)%></span> <%=pages_count%> <span id=word12 name=word12><!--טפסים--><%=arrTitles(12)%></span></font></td></tr>								
<%		
Else
%>
<tr><td align="center" class="title_sort1" colspan="15">&nbsp;</td></tr>
<%
  end if
  set rs_pages=nothing
%>	
</table></td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="18" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='addquestions.asp'><span id=word13 name=word13><!--הוספת טופס--><%=arrTitles(13)%></span></a></td></tr>
</table></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table>
</body>
<%set con=nothing%>
</html>
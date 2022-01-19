<script language="javascript" type="text/javascript">
<!--
    function getRealTop(imgElem)
	{
		yPos = eval(imgElem).offsetTop;
		tempEl = eval(imgElem).offsetParent;
		while (tempEl != null)
		{
			yPos += tempEl.offsetTop;
			tempEl = tempEl.offsetParent;
		}
		//alert(yPos);
		return yPos;
	}	
    function getRealLeft(imgElem)
	{
		yPos = eval(imgElem).offsetLeft;
		tempEl = eval(imgElem).offsetParent;
		while (tempEl != null)
		{
			yPos += tempEl.offsetLeft;
			tempEl = tempEl.offsetParent;
		}
		//alert(yPos);
		return yPos;
	}		
	
	function tfasimDropDown(obj)
	{
		document.all["div_tfasim"].style.zIndex=11; 
		document.all["div_tfasim"].style.position="absolute";		
		<%If trim(lang_id) = "1" Then%>
		offTop = getRealTop(obj) - 30; 
		offLeft = getRealLeft(obj) - 355;
		<%Else%>
		offTop = getRealTop(obj) - 30; 
		offLeft = getRealLeft(obj) + 110;
		<%End If%>		
		document.all["div_tfasim"].style.top=offTop;
		document.all["div_tfasim"].style.left=offLeft;
		document.all["div_tfasim"].style.display='inline';
		document.all["div_tfasim"].style.visibility="visible";
		return false; 
	}

	function tfasimClose()
	{
		document.all["div_tfasim"].style.display='none';
		document.all["div_tfasim"].style.visibility="hidden"; 
		return false;
	}	
//-->
</script>
<div dir="<%=dir_obj_var%>" id=div_tfasim name=div_tfasim style="display:none;visibility:hidden;width:350;height:150;background-color: #E6E6E6; overflow:scroll; overflow-x:hidden;border:1px solid #808080;SCROLLBAR-FACE-COLOR:#E6E6E6;SCROLLBAR-HIGHLIGHT-COLOR: #FFFFFF; SCROLLBAR-SHADOW-COLOR: #C0C0C0;SCROLLBAR-3DLIGHT-COLOR: #E6E6E6;SCROLLBAR-ARROW-COLOR: #000000; SCROLLBAR-TRACK-COLOR: #F4F4F4;">
<table cellpadding="0" cellspacing="0" border="0" width="100%">
    <tr>
		<td colspan="2" width="100%" bgcolor="#616161" height="1"></td>
	</tr>
	<tr>
        <td width="22" nowrap align=center class=title_form><INPUT type="image" src="../../images/close_icon.gif" onClick="return tfasimClose();" vspace=0 hspace=0 ID="Image1" NAME="Image1">              
        <td width="100%" nowrap class=title_form align="<%=align_var%>">&nbsp;</td>
        </td>                        
    </tr>
    <tr>
		<td colspan="2" width="100%" bgcolor="#616161" height="1"></td>
	</tr>	
   <tr>
        <td width="100%" class="card" align="<%=align_var%>">
        <table border="0" width="100%" cellspacing="5" cellpadding="0">
	<%
		If is_groups = 0 Then
			sqlstr = "SELECT product_id, Replace(product_name, Char(59), '&#59;')  FROM dbo.Products " & _
			" WHERE (isNULL(FORM_CLIENT,0) = 1) AND (product_number = '0') AND (SHOW_INSITE = 1) " & _
			" AND (ORGANIZATION_ID=" & OrgID & ") ORDER BY product_name"
		Else
			sqlstr = "EXEC dbo.get_products_list_enter_cl '" & OrgID & "','" & UserID & "'"
		End If
		'Response.Write sqlstr
		'Response.End
		set rs_products = con.GetRecordSet(sqlstr)
		if not rs_products.eof then 
			arr_products = rs_products.getRows()
		end if
		set rs_products=nothing	
		
		If IsArray(arr_products) Then
			If trim(projectID) = "" And trim(companyID) = "" Then
				from = "comp_list"
			Else
				from = "pop_up"
			End If	 
			
			For i=0 To  Ubound(arr_products,2) 				
					quest_id = arr_products(0,i)   	
					product_name = arr_products(1,i)%>
				<tr><td align="<%=align_var%>" dir="<%=dir_obj_var%>" colspan="2">
				<%if  quest_id=16504 then%>
				<a class="link1" href="javascript:void(0);" onclick="tfasimClose(); javascript:window.open('../appeals/TripForm16504.aspx?quest_id=<%=quest_id%>&companyID=<%=companyID%>&contactID=<%=contactID%>&from=<%=from%>','','top=50,left=100,width=600,height=500,scrollbars=1,resizable=1');" ><%=product_name%></a>
				<%else%>
				<a class="link1" href="javascript:void(0);" onclick="tfasimClose(); javascript:window.open('../appeals/edit_appeal.asp?quest_id=<%=quest_id%>&companyID=<%=companyID%>&contactID=<%=contactID%>&projectID=<%=projectID%>&from=<%=from%>','','top=50,left=100,width=600,height=500,scrollbars=1,resizable=1');" ><%=product_name%></a>
				<%end if%>
     			</td></tr><%	  
		Next	
	Else%>
		<tr><td height="30" nowrap></td></tr> 
		<tr>
		<td width="100%" align="center" class="title_table" dir="<%=dir_obj_var%>">
		<%If trim(lang_id) = "1" Then%>
		&nbsp;לא קיימים טפסים במערכת עבורך,&nbsp;<br>
		&nbsp;באפשרותך לבנות טופס בעזרת מחולל טפסים.&nbsp;
		<%Else%>
		&nbsp;Not found forms&nbsp;
		<%End If%>
		</td>
		</tr>	
	<%End If%>
	</table></td></tr>
</table>
</div>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<style>
.tooltip {
    position: relative;
	display: inline-block;
	 border-bottom: 1px dotted black; /* If you want dots under the hoverable text */
}

.tooltip .tooltiptext {
    visibility: hidden;
    width: 120px;
    background-color: #555;
    color: #fff;
    direction:rtl;
    text-align: center;
    border-radius: 6px;
    padding: 5px 0;
    position: absolute;
    z-index: 1;
    bottom: 125%;
    left: 50%;
    margin-left: -60px;
    opacity: 0;
    transition: opacity 0.3s;
}

.tooltip .tooltiptext::after {
    content: "";
    position: absolute;
    top: 100%;
    left: 50%;
    margin-left: -5px;
    border-width: 5px;
    border-style: solid;
    border-color: #555 transparent transparent transparent;
    
}

.tooltip:hover .tooltiptext {
    visibility: visible;
    opacity: 1;
}
</style>
<%
	If Request.QueryString("deleteId") <> nil Then
		delId = trim(Request.QueryString("deleteId"))
		If Len(delId) > 0 Then
			con.executeQuery("Delete From Suppliers Where supplier_Id = " & delId)
		End If
		Response.Redirect "default.asp"
	End If
	urlSort="default.asp?s=1"
	
	if IsNumeric(request("sort")) then
	sort=request("sort")
	else
	sort=0
	end if
	select case sort
	case 0
	sqlorder=" order by supplier_Name"
	case 1
	sqlorder=" order by supplier_Name"
	case 2
	sqlorder=" order by supplier_Name desc"
	case 3
	sqlorder=" order by isVisible"
	case 4
	sqlorder=" order by isVisible desc"
	case 5
	sqlorder=" order by isAllowedSalesReport"
	case 6
	sqlorder=" order by isAllowedSalesReport desc"
	end select
	
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 30 Order By word_id"				
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
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
	function checkDelete()
	{

		<%
			If trim(lang_id) = "1" Then
				str_confirm = "?האם ברצונך למחוק את הספק"
			Else
				str_confirm = "Are you sure want to delete the group ?"
			End If   
		%>
		
		return window.confirm("<%=str_confirm%>");
		
	}
	
	function openSupplier(ID)
	{
		h = 1400;
		w = 800;
		S_Wind = window.open("addSupplier.asp?supplier_Id=" + ID, "S_Wind" ,"scrollbars=1,toolbar=0,top=0,left=300,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}	

//-->
</SCRIPT>

</HEAD>

<body>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>" ID="Table1">
<tr><td width=100% align="<%=align_var%>">
<!--#include file="../../logo_top.asp"-->
</td></tr>
<tr><td width=100% align="<%=align_var%>">
  <%numOftab = 4%>
  <%numOfLink = 13%>
<%topLevel2 = 87 'current bar ID in top submenu - added 03/10/2019%>
<!--#include file="../../top_in.asp"-->
</td></tr>
<tr><td class="page_title" align="<%=align_var%>" valign="middle" width="100%">&nbsp;</td></tr>
<tr>
<td align="<%=align_var%>" valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0 ID="Table2">
<tr>
<td width=100% valign=top>
<table width="90%" cellspacing="1" cellpadding="2" align="<%=align_var%>" border="0" bgcolor="#ffffff" ID="Table3">
<tr>
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id=word3 name=word3><!--מחיקה--><%=arrTitles(3)%></span>&nbsp;</td>	
	<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="word4" name=word4><!--עדכון-->עדכון</span>&nbsp;</td>	
	<td  align="<%=align_var%>" class="title_sort">&nbsp;מדינה בה נותן שירותים</td>
	<td align="<%=align_var%>" class="title_sort" nowrap>יצירת סיסמה</td>
	<td align="<%=align_var%>" class="title_sort" nowrap>סיסמה</td>
	
	<td align="<%=align_var%>" class="title_sort" nowrap></td>
	<td  align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="למיין בסדר יורד"><%elseif trim(sort)="4" then%>&nbsp;<a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="למיין בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=3"  title="פעיל/לא פעיל"><%end if%>&nbsp;פעיל/לא פעיל&nbsp;<img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a>

	<td  align="center" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="5" OR trim(sort)="6" then%>_act<%end if%>"><%if trim(sort)="5" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="למיין בסדר יורד"><%elseif trim(sort)="6" then%>&nbsp;<a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="למיין בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=5"  title="תוצג לספק/לא תוצג לספק"><%end if%>&nbsp;תוצג לספק<br> Sales Report&nbsp;<img src="../../images/arrow_<%if trim(sort)="5" then%>bot<%elseif trim(sort)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a>
	<td align="<%=align_var%>" class="title_sort" nowrap>שם משתמש</td>
	<td align="<%=align_var%>" class="title_sort" nowrap>תאריך כניסה אחרון</td>
	
	<td width="100%" align="<%=align_var%>" dir="<%=dir_obj_var%>" nowrap class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort+1%>" name=word25 title="למיין בסדר יורד"><%elseif trim(sort)="2" then%>&nbsp;<a class="title_sort" href="<%=urlSort%>&amp;sort=<%=sort-1%>"  title="למיין בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&amp;sort=1"  title="שם הספק"><%end if%>&nbsp;שם הספק&nbsp;<img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="1"></a></td>
	
</tr>
<%	'set rs_supplier = con.GetRecordSet("Select Country_Name,Suppliers.supplier_Id,supplier_Email1,supplier_Email2,supplier_Email3,supplier_Email4,GUID,dbo.[get_SuppliersLoginDate](supplier_Id) as LastLoginDate,dbo.[getSuppliersLoginDateString](Supplier_Id) as  LoginDateHistory,supplier_Name,LOGINNAME,Password,TermsPayment,supplier_Details1,supplier_Details2,supplier_Details3,supplier_Details4,supplier_Descr,isVisible,isAllowedSalesReport from Suppliers left join pegasus.dbo.Countries on Countries.Country_id=Suppliers.Country_id "  & sqlorder)
    set rs_supplier = con.GetRecordSet("Select Country_Name,Suppliers.supplier_Id,supplier_Email1,supplier_Email2,supplier_Email3,supplier_Email4,GUID,dbo.[get_SuppliersLoginDate](supplier_Id) as LastLoginDate,dbo.[getSuppliersLoginDateString](Supplier_Id) as  LoginDateHistory,supplier_Name,LOGINNAME,Password,TermsPayment,supplier_Details1,supplier_Details2,supplier_Details3,supplier_Details4,supplier_Descr,isVisible,isAllowedSalesReport from Suppliers left join " & pegasusDBName & ".dbo.Countries on Countries.Country_id=Suppliers.Country_id "  & sqlorder)
   if not rs_supplier.eof then 
		do while not rs_supplier.eof
    	supplier_Id = trim(rs_supplier("supplier_Id"))
    	supplier_Name = trim(rs_supplier("supplier_Name"))
    	Country_Name= trim(rs_supplier("Country_Name"))
    	LOGINNAME= trim(rs_supplier("LOGINNAME"))
    	Password= trim(rs_supplier("Password"))
    	isVisible= trim(rs_supplier("isVisible"))
    	isAllowedSalesReport=trim(rs_supplier("isAllowedSalesReport"))
    	LastLoginDate=trim(rs_supplier("LastLoginDate"))
    	LoginDateHistory=trim(rs_supplier("LoginDateHistory"))
    	GUID=trim(rs_supplier("GUID"))
    	supplier_Email1=trim(rs_supplier("supplier_Email1"))
    	supplier_Email2=trim(rs_supplier("supplier_Email2"))
    	supplier_Email3=trim(rs_supplier("supplier_Email3"))
    	supplier_Email4=trim(rs_supplier("supplier_Email4"))
    	'''response.Write "isVisible="& isVisible
%>
<tr>
	<td align="center" bgcolor="#e6e6e6" nowrap><a href="default.asp?deleteId=<%=supplier_Id%>" ONCLICK="return checkDelete(<%=countComp %>)"><IMG SRC="../../images/delete_icon.gif" BORDER=0></a></td>
	<td align="center" bgcolor="#e6e6e6" nowrap><a  href='#' OnClick="return openSupplier('<%=supplier_Id%>')" ><IMG SRC="../../images/edit_icon.gif" BORDER=0 alt="<%=alt_edit%>"></a></td>	
	<td  align="<%=align_var%>" dir=rtl bgcolor="#e6e6e6" nowrap><%=Country_Name%></td>
	<td  align="<%=align_var%>" dir=rtl bgcolor="#e6e6e6" nowrap>&nbsp;<%if trim(LOGINNAME)<>"" then%>
	<%if supplier_Email1="" and  supplier_Email2="" and supplier_Email3=""  and  supplier_Email4="" then%>
	<%else%>
	<a href="send_password.aspx?supplierId=<%=supplier_Id%>" onclick="return window.confirm('?האם ברצונך ליצור סיסמה חדשה  לספק')" target="_self" class="button_delete_1">יצירת סיסמה</a>&nbsp;<%end if%>
	<%end if%></td>
	<td align="<%=align_var%>" bgcolor="#e6e6e6"><%if len(Password)>0 then%>*******<%end if%>&nbsp;</td>	
    <!--<td align="<%=align_var%>" bgcolor="#e6e6e6" nowrap>&nbsp;<a class="button_edit_1" href="https://www.pegasusisrael.co.il/suppliers/?pp=<%=GUID%>&amp;crmadm=1" title="תצוגה דף אישי" target="_blank" style="width:auto" dir="rtl">תצוגה דף אישי&nbsp;<img src="../../images/selPerson.gif" border="0"></a>&nbsp;</td>-->
	<td align="<%=align_var%>" bgcolor="#e6e6e6" nowrap>&nbsp;<a class="button_edit_1" href="<%=Application("SiteUrl") %>suppliers/?pp=<%=GUID%>&amp;crmadm=1&uu=<%=UserId%>" title="תצוגה דף אישי" target="_blank" style="width:auto" dir="rtl">תצוגה דף אישי&nbsp;<img src="../../images/selPerson.gif" border="0"></a>&nbsp;</td>
    <Td align="center" bgcolor="#e6e6e6"><a href="vsbPress.asp?idsite=<%=supplier_Id%>"><%if isVisible = "False" then%><img src="../../images/lamp_off.gif" alt="פעיל" border="0" WIDTH="13" HEIGHT="18"><%else%><img src="../../images/lamp_on.gif" alt=" לא פעיל" border="0" WIDTH="13" HEIGHT="18"><%end if%></a></Td>
	<Td align="center" bgcolor="#e6e6e6"><a href="vsbPress.asp?isAllowed=<%=supplier_Id%>"><%if isAllowedSalesReport= "False" then%><img src="../../images/lamp_off.gif" alt="תוצג לספק " border="0" WIDTH="13" HEIGHT="18"><%else%><img src="../../images/lamp_on.gif" alt=" לא תוצג לספק " border="0" WIDTH="13" HEIGHT="18"><%end if%></a></Td>

	<td align="<%=align_var%>" bgcolor="#e6e6e6"><%=LOGINNAME%>&nbsp;</td>	
	<td align="center" bgcolor="#e6e6e6" nowrap><%If LastLoginDate<>"" then%> 
		<div class="tooltip"><%=LastLoginDate%> <%='FormatDateTime(LastLoginDate,0)%>
       <span class="tooltiptext"><%=Replace(LoginDateHistory,",","<BR>")%>
       <%
       
       
       %>
       </span>
</div>
		
	<%end if%></td>
	
	<td align="<%=align_var%>" bgcolor="#e6e6e6"><b><%=supplier_Name%></b>&nbsp;</td>	
</tr>
<%		rs_supplier.MoveNext
		loop
	end if
	set rs_supplier=nothing%>
</table>
</td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100% ID="Table4">
<tr><td align="<%=align_var%>" colspan=2 height="18" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='#' OnClick="return openSupplier('')"><span id=word6 name=word6><!--הוספת קבוצה-->הוספת ספק</span></a></td></tr>

</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table></td></tr>
</table>
</BODY>
</HTML>

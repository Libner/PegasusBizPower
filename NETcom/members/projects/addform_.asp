<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
<!--
function SelectDropDown(obj)
{
    oPopup.document.body.innerHTML = Select_Popup.innerHTML; 
    oPopup.document.charset = "windows-1255";
    oPopup.show(-(250-obj.offsetWidth), 21, 250, 210, obj);
    //window.alert(obj.offsetWidth);
}
function CheckDel(txt) {
  return (confirm(txt))    
}

function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}


<!--End-->
</script>  

</head>
<%
	sqlstr = "Select PRODUCT_ID FROM PRODUCTS WHERE PRODUCT_TYPE = 1 AND ORGANIZATION_ID=" & trim(OrgID) 
	set rs_pr = con.getRecordSet(sqlstr)
	If not rs_pr.eof Then
		prodId = trim(rs_pr(0))
	End If
	set rs_pr = nothing
	
	SubjId = request("SubjId")
	id=request("id")
	place=Request.QueryString("place")
	par_id = Request.QueryString("par_id")
	
	If trim(prodId) <> "" Then
	Set product=con.GetRecordSet("SELECT PRODUCT_NUMBER,Product_Name,Langu FROM Products where Product_Id="&prodId&" and ORGANIZATION_ID=" & trim(OrgID) )
	if not product.Eof then
		productName=product("Product_Name")
		productNumber=product("PRODUCT_NUMBER")
		if product("Langu") = "eng" then
			dir_align = "ltr"
			td_align = "left"
			pr_language = "eng"
		else
			dir_align = "rtl"
			td_align = "right"
			pr_language = "heb"
		end if
    end if 'not product.Eof
    Set product= nothing
    End If
	
PathCalImage = "../../"

'order_field in form  *******8
if Request.QueryString("down")<>nil then
newPlace=CInt(place)+1

	con.ExecuteQuery("update Form_Field set Field_Order=-10 WHERE Field_Order=" & newPlace  & " and ORGANIZATION_ID=" & trim(OrgID) )
	con.ExecuteQuery("update Form_Field set Field_Order=" & newPlace & " WHERE Field_Order=" & place  & " and ORGANIZATION_ID=" & trim(OrgID) )
	con.ExecuteQuery("update Form_Field set Field_Order=" & place & " WHERE Field_Order=-10 "  & " and ORGANIZATION_ID=" & trim(OrgID) )
	Response.Redirect "addform.asp?prodId="&prodId
end if

if Request.QueryString("up")<>nil then
newPlace=CInt(place)-1

	con.ExecuteQuery("update Form_Field set Field_Order=-10 WHERE Field_Order=" & newPlace  & " and ORGANIZATION_ID=" & trim(OrgID) )
	con.ExecuteQuery("update Form_Field set Field_Order=" & newPlace & " WHERE Field_Order=" & place  & " and ORGANIZATION_ID=" & trim(OrgID) )
	con.ExecuteQuery("update Form_Field set Field_Order=" & place & " WHERE Field_Order=-10 "  & " and ORGANIZATION_ID=" & trim(OrgID) )
	Response.Redirect "addform.asp?prodId="&prodId
end if
'end order field in form *********

if Request("delField")<>nil then
	con.ExecuteQuery("delete from Form_field where Field_Id="&Id)
	con.ExecuteQuery("Delete from Form_Select where  Field_Id="&Id)
	con.ExecuteQuery("Delete from Form_PROJECT_Value where  Field_Id="&Id)
end if

%>

<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 4%>
<%numOfLink = 2%>
<%If trim(wizard_id) = "" then%>
<!--#include file="../../top_in.asp"-->
<%End if%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<tr>   	
	<td class="page_title" width=100%>&nbsp;עריכת טופס <%=trim(Request.Cookies("bizpegasus")("Projectone"))	%>&nbsp;</td>
</tr>
<tr>
<td align=center colspan=3 valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0 ID="Table2">
<form name=form1 id=form1 target=_self>
<tr>
<td width=100% valign=top>
<table width="100%" align=center border="0" cellpadding="1" cellspacing="1"  bgcolor="#ffffff" >
<tr>
	<td width="50" nowrap class="title_sort" align=center>&nbsp;מחיקה&nbsp;</td>
	<td width="50" nowrap class="title_sort" align=center>&nbsp;עריכה&nbsp;</td>		
	<td width="100%" class="title_sort" align="center">&nbsp;&nbsp;&nbsp;&nbsp;</td>
    <td width="25" nowrap colspan=2 class="title_sort" align="right">&nbsp;</td>
</tr>
<!--#INCLUDE FILE="getfields.asp"-->
</form>
</table>
</td>
<td width=110 nowrap align=right valign=top class="td_menu">
<table cellpadding=1 cellspacing=0 width=100%>
<tr><td align="right" colspan=2 height="17" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='addfield.asp?prodId=<%=prodId%>'>הוספת שדה</a></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='#' onclick="window.open('check_form.asp?prodId=<%=prodId%>','Preview','left=20,top=20,tollbar=0,menubar=0,scrollbars=1,resizable=1,width=660,height=450');">תצוגה מקדימה</a></td></tr>
<tr><td align="right" colspan=2 height="10" nowrap></td></tr>
</table></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table>
</body>
<%set con=nothing%>
</html>

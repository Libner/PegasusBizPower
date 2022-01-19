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
function CheckDel(txt) {
  <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את השדה"
     Else
		str_confirm = "Are you sure want to delete the field?"
     End If   
  %>
  return (window.confirm("<%=str_confirm%>"))
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

'order_field in form  ********
if Request.QueryString("down")<>nil then
	newPlace=CInt(place)+1
	con.ExecuteQuery("update Form_Field set Field_Order=-10 WHERE Field_Order=" & newPlace  & " AND PRODUCT_ID = " & prodId & " and ORGANIZATION_ID=" & trim(OrgID) )
	con.ExecuteQuery("update Form_Field set Field_Order=" & newPlace & " WHERE Field_Order=" & place & " AND PRODUCT_ID = " & prodId  & " and ORGANIZATION_ID=" & trim(OrgID) )
	con.ExecuteQuery("update Form_Field set Field_Order=" & place & " WHERE Field_Order=-10 " & " AND PRODUCT_ID = " & prodId & " and ORGANIZATION_ID=" & trim(OrgID) )
	Response.Redirect "addform.asp?prodId="&prodId
end if

if Request.QueryString("up")<>nil then
	newPlace=CInt(place)-1
	con.ExecuteQuery("update Form_Field set Field_Order=-10 WHERE Field_Order=" & newPlace & " AND PRODUCT_ID = " & prodId & " and ORGANIZATION_ID=" & trim(OrgID) )
	con.ExecuteQuery("update Form_Field set Field_Order=" & newPlace & " WHERE Field_Order=" & place  & " AND PRODUCT_ID = " & prodId & " and ORGANIZATION_ID=" & trim(OrgID) )
	con.ExecuteQuery("update Form_Field set Field_Order=" & place & " WHERE Field_Order=-10 " & " AND PRODUCT_ID = " & prodId & " and ORGANIZATION_ID=" & trim(OrgID) )
	Response.Redirect "addform.asp?prodId="&prodId
end if
'end order field in form *********

if Request("delField")<>nil then
	con.ExecuteQuery("delete from Form_field where Field_Id="&Id)
	con.ExecuteQuery("Delete from Form_Select where  Field_Id="&Id)
	con.ExecuteQuery("Delete from Form_PROJECT_Value where  Field_Id="&Id)
end if

%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 21 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(trim(arr_title(0))) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing	
	  
	  sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	  set rsbuttons = con.getRecordSet(sqlstr)
	  If not rsbuttons.eof Then
		arrButtons = ";" & rsbuttons.getString(,,",",";")		
		arrButtons = Split(arrButtons, ";")
	  End If
	  set rsbuttons=nothing	  
%>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 4%>
<%numOfLink = 2%>
<%If trim(wizard_id) = "" then%>
<!--#include file="../../top_in.asp"-->
<%End if%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<tr>   	
	<td class="page_title" width=100%>&nbsp;&nbsp;</td>
</tr>
<tr>
<td align=center colspan=3 valign=top>
<table cellpadding=0 cellspacing=0 width=100% border=0>
<form name=form1 id=form1 target=_self>
<tr>
<td width=100% valign=top>
<table width="100%" align=center border="0" cellpadding="1" cellspacing="1"  bgcolor="#ffffff" >
<tr>
	<td width="50" nowrap class="title_sort" align=center>&nbsp;<span id=word3 name=word3><!--מחיקה--><%=arrTitles(3)%></span>&nbsp;</td>
	<td width="50" nowrap class="title_sort" align=center>&nbsp;<span id="word4" name=word4><!--עריכה--><%=arrTitles(4)%></span>&nbsp;</td>		
	<td width="100%" class="title_sort" align="center">&nbsp;&nbsp;&nbsp;&nbsp;</td>
    <td width="25" nowrap colspan=2 class="title_sort" align="<%=align_var%>">&nbsp;</td>
</tr>
<!--#INCLUDE FILE="getfields.asp"-->
</form>
</table>
</td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=0 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='addfield.asp?prodId=<%=prodId%>'><span id="word6" name=word6><!--הוספת שדה--><%=arrTitles(6)%></span></a></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='addquest.asp?prodId=<%=prodId%>'><span id="word7" name=word7><!--הוספת נושא--><%=arrTitles(7)%></span></a></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='#' onclick="window.open('check_form.asp?prodId=<%=prodId%>','Preview','left=20,top=20,tollbar=0,menubar=0,scrollbars=1,resizable=1,width=660,height=450');"><span id="word8" name=word8><!--תצוגה מקדימה--><%=arrTitles(8)%></span></a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
</table></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table>
</body>
<%set con=nothing%>
</html>

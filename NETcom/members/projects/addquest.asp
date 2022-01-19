<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
function CheckFields()
{
	if(document.formField.Title.value=='')
	{ 
		document.formField.Title.focus();
	<%
		If trim(lang_id) = "1" Then
			str_alert = "נא למלא את השדה כותר"
		Else
			str_alert = "Please, insert the subject title!"
		End If   
	%>  		
		window.alert("<%=str_alert%>");		return false;
	}
	return true;
}
//-->
</SCRIPT>

</head>
<%
 prodId=Request("prodId")
 ID = Request("ID")
 
  Set product=con.GetRecordSet("SELECT Product_Name,Langu FROM Products where Product_Id="&prodId)
  if not product.Eof then
		productName=product("Product_Name")
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

 %>
 <%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 23 Order By word_id"				
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
<!--#INCLUDE FILE="../../top_in.asp"-->
<%End if%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<tr><td class="page_title" width=100%><%If IsNumeric(wizard_id) Then%> <%=page_title_%> <%Else%>&nbsp;<%If trim(ID) = "" Then%><span id=word3 name=word3><!--הוסף--><%=arrTitles(3)%></span><%Else%><span id="word4" name=word4><!--עדכן--><%=arrTitles(4)%></span><%End If%>&nbsp;<span id="word5" name=word5><!--נושא--><%=arrTitles(5)%></span>&nbsp;<%End If%></td></tr>

<tr><td width=100% align="<%=align_var%>">
<% wizard_page_id = 1 %>
<!--#include file="../../wizard_inc.asp"-->
</td></tr>
<tr><td align=center colspan=3>
<table cellpadding=1 cellspacing=1 width=100% border=0>
<tr><td width=100% bgcolor="#E6E6E6">
<table cellpadding=0 cellspacing=0 width=80% border=0 align=center>
<form method="post" action="addfield.asp" id=formField name=formField>
<% 
if ID<>nil then
   set p_field=con.GetRecordSet("select FIELD_TITLE,FIELD_TYPE from FORM_FIELD where FIELD_ID ="& ID)
   if not p_field.eof then
	FIELD_TITLE = p_field("FIELD_TITLE")
	FIELD_TYPE = p_field("FIELD_TYPE")    
   end if 
 else
	FIELD_TYPE = 10
 end if
	
%>
	<tr>
		<td align="<%=align_var%>"><TEXTAREA dir="<%=dir_align%>" class="txt" rows=2 style="width:400" name="Title"><%=FIELD_TITLE%></textarea></td>
		<th align="center" class="title_Show_form" valign=top><span id=word6 name=word6><!--כותר לנושא--><%=arrTitles(6)%></span>&nbsp;</th>
	</tr>	
	<tr>
	<td colspan="2" height=40 valign="bottom">
	<table cellspacing=0 cellpadding=0 width=70% align=center>
	<tr>
	<td align="center" width=50% nowrap>
	<INPUT class="but_menu" style="width:90" type="button" value="<%=arrButtons(2)%>" id=button2 name=button2 onclick="document.location.href='addform.asp?prodId=<%=prodid%><%if ID <> nil then%>#link<%=ID%><%end if%>'"></td>
	<td align="center" width=50% nowrap>
		<input type=submit style="width:90" value="<%=arrButtons(1)%>" class="but_menu" onClick="return CheckFields()" id=button1 name=button1>
			<input type="hidden" name="prodId" value="<%=prodId%>">			
			<input type="hidden" name="ID" value="<%=ID%>">
			<INPUT type="hidden" id=type_field name=type_field value="<%=FIELD_TYPE%>">			
			<input type="hidden" name="new_new" value="true" ID="new_new">
			<input type="hidden" name="update" value="true" ID="update">
		</td>
	</tr>
	</form>
 </table>
 </td></tr></table>
 </td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100% ID="Table3">
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='addfield.asp?prodId=<%=prodId%>'><span id=word7 name=word7><!--הוספת שדה--><%=arrTitles(7)%></a></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='addform.asp?prodId=<%=prodId%>'><span id="word8" name=word8><!--עריכת טופס--><%=arrTitles(8)%></a></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='#' onclick="window.open('check_form.asp?prodId=<%=prodId%>','Preview','left=20,top=20,tollbar=0,menubar=0,scrollbars=1,resizable=0,width=660');"><span id="word9" name=word9><!--תצוגה מקדימה--><%=arrTitles(9)%></span></a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table>
</body>
<%set con=nothing%>
</html>





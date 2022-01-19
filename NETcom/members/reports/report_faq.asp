<%
	quest_id = Request.Form("quest_id")
%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="javascript">
<!--
function report_open(fid){	
	form1.FieldId.value = fid;
	form1.submit();
	return false;
}
//-->
</script>
</head>

<body marginwidth="0" marginheight="0" hspace="0" vspace="0" topmargin="0" leftmargin="0" rightmargin="0">
<div align="center">
<center>
<!--#INCLUDE FILE="logo_top.asp"-->
<table width=760 bgcolor="#CCCCCC" align="center" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="20" style="color:#000000; font-size:13pt" align=right valign=bottom><b>דוח נרשמים לשדה חופשי</b>&nbsp;&nbsp;&nbsp;</td>
	</tr>
	<tr>
		<td width="100%" bgcolor="#999999" height="1"></td>
	</tr>					
</table>
</center></div>
<table border="0" width="760" align=center cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top" align="center" width="100%"><div align="center"><center>
	<table border="0" width="100%" cellspacing="0" cellpadding="0">	       
      <tr><td height=10 nowrap></td></tr> 
      <tr>
        <td width="100%" align="center">
<!-- start code -->
			<table width="710" BGCOLOR="#999999" ALIGN="center" BORDER="0" CELLPADDING="1" cellspacing="0">
			<tr>
			<td align="right" width="100%" valign="top"> 
			<%	
			sqlStr = "Select * from products where product_id=" & quest_id
			'Response.Write sqlStr
			'Response.End
			set prod = con.GetRecordSet(sqlStr)
			if not prod.eof then
				productName=prod("Product_Name")						
				if prod("Langu") = "eng" then
					td_align = "left"
				else
					td_align = "right"
				end if
			end if
			set prod = nothing%> 
			<table border="0" cellpadding="2" cellspacing="0" width="100%" align="right" bgcolor="#F0F0F0">
			<tr><td class="form" align=center width=100% bgcolor=#999999 colspan=4><INPUT type="hidden" id=product_id name=product_id value="<%=quest_id%>">
					<font color=#FFFFFF size=3><%=productName%></font>
				</td>
			</tr>
	
	<tr><td width="100%" height=10></td></tr>
<%	set f=con.GetRecordSet("SELECT Field_Id,FIELD_TITLE FROM FORM_FIELD Where product_id = " & quest_id &_
    " and (Field_Type =1 or Field_Type=2 or Field_Type=7) Order by Field_Order")
	if not f.eof then
		isanswers = true%>
	
	<tr><td width="100%" height=5></td></tr>
	   <tr>
		<td>
			<TABLE align=center BORDER=0 CELLSPACING=1 CELLPADDING=1 width=100%>
    <%DO WHILE NOT f.EOF
	FIELD_TITLE=trim(f("FIELD_TITLE"))
    Field_Id=f("Field_Id")
    if FIELD_TITLE<>"" then%>
			<tr>
				<tD width="100%" align=<%=td_align%> valign="top"  bgcolor=#DADADA style="padding-right:3px;" dir="rtl">&nbsp;&nbsp;<a class=linkFaq href='' onclick='return report_open(<%=Field_Id%>)'><%=FIELD_TITLE%></a></tD>
			</tr>
    <%else%>
			<tr>
				<tD width="100%" align=<%=td_align%> valign="top"  bgcolor=#DADADA nowrap>&nbsp;&nbsp;<a class=linkFaq href='' onclick='return report_open(<%=Field_Id%>)'>הבושת</a>&nbsp;&nbsp;</tD>
			</tr>
    
    <%end if    %>
   
    <%
	f.MOVENEXT
    LOOP%>
			</TABLE>		
	 	</td>
	</tr>
	

<%set f=nothing
	
	end if 'if not eof 
%>	
				
<!-- end code --> 
	<tr><td width="100%" height=20 colspan=2 align=center valign=middle>
	<%if isanswers = false then%>
		<font size="3" color="#ff0000">אין בטופס הזה שדות חופשיים</font><br><br>
	<%end if%>
	<FORM action="report_text.asp" method=POST id=form1 name=form1 target=_self>	<INPUT type="hidden" id=quest_id name=quest_id value="<%=quest_id%>">			<INPUT type="hidden" id=QId name=QId value="">	<INPUT type="hidden" id=FieldId name=FieldId value="">
	</FORM>
	</td></tr>		
</TABLE>
</body>
<%
set con = nothing%>
</html>

<%SERVER.ScriptTimeout=3000%>
<!--#include file="../../NETcom/connect.asp"-->
<!--#include file="../../NETcom/reverse.asp"-->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
</head>
<body>
<% prod_id = Request.Form("selFormId")
	set prod = con.GetRecordSet("Select * from products where product_id=" & prod_id & " and USER_ID=" & trim(Request.Cookies("biznetcom")("UserID")) )
		if not prod.eof then
		productName=prod("Product_Name")
		PRODUCT_DESCRIPTION = trim(prod("PRODUCT_DESCRIPTION"))
		if prod("Langu") = "eng" then
			dir_align = "ltr"
			td_align = "left"
			Langu = "eng"
		else
			dir_align = "rtl"
			td_align = "right"
			Langu = "heb"
		end if
	end if
	set prod = nothing
	htmlForm = "<div name=tofes><table border=0 cellpadding=2 cellspacing=1 width='100%'>"
	if PRODUCT_DESCRIPTION <> "" then
		htmlForm = htmlForm & "<tr><td align=" & td_align & " width='100%' colspan=2>"
		htmlForm = htmlForm & "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=3><TR><td class=form width=100% align=" & td_align & "><span dir=" & dir_align & " >" & breaks(PRODUCT_DESCRIPTION) & "</span></td></tr></TABLE>"
        htmlForm = htmlForm & "</td></tr>"
	end if
	
	' start fields dynamics -->
	
	set fields=con.GetRecordSet("SELECT * FROM Form_View Where product_id=" & prod_id &" Order by Field_Order")
				
	do while not fields.EOF 
		Field_Id = fields("Field_Id")
		Field_Title = trim(fields("Field_Title"))
		Field_Type = fields("Field_Type")
		Field_Size = fields("Field_Size")
		Field_Align = trim(fields("Field_Align"))
		Field_Scale = fields("Field_Scale")
		
		if Langu = "eng" then
			htmlForm = htmlForm & "<tr><td valign=top bgcolor=transparent  style='padding-left:3px;padding-right:15px'  class='form'  align=left nowrap>&nbsp;"
			if cstr(Field_Type) = "5" then
				htmlForm = htmlForm & GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng","","" & Field_Title & "")
			end if
			htmlForm = htmlForm & Field_Title & "&nbsp;</td>"
			htmlForm = htmlForm & "<td valign=top align=left bgcolor=transparent width='100%'>"
			if cstr(Field_Type) <> "5" then
				htmlForm = htmlForm & GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"eng","","" & Field_Title & "")
			end if
			htmlForm = htmlForm & "</td></tr>"
		else ' Langu = "eng"
			htmlForm = htmlForm & "<tr><td valign=top align=right bgcolor=transparent width='100%'>"
			if cstr(Field_Type) <> "5" then
				htmlForm = htmlForm & GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","","" & Field_Title & "")
			end if
			htmlForm = htmlForm & "</td><td valign=top bgcolor=transparent class='form' align=right style='padding-right:3px;padding-left:15px' nowrap><span dir=rtl>" & Field_Title & "</span>"
			if cstr(Field_Type) = "5" then
				htmlForm = htmlForm & GetFormField(Field_Id,Field_Type,Field_Size,Field_Align,Field_Scale,PathCalImage,"heb","","" & Field_Title & "")
			end if
			htmlForm = htmlForm & "</td></tr>"
		end if ' Langu = "eng"
	fields.moveNext()
	loop
	set fields=nothing 	
	' end fields dynamics -->
	htmlForm = htmlForm & "<tr><td height=10 width='100%' colspan=2><table width=100% align=center cellpadding=0 cellspacing=0><tr><td width=100% bgcolor=#808080 height=1><img src='../../NETcom/images/line.gif' WIDTH=5 HEIGHT=1></td></tr></table></td></tr>"
	htmlForm = htmlForm & "<tr><td align=center colspan=2><IMG "
	if Langu = "eng" then
		htmlForm = htmlForm & "SRC='http://www.cyberserve.co.il/NETcom/images/SEND-ENG.gif'"
	else
		htmlForm = htmlForm & "SRC='http://www.cyberserve.co.il/NETcom/images/b-send-a.gif'"
	end if
	htmlForm = htmlForm & " border=0 hspace=60></td></tr>"
	htmlForm = htmlForm & "</table></div><div name=endtofes></div>"
	%> 
	<SCRIPT LANGUAGE=javascript>
	<!--
		window.parent.main_frame.txtForm.value = "<%=htmlForm%>";
		window.parent.main_frame.btnOKClick();
	//-->
	</SCRIPT>
<%set con = nothing%>	
</body>
</html>

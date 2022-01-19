<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="..\..\netcom/reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title>Bizpower Administration</title>
<meta charset="windows-1255">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
</head>
<%
	OrgID = Request.QueryString("ORGANIZATION_ID")		
	if trim(OrgID) <> "" then
		sqlStr = "select ORGANIZATION_NAME from  ORGANIZATIONS where ORGANIZATION_ID = " & OrgID
		set rs_ORGANIZATIONS = con.GetRecordSet(sqlStr)
		if not rs_ORGANIZATIONS.eof then
			if trim(rs_ORGANIZATIONS("ORGANIZATION_NAME")) <> "" then
				ORGANIZATION_NAME = rs_ORGANIZATIONS("ORGANIZATION_NAME")
			end if
		end if
		set rs_ORGANIZATIONS = nothing
	end if	
%>
<body bgcolor="#FFFFFF" onload="window.focus();">
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td class="page_title"><%=ORGANIZATION_NAME%>
    <br>
    רשימת קבצים המצורפים לטפסים
    </td>
  </tr> 
</table>
<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1">
<tr bgcolor="#B4B4B4">   
   <td width="15%" align="center" class="11normalB">גודל קובץ</td>
   <td width="35%" align="center" class="11normalB">שם קובץ</td>
   <td width="45%" align="center" class="11normalB">שם טופס</td>
   <td width="5%" align="center" class="11normalB">מספר</td>
</tr>
<%
	set qs=con.GetRecordSet("SELECT PRODUCT_NAME,DATE_START,FILE_ATTACHMENT,ATTACHMENT_TITLE,ATTACHMENT_SIZE FROM PRODUCTS WHERE ORGANIZATION_ID=" & trim(OrgID) & " AND Len(FILE_ATTACHMENT)>0 AND PRODUCT_NUMBER = '0'" )
	If not qs.eof Then
	count_size = 0 : i=0
	While not qs.eof
	    i=i+1
		prodName = qs("product_name")			
		prodDate = qs("DATE_START")
		attachment = qs("FILE_ATTACHMENT")
		attachment_title = trim(qs("ATTACHMENT_TITLE"))
		file_size = 0
		file_size = trim(qs("ATTACHMENT_SIZE"))		
		If IsNumeric(trim(file_size)) Then
			count_size = count_size + file_size
			file_size = Round(file_size / 1024,2)
		End If
%>
	<tr>
		<td align="center" valign="top" bgcolor="#DDDDDD"><%=file_size%> kb</td>
		<td align="center" valign="top" bgcolor="#DDDDDD"><a href="../../../download/products/<%=attachment%>" target=_blank><%=attachment%></a></td>
	    <td align="center" valign="top" bgcolor="#DDDDDD"><font size="2"><strong><%=prodName%></strong></font></td>
	    <td align="center" valign="top" bgcolor="#DDDDDD"><%=i%></td>
	</tr>
<%
	qs.movenext
	Wend
	count_size = FormatNumber(count_size / 1024 / 1024, 2, -1)
%>
	<tr><td colspan=4 height=1 bgcolor=#808080></td></tr>
	<tr>
		<td align="center" valign="top" bgcolor="#C0C0C0" CLASS=11normalB><%=count_size%> mb</td>
		<td align="center" valign="top" bgcolor="#C0C0C0" CLASS=11normalB>סה"כ גודל קבצים מצורפים</td>
	    <td align="center" valign="top"></td>
	</tr>	
<%	
	Else
%>
	<tr>
		<td align="center" valign="top" bgcolor="#DDDDDD" colspan=4>לא נמצאו קבצים המצורפים לטפסים</td>
	</tr>	
<%	
	End If
	set qs = nothing
%>
</table>
</body>
</html>
<%
set con = nothing
%>
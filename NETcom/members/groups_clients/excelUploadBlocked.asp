<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<%  groupId = trim(Request.QueryString("groupId"))
    UserId = trim(Request.Cookies("bizpegasus")("UserID"))
    OrgID = trim(Request.Cookies("bizpegasus")("OrgID"))
 %>   
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language=javascript>
<!--
	function ifFieldEmpty(){
	  if (document.form1.UploadFile.value=='')
		{
		alert('��� ����');
		document.form1.UploadFile.focus();		
		return false;
		}	  
		return true;		
}
//-->
</script>
</head>
<body style="margin:0px;background-color:#E5E5E5" onload="window.focus();">
<table cellpadding=0 cellspacing=0 width=100%>
<tr><td width="100%" class="page_title" dir=rtl>&nbsp;����� ������&nbsp;(Excel)</td></tr>         
<tr><td width="100%" colspan=2 height=10></td></tr>
<tr><td width="100%" colspan=2>
<table border="0" width="95%" align=center cellspacing="1" cellpadding="5">
<tr><td width=100%>
<form name="form1" id="form1" target="_self" ACTION="blockedFileadd.asp?C=1" METHOD="post" ENCTYPE="multipart/form-data">
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
<td width=100% align=right>
<select name="groupID" id="groupID" style="width:200" width="150" dir="rtl" class="norm">
<option value="">��� ��� �������</option>
<%
	sqlStr = "select groupId, groupName from groups where ORGANIZATION_ID=" & OrgID & " ORDER BY rtrim(ltrim(groupName))"
	set rs_groups = con.GetRecordSet(sqlStr)
	do while not rs_groups.eof
%>
<option value="<%=trim(rs_groups(0))%>" <%If trim(groupId) = trim(rs_groups(0)) Then%> selected <%End if%>><%=trim(rs_groups(1))%></option>
<%
	rs_groups.movenext		
	loop
	set rs_groups = nothing
%>
</select>
</td>
<td width=220 nowrap align=right>��� ����� ���� ������ ����� �� �������</td>
<td width=20 nowrap align=center><b>1</b></td></tr>
</table>
</td></tr>
<tr><td width=100%>
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
<td width=100% align=right><INPUT type=button class="but_menu" onclick="window.open('../../../download/list.xls')" value="����" style="width:70"></td>
<td width=220 nowrap align=right>Excel ���� ����� ������ ������</td>
<td width=20 nowrap align=center><b>2</b></td></tr>
</table>
</td></tr>
<tr><td width=100%>
<table cellpadding=0 cellspacing=0 width=100%">
<tr>
<td width=100% align=right>.����� ���� ����� ������ ��� ,Excel���� �� ����� ������� ������ �</td>
<td width=20 nowrap align=center><b>3</b></td></tr>
</table>
</td></tr>
<tr><td width=100%>
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
<td width=100% align=right>.������ ��� Excel���� �� ����� � Browse ��� �� ����</td>
<td width=20 nowrap align=center><b>4</b></td></tr>
</table>
</td></tr>
<tr>
	<td align=right>
	<table border=0 width=100% cellpadding=0 cellspacing=0>		
		<tr height=25>																							
			<td align="right" width=100%>&nbsp;<INPUT style="font-size:11pt" TYPE="FILE" NAME="UploadFile" style="width:300" ID="File1"></td>														
		</tr>		
	</table>					
	</td>
</tr>
<tr><td width=100%>
<table cellpadding=0 cellspacing=0 width=100%>
<tr>
<td width=100% align=right><INPUT type="submit" class="but_menu" value="��� ������" id="upload_excel" name="upload_excel" style="width:100" onClick="return ifFieldEmpty();"></td>
<td width=150 nowrap align=right>��� �� ���� ��� ������</td>
<td width=20 nowrap align=center><b>5</b></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>	
</table></form></td>
</tr></table>
</body>		
</html>		
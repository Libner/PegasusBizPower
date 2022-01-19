
<html>
<head>
<title>test</title>
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="javascript">

</script>
</head>
<%
response.write "<br>hiddenfield=" &request.form("hiddenfield")
response.write "<br>contactID=" &request.form("contactID")
response.write "<br>textfield=" &request.form("textfield")
%>
<body>
<table border="0" width="100%" align=center cellspacing="0" cellpadding="0" ID="Table1">
		  <form action="testPost.asp" id="form1" name="form1" method="post" ENCTYPE="multipart/form-data">			
			<tr>
				<td  style="font-size:16pt" align=center width=100% >
				<INPUT type="hidden" id="hiddenfield" name="hiddenfield" value="hiddenvalue">
				<input type="hidden" name="contactID" id="contactID" value="12345">
				<INPUT type="text" id="textfield" name="textfield" value="textvalue">
				</td>
			</tr>
			<tr>
				<td  style="font-size:16pt" align=center width=100% >
				<INPUT type="submit" id="sub" name="sub" value="submit">
				</td>
			</tr>
			</form>
			</table>
</body>
</html>

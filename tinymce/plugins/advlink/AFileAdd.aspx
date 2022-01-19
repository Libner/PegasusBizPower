<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="AFileAdd.aspx.vb" Inherits="bizpower_pegasus2018.AFileAdd" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>File Browser</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../themes/advanced/css/editor_popup.css" rel="stylesheet" type="text/css">
	</HEAD>
	<body>
		<form id="form1" runat="server" method="post" ENCTYPE="multipart/form-data">
			<p>File Upload</p>
			<p><input type="file" runat="server" style="WIDTH:250px" id="UploadFile1" name="UploadFile1"></p>
			<p>
				<asp:RegularExpressionValidator id="valPicFile" runat="server" CssClass="validator" ErrorMessage="הקובץ אשר בחרת אינו חוקי"
					Display="Dynamic" ControlToValidate="UploadFile1" ValidationExpression=".+\.([dD][oO][cC]|[rR][tT][fF]|[tT][xX][tT]|[xX][lL][sS]|[pP][pP][tT]|[pP][dD][fF]|[hH][tT][mM]|[jJ][pP][gG]|[gG][iI][fF]|[pP][nN][gG]|[dD][oO][cC][xX]|[xX][lL][sS][xX]|[xX][lL][sS][mM]|[pP][pP][tT][xX]|[dD][oO][cC][mM])$"></asp:RegularExpressionValidator>
			</p>
			<p><input type="submit" value="Save"></p>
		</form>
	</body>
</HTML>

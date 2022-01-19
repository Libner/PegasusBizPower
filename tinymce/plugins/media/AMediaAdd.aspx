<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="AMediaAdd.aspx.vb" Inherits="bizpower_pegasus2018.AMediaAdd" %>
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
			<p>Media Upload</p>
			<p><input type="file" runat="server" style="WIDTH:250px" id="UploadFile1" name="UploadFile1"></p>
			<p>
				<asp:RegularExpressionValidator id="Regularexpressionvalidator1" runat="server" CssClass="validator" ErrorMessage="אנא בחר קובץ מדיה חוקי"
					Display="Dynamic" ControlToValidate="UploadFile1" validationexpression="^([a-zA-Z].*|[1-9].*)\.(((w|W)(m|M)(v|V))|((s|S)(w|W)(f|F))|((m|M)(i|I)(d|D))|((m|M)(p|P)(e|E)(g|G))|((a|A)(v|V)(i|I))|((c|C)(d|D)(a|A))|((w|W)(a|A)(v|V))|((w|W)(m|M)(a|A))|((m|M)(p|P)(3)))$"></asp:RegularExpressionValidator>
			</p>
			<p><input type="submit" value="Save"></p>
		</form>
	</body>
</HTML>

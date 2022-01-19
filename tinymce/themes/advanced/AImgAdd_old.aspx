<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="AImgAdd.aspx.vb" Inherits="Newpan.AImgAdd" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>File Browser</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">		
		<link href="css/editor_popup.css" rel="stylesheet" type="text/css" />		
	</HEAD>
	<body>
	<form id="form1" runat="server" method="post" ENCTYPE="multipart/form-data">
	<p>Image Upload</p>
	<p><input type="file" runat="server" style="width:250" id="UploadFile1" name="UploadFile1"></p>
	<p>
	<asp:RegularExpressionValidator id="valPicFile" runat="server" CssClass="validator" ErrorMessage="אנא בחר קובץ תמונה חוקי"
	Display="Dynamic" ControlToValidate="UploadFile1" 
	ValidationExpression="^.+\.(([jJ][pP][eE]?[gG])|([gG][iI][fF])|([pP][nN][gG]))$"></asp:RegularExpressionValidator>					
	</p>
	<p><input type="submit" value="Save"></p>
	</form>
	</body>
</html>	

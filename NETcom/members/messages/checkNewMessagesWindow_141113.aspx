<%@ Page Language="vb" AutoEventWireup="false" EnableViewState="False" Codebehind="checkNewMessagesWindow.aspx.vb" Inherits="bizpower_pegasus.checkNewMessagesWindow"%>
<%@ OutputCache Duration="1" VaryByParam="none" Location="none"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>checkNewMessagesWindow</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
	<link href="style.css?v=<%=Application("sversion")%>" type="text/css" rel="STYLESHEET">         
		<script type="text/javascript" src="../../../javaScript/jquery-1.3.2.min.js"></script> 
<script language="javascript">
<!--	
	var oMultiChatInterval = "";
	function Refresh_Function()
	{
//alert("Refresh")
			$(document).ready(function() {
				
				if (idRef == -1){
					$.post("ajaxNewMessages.aspx", function(data){
										$("#mesagesDiv").html(data);
					});					
				}
				else
				{
					$.post("ajaxNewMessages.aspx?idRef=" + idRef, function(data){
											$("#mesagesDiv").html(data);
					});					
				}	

				if (idRef != 5)
					idRef = idRef+1;	
				else
					idRef = 0	
						
			});			
	}			
			
	var idRef = -1;			
	Refresh_Function();
	oMultiChatInterval = window.setInterval("Refresh_Function()",120000);		
//120000 milisec=2min
//-->
</script>	
  </head>
	<body bgcolor="#F4F4F4" leftmargin="0" topmargin="0" dir="rtl" rightmargin="0" bottommargin="0">
    <form id="Form1" method="post" runat="server">		
			<div id="mesagesDiv" name="mesagesDiv" style="border:0px solid red; width:178px; height:20px; overflow: auto;"></div>
    </form>
  </body>
</html>
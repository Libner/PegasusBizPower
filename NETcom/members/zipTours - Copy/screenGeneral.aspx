<%@ Page Language="vb" AutoEventWireup="false" Codebehind="screenGeneral.aspx.vb" Inherits="bizpower_pegasus.screenGeneral"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>screenGeneral</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../IE4.css" rel="STYLESHEET" type="text/css">

  </head>
  <body MS_POSITIONING="GridLayout">

    <form id="Form1" method="post" runat="server">
   <table dir="rtl" border="0" cellpadding="0" cellspacing="1" width="100%" ID="Table2">
      <tr>
    <td align="center" valign=top   width=100% height=10>   </td></tr>
    <tr>
    <td align="center" valign=top   width=100%>   
      <table dir="rtl" border="0" cellpadding="0" cellspacing="1" ID="Table2">
    <tr>
    <td align="right" width="100%" valign=top >   
    	<asp:DataList ID="dtlGeneral" Runat="server" RepeatColumns="1" RepeatDirection="Vertical">
										<ItemTemplate>
											<table border=0 cellpadding=2 cellspacing=2>
                                            <tr><td><input type="checkbox" ID="chkSpec" runat="server"></td>
                                            <td style="font-weight:bold;font-family:Arial;font-size:12pt"> <%#Container.DataItem("Column_Name")%></td>

                                        
                                            </tr>
                                            
                                          
                                            </table>
										</ItemTemplate>
									</asp:DataList>
									</td></tr>
									  <tr><td height="20" valign="middle" colspan="2" align="center" class="td_admin_4"><asp:button  ID=btnSave EnableViewState="false" runat="server"  Text="�����"></asp:Button> </td></tr>


									
									</table>		</td></tr></table>
    </form>

  </body>
</html>

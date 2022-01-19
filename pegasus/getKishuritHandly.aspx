<%@ Page Language="vb" AutoEventWireup="false" Codebehind="getKishuritHandly.aspx.vb" Inherits="bizpower_pegasus2018.getKishuritHandly" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML lang="en">
	<HEAD>
		<title>
			<%=fTitle%>
		</title>
		<meta charset="utf-8">
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
			<link rel="stylesheet" href="/resources/demos/style.css">
				<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
				<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
				
		<script src="../netCom/include/CalendarPopup.js" charset="windows-1255"></script>
				</script>
				<link href="../netCom/IE4.css" rel="STYLESHEET" type="text/css">
	</HEAD>
	<body style="MARGIN:0px" onload="self.focus();">
		<center>
			<form id="Form1" method="post" runat="server">
				<table width="100%" align="center" cellspacing="0" cellpadding="2" border="0">
					<tr>
						<td align="center" class="td_admin_4" height="10"></td>
					</tr>
					<tr>
						<Td  align=center nowrap valign="bottom"><a href="" onclick="cal1xx.select(document.getElementById('fromDate'),'AfromDate','dd/MM/yy'); return false;"
										id="AfromDate"><img src="../netCom/images/calendar.gif" border="0" align="center"></a>
										<input runat="server" type="text" id="fromDate" 
										NAME="fromDate" readonly>
										</td>
									</tr>
						<tr>
						<Td  align=center nowrap valign="bottom">			
									<a href="" onclick="cal1xx.select(document.getElementById('toDate'),'AtoDate','dd/MM/yy'); return false;"
										id="AtoDate"><img src="../netCom/images/calendar.gif"  border="0" align="center"></a>
										<input runat="server" type="text" id="toDate"  
										NAME="toDate" readonly></Td>
					</tr>
					<tr>
						<td class="td_admin_5" align="center" valign="bottom">
						<asp:LinkButton Runat="server" ID="btnSearch" CssClass="button_edit_1" style="width:70px">&nbsp;ωμη&nbsp;</asp:LinkButton>
								
						</td>
					</tr>
					<tr>
						<td align="center" class="td_admin_4" height="10"></td>
					</tr>
					<tr>
						<td align="center" class="td_admin_4" ><asp:Literal ID="ltMess" Runat=server></asp:Literal></td>
					</tr>
				</table>
			</form>
			<SCRIPT type="text/javascript">
            <!--
            var cal1xx = new CalendarPopup('CalendarDiv');
                cal1xx.setYearSelectStartOffset(120);
	cal1xx.setYearSelectEndOffset(2);
	//cal1xx.setYearSelectStart(1910); //Added by Mila
                cal1xx.showNavigationDropdowns();
                //cal1xx.offsetX = -50;
                //cal1xx.offsetY = -40;
          
                //-->
			</SCRIPT>
			<DIV ID='CalendarDiv' STYLE='POSITION:absolute;VISIBILITY:hidden;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>

		</center>
	</body>
</HTML>

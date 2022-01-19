<%@ LANGUAGE=VBScript %>
<%response.Write ("edit="& request.Cookies("bizpegasus")("salesControl_Points_Edit"))%>
<%Response.Buffer = True%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
	<head>
		<!-- #include file="../../title_meta_inc.asp" -->
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<style>
html,body {
    height:100%;
}
.h_iframe iframe {
    position:relative;
    top:0px;
    left:0;
    width:100%;
    height:90%;
}
		</style>
		<script type="text/javascript">
  function resizeIframe(iframe) {
   iframe.height = iframe.contentWindow.document.body.scrollHeight + "px";
   /*size=iframe.contentWindow.document.body.scrollHeight;
    var size_new=size-510; 
    iframe.height=size_new + "px";*/
 
  }
		</script>
	</head>
	<script>
	
function openReportExcel()
{
	
			window.open("Report_VouchersProviderStatus.aspx","ReportExcel","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

}
		


	
	</script>
	<body>
		<table border="0" width="100%" style="height:100%" cellspacing="0" cellpadding="0" ID="Table1">
			<tr>
				<td width="100%" align="<%=align_var%>">
					<!--#include file="../../logo_top.asp"-->
				</td>
			</tr>
			<tr>
				<td width="100%" align="<%=align_var%>">
					<%numOftab = 108%>
					<%numOfLink =1%>
					<!--#include file="../../top_in.asp"-->
				</td>
			</tr>
			<tr>
				<td class="page_title" align="center" width=100%>
					<table border="0" cellpadding="2" cellspacing="0" align="center" width=100% ID="Table2">
						<tr>
						<td>&nbsp;</td>
						</tr>
					</table> <!--אנא הקלד את שם החברה מימין או את שם איש הקשר לאיתור במאגר הלקוחות--><%'=arrTitles(25)%></td>
			</tr>
			<tr>
				<td width="100%" valign="top" align="center" style="height:100%">
					<table dir="<%=dir_var%>" border="0" cellpadding="0" cellspacing="1" width="100%" ID="Table3" style="height:100%">
						<tr>
							<td align="left" width="100%" valign="top" style="height:100%">
								<div style="position:relative;height:100%" >
									<div class="h_iframe" style="height:100%">
										<iframe name="frameScrreen" id="frameScrreen" src="targetScreen.aspx?tab=<%=request.querystring("tab")%>&PEdit=<%=request.Cookies("bizpegasus")("salesControl_Points_Edit")%>"	ALLOWTRANSPARENCY=true  marginwidth="0" marginheight="0" hspace="0" vspace="0"  frameborder="0">
										</iframe>
									</div>
								</div>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</body>
</html>
<%set con=Nothing%>

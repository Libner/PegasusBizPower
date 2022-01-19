<%@ LANGUAGE=VBScript %>
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
					<%numOftab = 82%>
					<%numOfLink =0%>
					<!--#include file="../../top_in.asp"-->
				</td>
			</tr>
			<tr>
				<td class="page_title" align="center">
					<table border="0" cellpadding="2" cellspacing="0" align="center">
						<tr>
						<td><a href="javascript:void(0)" onclick="return openReportExcel();" class="button_edit_1" style="width:290;">דוח רשימת חשבוניות שטרם בוצעה עבורם התאמה</a></td>
						<td width="10"></td>
							<td><a href="default.asp?tab=8" class="button_edit_1" style="width:105;">שובר ללא התאמה</a></td>
							<td width="10"></td>
							<td><a href="default.asp?tab=7" class="button_edit_1" style="width:105;">טיולי החודש</a></td>
							<td width="10"></td>
							<td><a href="default.asp?tab=10" class="button_edit_1" style="width:115;">רשימת בריפים 
									קדימה</a></td>
							<td width="10"></td>
							<td><a href="default.asp?tab=5" class="button_edit_1" style="width:135;">הטיולים שיצאו 
									ואין תמחור</a></td>
							<td width="10"></td>
							<td><a href="default.asp?tab=4" class="button_edit_1" style="width:125;" nowrap>לא תואם 
									בריף החזרה</a></td>
							<td width="10"></td>
							<%if false then%>
							<td><a href="default.asp?tab=10" class="button_edit_1" style="width:145;">רשימת בריפים 
									מהיום והלאה</a></td>
							<td width="10"></td>
							<%end if%>
							<td><a href="default.asp?tab=9" class="button_edit_1" style="width:165;">לא תואם בריף</a></td>
							<td width="10"></td>
							<td><a href="default.asp?tab=11" class="button_edit_1" style="width:105;">החזרת טיולים 
									היום</a></td>
							<td width="10"></td>
							<td><a href="default.asp?tab=2" class="button_edit_1" style="width:95;">בריפים היום </a>
							</td>
							<td width="10"></td>
							<td><a href="default.asp?tab=3" class="button_edit_1" style="width:105;">בחו"ל ולא 
									התקשרו</a></td>
							<td width="10"></td>
							<td><a href="default.asp?tab=1" class="button_edit_1" style="width:95;">כרגע בחו"ל</a></td>
						</tr>
					</table> <!--אנא הקלד את שם החברה מימין או את שם איש הקשר לאיתור במאגר הלקוחות--><%'=arrTitles(25)%></td>
			</tr>
			<tr>
				<td width="100%" valign="top" align="center" style="height:100%">
					<table dir="<%=dir_var%>" border="0" cellpadding="0" cellspacing="1" width="100%" ID="Table2" style="height:100%">
						<tr>
							<td align="left" width="100%" valign="top" style="height:100%">
								<div style="position:relative;height:100%" >
									<div class="h_iframe" style="height:100%">
										<iframe name="frameScrreen" id="frameScrreen" src="WorkScreen.aspx?tab=<%=request.querystring("tab")%>"	ALLOWTRANSPARENCY=true  marginwidth="0" marginheight="0" hspace="0" vspace="0"  frameborder="0">
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

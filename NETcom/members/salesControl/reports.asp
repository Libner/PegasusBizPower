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
	
function report_open(fname){
if (document.getElementById("currentYear").value=='0')
{
  alert("אנא בחר שנה");
	return false;
	}
	    formsearch.action=fname;
		formsearch.target = '_blank';
		formsearch.submit();
	
}

	
	</script>
	<body>
		<FORM action="" method=POST id="Form1" name=formsearch target=_self>

		<table border="0" width="100%" style="height:100%" cellspacing="0" cellpadding="0" ID="Table1">
			<tr>
				<td width="100%" align="<%=align_var%>">
					<!--#include file="../../logo_top.asp"-->
				</td>
			</tr>
			<tr>
				<td width="100%" align="<%=align_var%>">
					<%numOftab = 108%>
					<%numOfLink =2
					dtCurrentDate=Now()
					%>
<%topLevel2 = 113 'current bar ID in top submenu - added 03/10/2019%>
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
									<br clear=all>
		<table border="0" cellspacing="1" width=450 align=center bgcolor="#646E77" cellpadding="4" dir="<%=dir_var%>" ID="Table7">


											<tr>
			<td width="80%" bgcolor="#DBDBDB" align="right" dir="rtl">	
			<SELECT NAME="currentYear" CLASS="norm" ID="currentYear" dir="<%=dir_var%>" style="width:60" >
	      <% For counter = -10 to 0 %>
	        <OPTION VALUE="<%=Year(dtCurrentDate)+counter%>" <% If (DatePart("yyyy", dtCurrentDate) = Year(dtCurrentDate)) Then Response.Write "SELECTED"%>><%=Year(dtCurrentDate)+counter%></OPTION>
	      <% Next %>
	      </SELECT>
			</td>
			<td width="20%" class="subject_form" nowrap align="right">&nbsp;<span id="Span4" name="word4">שנת קלנדרי</span>&nbsp;</td>
	</tr>
		<tr id=but_tr>
			<td bgcolor="#dbdbdb" align="center" colspan="2">
				<TABLE WIDTH=100% align=center BORDER=0 CELLSPACING=5 CELLPADDING=0 ID="Table6">
				 	<TR>
						<td width=5%>&nbsp;</td>
						<TD align=right width=90%><A class="but_menu" href="#" onclick="report_open('graphSalesControl1.aspx');"><span id=word9 name=word9>מכירות יעד מול ביצוע</span></a></TD>
						<td width=5%>&nbsp;1</td>
					</TR>	
						<TR>
						<td width=5%>&nbsp;</td>
						<TD align=right width=90%><A class="but_menu" href="#" onclick="report_open('graphSalesControl2.aspx');"><span id="Span3" name=word9>יעד מול ביצוע, רבעונים</span></a></TD>
						<td width=5%>&nbsp;2</td>
					</TR>			
					
								<TR>
						<td width=5%>&nbsp;</td>
						<TD align=right width=90%><A class="but_menu" href="#" onclick="report_open('graphSalesControl3.aspx');"><span id="Span1" name=word9>ביצוע שנה נוכחית  מול ביצוע שנה קודמת</span></a></TD>
						<td width=5%>&nbsp;3</td>
					</TR>	
					
							<TR>
						<td width=5%>&nbsp;</td>
						<TD align=right width=90%><A class="but_menu" href="#" onclick="report_open('graphSalesControl4.aspx');"><span id="Span2" name=word9>דוח טפסי מתעניין שנה נוכחית מול טפסי מתעניין שנה קודמת</span></a></TD>
						<td width=5%>&nbsp;4</td>
					</TR>	
					
							<TR>
						<td width=5%>&nbsp;</td>
						<TD align=right width=90%><A class="but_menu" href="#" onclick="report_open('graphSalesControl5.aspx');" ><span id="Span5" name=word9>אחוז סגירה שנת נוכחית מול אחוז סגירה שנת הקודמת</span></a></TD>
						<td width=5%>&nbsp;5</td>
					</TR>	
					
	
							<TR>
						<td width=5%>&nbsp;</td>
						<TD align=right width=90%><A class="but_menu" href="#" onclick="report_open('graphSalesControl6.aspx');" ><span id="Span6" name=word9>דוח שיחות מכירה שנה נוכחית מול שיחות מכירה שנת הקודמת</span></a></TD>
						<td width=5%>&nbsp;6</td>
					</TR>	
						
							<TR>
						<td width=5%>&nbsp;</td>
						<TD align=right width=90%><A href="#" style="background-color:#736BA6;color:#ffffff;font-weight:bold;text-decoration:none;padding:1px;padding-right:5px;padding-left:5px" onclick="report_open('graphSalesControl7.aspx');" ><span id="Span7" name=word9>כמות קישוריות המכירה בשנת נוכחית מול כמות קישוריות המכירה&nbsp;<br> בשנה קודמת </span></a></TD>
						<td width=5%>&nbsp;7</td>
					</TR>	
					<tr><td colspan=3 height=3></td></tr>
					
		
							<TR>
						<td width=5%>&nbsp;</td>
						<TD align=right width=90%><A style="background-color:#736BA6;color:#ffffff;font-weight:bold;text-decoration:none;padding:1px;padding-right:5px;padding-left:5px" href="#" onclick="report_open('graphSalesControl8.aspx');" ><span id="Span8" name=word9>&nbsp;אחוז הקישוריות המכירה שנת נוכחית מול אחוז קישוריות המכירה&nbsp;<br> שנת הקודמת</span></a></TD>
						<td width=5%>&nbsp;8</td>
					</TR>
						<tr><td colspan=3 height=3></td></tr>	
								<TR>
						<td width=5%>&nbsp;</td>
						<TD align=right width=90%><A style="background-color:#736BA6;color:#ffffff;font-weight:bold;text-decoration:none;padding:1px;padding-right:5px;padding-left:5px" href="#" onclick="report_open('graphSalesControl9.aspx');" ><span id="Span9" name=word9>כמות קישוריות השירות שנת נוכחית מול אחוז סגירה שנת הקודמת</span></a></TD>
						<td width=5%>&nbsp;9</td>
					</TR>	
						<tr><td colspan=3 height=3></td></tr>
								<TR>
						<td width=5%>&nbsp;</td>
						<TD align=right width=90%><A  style="background-color:#736BA6;color:#ffffff;font-weight:bold;text-decoration:none;padding:1px;padding-right:5px;padding-left:5px" href="#" onclick="report_open('graphSalesControl10.aspx');" ><span id="Span10" name=word9>&nbsp;אחוז הקישוריות השירות שנת נוכחית מול אחוז הקישוריות השירות&nbsp;<br>שנת הקודמת</span></a></TD>
						<td width=5%>&nbsp;10</td>
					</TR>	
					
									
</table></td></tr>
										</table>
									</div>
								</div>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</form>
	</body>
</html>
<%set con=Nothing%>

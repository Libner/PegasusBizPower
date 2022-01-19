<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SalesEfficiency1.aspx.vb" Inherits="bizpower_pegasus2018.SalesEfficiency1"%>
<!DOCTYPE HTML>
<html>
	<head>
		<title>SalesEfficiency</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
		<script type="text/javascript" src="//code.jquery.com/jquery-1.12.4.js"></script>
		<link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
		<script src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>
		<script type='text/javascript' src="https://jquery-ui.googlecode.com/svn-history/r3982/trunk/ui/i18n/jquery.ui.datepicker-he.js"></script>
		<style>
  button.new {
    background: -moz-linear-gradient(#00BBD6, #EBFFFF);
    background: -webkit-gradient(linear, 0 0, 0 100%, from(#00BBD6), to(#EBFFFF));
    filter: progid:DXImageTransform.Microsoft.gradient(startColorstr='#00BBD6', endColorstr='#EBFFFF');
    padding: 3px 7px;
    color: #333;
    -moz-border-radius: 5px;
    -webkit-border-radius: 5px;
    border-radius: 5px;
    border: 1px solid #666;
   }
   
.tooltip {
    position: relative;
	display: inline-block;
	// border-bottom: 1px dotted black; /* If you want dots under the hoverable text */
}

.tooltip .tooltiptext {
    visibility: hidden;
    width: 200px;
    background-color: #555;
    color: #fff;
    direction:rtl;
    text-align: center;
    border-radius: 6px;
    padding: 5px 0;
    position: absolute;
    z-index: 1;
    bottom: 125%;
    left: 50%;
    margin-left: -80px;
    opacity: 0;
    transition: opacity 0.3s;
}

.tooltip .tooltiptext::after {
    content: "";
    position: absolute;
    top: 100%;
    left: 50%;
    margin-left: -5px;
    border-width: 5px;
    border-style: solid;
    border-color: #555 transparent transparent transparent;
    
}

.tooltip:hover .tooltiptext {
    visibility: visible;
    opacity: 1;
}
</style>
		<script>
$(document).ready(function()
{ 
$('#Button2Div').hide();
$('#chkCountry').change(function(){
if(this.checked)
{
$('#plhSelect_Country').show();
$('#Button1Div').hide();
$('#Button2Div').show();
}
else
{
$('#plhSelect_Country').hide();
$('#Button1Div').show();
$('#Button2Div').hide();
}
}
);
//alert($("#chkCountry").checked)
$("input[name$='RadioType']").click(function() 
{
      var test = $(this).val();
      $("div div").hide();
        $("#plhSelect_1").hide();
           $("#plhSelect_2").hide();
      $("#plhSelect_"+test).show();
  }
  ); 
});
		</script>
		<script>
		  $(function() {
      $.datepicker.setDefaults($.datepicker.regional["he"]);
      $.datepicker.setDefaults({ dateFormat: 'dd/mm/y' });
	$('#dateStart').datepicker();
	$('#dateEnd').datepicker();
	
	$('#icDatePicker').on('click', function() {
      $('#dateStart').datepicker('show');
      return false
   });
   	$('#icDatePickerEnd').on('click', function() {
      $('#dateEnd').datepicker('show');
      return false
   });
});

		</script>
		<script>
		function openDays(uid)
		{
	
		  h = 400;
          w = 800;
          S_Wind = window.open("List_openDays.aspx?uid=" + uid+"&dStart="+document.getElementById("dateStart").value+"&dEnd="+document.getElementById("dateEnd").value, "W_openDays", "scrollbars=1,toolbar=0,top=150,left=50,width=" + w + ",height=" + h + ",align=center,resizable=1");
          S_Wind.focus();
          return false;
		}
			function openExcel()
		{
		
		//	window.open("ExcelWorkScreen.aspx","ExcelWork","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");


 var f = document.getElementById('Form2');
f.dateStartEx.value = document.getElementById("dateStart").value;
f.dateEndEx.value=document.getElementById("dateEnd").value;
if (document.getElementById('radio1').checked) {
  f.RadioTypeEx.value = document.getElementById('radio1').value;
}
if (document.getElementById('radio2').checked) {
  f.RadioTypeEx.value = document.getElementById('radio2').value;
}
f.seldepEx.value=document.getElementById("seldep").value;
f.selUserEx.value=document.getElementById("selUser").value;
f.action="ExcelSalesEfficiency.aspx"

f.target="_blank"
 //// window.open('ExcelWorkScreen.aspx', 'ExcelWork');
 f.submit();
}
  


		</script>
	</head>
	<body style="margin:0px">
		<form id="Form1" method="post" runat="server" name="Form1">
			<table border="0" cellpadding="2" cellspacing="0" align="center" width="100%">
				<tr>
					<td width="10"></td>
				</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="left">
				<tr>
					<td align="center" colspan="2" style="color:#000000;">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td height="30" align="center"><span style="COLOR: #6F6DA6;font-size:14pt">דו"ח איכות 
										עבודת נציג</span></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td colspan="2" width="100%" style="color:#000000;">
						<table align="center" cellpadding="0" cellspacing="0" style="color:#000000;background-color:#e1e1e1;border:solid 1px #6F6DA6">
							<tr>
								<td valign="top" style="border-right:solid 1px #6F6DA6"><table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td >
												<table border="0" cellpadding="2" cellspacing="2" width=100%>
													<tr>
														<td align="right"><b>(חלוקה ליעדים (דו"ח אקסל</b><input type="checkbox" id="chkCountry"></td>
													</tr>
													<tr>
														<td align="center" height="10" width=300>&nbsp;</td>
													</tr>
													<tr>
														<td align="right">
															<div id="plhSelect_Country" style="display:none">
																<table border="0" cellpadding="0" cellspacing="0">
																	<tr>
																		<TD>
																			<select runat="server" id="selCountry" class="searchList" name="selUser" style="width: 220px;height:60px;direction:rtl;font-size:8pt"
																				multiple="">
																			</select>
																		</TD>
																	</tr>
																</table>
															</div>
														</td>
													</tr>
													<tr>
														<td align="center" height="10">&nbsp;</td>
													</tr>
													<tr>
														<td align="center">
															<table>
																<tr>
																	<td align="center">
																<div id=Button2Div><asp:Button runat=server  id="Button2"></asp:Button></div></td>
											</td>
																</tr>
															</table>
														</td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</td>
								<td>
									<table border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td valign="middle">
												<table border="0" cellpadding="2" cellspacing="2">
													<tr>
														<td valign="middle">
															<table border="0" cellpadding="0" cellspacing="0">
																<tr>
																	<td valign="top" colspan="2" align="right"><b>טווח תאריכים</b></td>
																</tr>
																<tr>
																	<TD align="right" nowrap>
																		<a id="icDatePicker" href=""><img class="calendar_icon" src="../../images/calend.gif" class="iconButton" border="0"></a>
																		<input  id="dateStart" type="text"  name="dateStart" dir="ltr" class="passw" size=8 value="<%=dateStart%>"></TD>
																	<TD align="right">&nbsp;<span id="Span1" name="word7"><!--- תאריך מ--> מתאריך</span>&nbsp;</TD>
																</tr>
																<tr>
																	<TD align="right">
																		<a id="icDatePickerEnd" href=""><img class="calendar_icon" src="../../images/calend.gif" class="iconButton" border="0"></a>
																		<input  id="dateEnd" type="text"  name="dateEnd" dir="ltr" class="passw" size=8 value="<%=dateEnd%>"></TD>
																	<td align="right">&nbsp;<span id="Span1" name="word7"><!--- תאריך מ--> עד תאריך</span>&nbsp;</td>
																</tr>
															</table>
														</td>
													</tr>
												</table>
											</td>
											<td>
												<table border="0" cellpadding="2" cellspacing="2">
													<tr>
														<td align="right">מחלקות<input type="radio" id="radio1" name="RadioType" value="1" <%if RadioType=1 then%>checked<%end if%>></td>
													</tr>
													<tr>
														<td align="right">נציגים<input type="radio" id="radio2" name="RadioType" value="2"  <%if RadioType=2 then%>checked<%end if%>></td>
													</tr>
													<tr>
														<td valign="top">
															<div id="plhSelect_2" <%if RadioType=2 then%> style="display:block"<%else%>style="display:none" <%end if%>>
																<table border="0" cellpadding="0" cellspacing="0">
																	<tr>
																		<TD>
																			<select runat="server" id="selUser" class="searchList" name="selUser" style="width: 220px;height:60px;direction:rtl;font-size:8pt"
																				multiple="">
																			</select></TD>
																	</tr>
																</table>
															</div>
														</td>
													</tr>
													<tr>
														<td align="right" valign="top">
															<div id="plhSelect_1" <%if RadioType=1 then%> style="display:block"<%else%>style="display:none" <%end if%>>
																<table border="0" cellpadding="0" cellspacing="0">
																	<tr>
																		<TD><select runat="server" id="seldep" class="searchList" name="seldep" style="width: 220px;height:60px;direction:rtl;font-size:8pt"
																				multiple="">
																			</select></TD>
																	</tr>
																</table>
															</div>
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td align="center" colspan="2">
												<table>
													<tr>
														<td><div id=Button1Div><asp:Button runat=server  id="Button1" name="Button1"></asp:Button></div></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td height="20">&nbsp;</td>
				</tr>
				<tr><td align=center>
				<table border=0 cellpadding=0 cellspacing=0 width=80%>
				<tr><td><a href="#" onclick="openExcel();" class="button_small1" style="width:30px">Excel</a></td>
				<td align=center><span style="COLOR: #6F6DA6;font-size:14pt"><%if RadioType=1 then%>מחלקת <%=depName%><%else%>נציגים<%end if%> בין התאריכים <%=dateStart%> -  <%=dateEnd%></span></td></tr>
							</table>
							</td></TR>
				<tr>
					<td>
						<table cellpadding="1" cellspacing="1" width="80%" align="center" style="border:solid 0px #d3d3d3">
							<tr style="height:50px">
								<td colspan="6" class="title_sort" style="background-color:#BED49B;font-weight:bold"
									align="center">כמויות המכירות</td>
								<td colspan="3" class="title_sort" style="background-color:#BED49B;font-weight:bold"
									align="center">כמויות מתעניינים</td>
								<td colspan="5" class="title_sort" style="background-color:#BED49B;font-weight:bold"
									align="center">כמויות שיחות</td>
								<td class="title_sort" style="background-color:#BED49B;font-weight:bold" align="center">נוכחות</td>
								<td colspan="2" class="title_sort" style="background-color:#BED49B"></td>
								<td  class="title_sort" style="background-color:#BED49B"></td>
							</tr>
							<tr>
								<Td class="title_sort" align="center"><div class="tooltip">סגירה של מכירות ב"תהליך לא מלא<span class="tooltiptext">אחוז מכירות הנציג ב"תהליך לא מלא" מתוך סך המכירות שלו בתקופה המבוקשת</span></div>
								</Td>
								<Td class="title_sort" align="center"><div class="tooltip">כמות מכירות ב"תהליך לא מלא<span class="tooltiptext">כמות האנשים בטפסי "טופס רישום חתום" אשר פתח המשתמש בתאריכים המבוקשים ואשר אינם בסטטוס "מבוטל" תוך בדיקה שטפסי "מתעניין בטיול" אינם משויכים לאותו הנציג</span></div>
								</Td>
								<Td class="title_sort" align="center"><div class="tooltip">% סגירה של מכירות ב"תהליך מלא"<span class="tooltiptext">אחוז מכירות הנציג ב"תהליך מלא" מתוך סך המכירות שלו בתקופה המבוקשת</span></div>
								</Td>
								<Td class="title_sort" align="center"><div class="tooltip">כמות מכירות ב"תהליך מלא"<span class="tooltiptext">כמות האנשים בטפסי "טופס רישום חתום" אשר פתח המשתמש בתאריכים המבוקשים ואשר אינם בסטטוס "מבוטל" תוך בדיקה שגם טפסי "מתעניין בטיול" משויכים לאותו הנציג</span></div>
								</Td>
								<td class="title_sort" align="center">
								<div class="tooltip">כמות לקוחות שביטלו הרשמה<span class="tooltiptext">כמות האנשים (לקוחות) בטפסי "טופס רישום חתום" של הנציג אשר נמצאים בסטטוס מבוטל</span></div>
								</td>
								<td class="title_sort" align="center">
								<div class="tooltip">כמות מכירות כללית של הנציג<span class="tooltiptext">כמות האנשים בטפסי "טופס רישום חתום" אשר פתח המשתמש בתאריכים המבוקשים ואינם בסטטוס "בוטל"</span>
								</td>
								<td class="title_sort" align="center">
								<div class="tooltip">% סגירה של הנציג<span class="tooltiptext">אחוז הסגירה = כמות "טופס רישום חתום" של הנציג ביחס לכמות "טפסי מתעניין בטיול" שפתח"</span></div>
								</td>
								<Td class="title_sort" align="center"><div class="tooltip">מתוכם בכמה בוצעה מכירה<span class="tooltiptext">כמה מתוך סך טפסי המתעניין של הנציג בוצעה בגינם מכירה (טופס רישום חתום של הנציג). לא כולל טפסים אשר נמצאים בסטטוס "בוטל"</span></div>
								</Td>
								<TD class="title_sort" align="center">כמות טפסי המתעניין</TD>
								<td class="title_sort" align="center">
								<div class="tooltip">ממוצע שיחות יומי<span class="tooltiptext">חלוקת הנתון מעמודת "כמות כללית של סך השיחות" בימי העבודה שנכח הנציג</span></div>
								</td>
								<td class="title_sort" align="center">
								<div class="tooltip">כמות כללית של סך השיחות<span class="tooltiptext">פסי מתעניין בטיול + תיעודי שיחות נכנסות + תיעודי שיחות יוצאות אמיתיות (לא כולל "אין מענה")</span></div>
								</td>
								<td class="title_sort" align="center">כמות טפסי המתעניין</td>
								<td class="title_sort" align="center">
										<div class="tooltip">כמות תיעודי שיחה יוצאת<span class="tooltiptext">כמות תיעודי השיחה היוצאים עם חלוקה לתיעודים עם תוכן לאומת תיעודים של "אין מענה"</span></div>
						        </td>
								<td class="title_sort" align="center">כמות תיעודי שיחה נכנסת</td>
								<td class="title_sort" align="center">
								<div class="tooltip">כמות ימים שלא נכח בעבודה <span class="tooltiptext">כמות הימים בהם הנציג לא ביצע לוג אין למערכת </span></div>
						        </td>
								<td class="title_sort" align="center" nowrap>מחלקה</td>
								<td class="title_sort" align="center" nowrap>נציג</td>
								<td class="title_sort" align="center" nowrap>&nbsp;</td>
							</tr>
							<asp:Repeater ID="rptData" Runat="server">
								<ItemTemplate>
									<tr style="height:30px">
										<Td style="background-color:#ffffff;border-left:solid 1px #e1e1e1;border-right:solid 1px #e1e1e1"
											align="center">---15</Td>
										<Td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center">--14</Td>
										<Td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center">--13"</Td>
										<Td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center">--12</Td>
										<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#Container.Dataitem("numberOf16735_40660Bitul")%></td>
										<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#Container.dataItem("numberOf16735_40660")%></td>
										<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center">--9</td>
										<Td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center">--8</Td>
										<TD style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#Container.DataItem("numberOf16504")%></TD>
										<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center">--6</td>
										<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center">--5</td>
										<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#IIF(Container.DataItem("numberOf16504")=0,"",Container.DataItem("numberOf16504"))%></td>
										<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#IIF(Container.DataItem("numberOf16470_40811_out")=0 ,"",Container.DataItem("numberOf16470_40811_out"))%></td>
										<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#IIF(Container.DataItem("numberOf16470_40811_in")=0 ,"",<a href="">Container.DataItem("numberOf16470_40811_in"))%></a></td>
										<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%'#Container.dataItem("numberOfDays")%><asp:Label id=pDays Runat=server></asp:Label></td>
										<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="right" nowrap><%'#Container.DataItem("departmentName")%></td>
										<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="right" nowrap><%'#Container.Dataitem("User_Name")%></td>
										<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="right" nowrap>&nbsp;<%#Container.ItemIndex+1%>&nbsp;</td>
									</tr>
								</ItemTemplate>
								<SeparatorTemplate>
									<tr>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
								    	<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
										<td style="background-color:#e1e1e1;height:0.1px"></td>
									</tr>
								</SeparatorTemplate>
								<FooterTemplate>
									<tr style="background-color:#ffd011;height:30px">
										<Td align="center">--15</Td>
										<Td align="center">--14</Td>
										<Td align="center">--13"</Td>
										<Td align="center">--12</Td>
										<td align="center"><%=SumVar16735_40660Bitul%></td>
										<td align="center"><%=SumVar16735_40660%></td>
										<td align="center">---9</td>
										<Td align="center">--8</Td>
										<TD align="center"><%=SumVar16504%></TD>
										<td align="center">--6</td>
										<td align="center">--5</td>
										<td align="center"><%=SumVar16504%></td>
										<td align="center"><%=SumVar16470_40811_out%></td>
										<td align="center"><%=SumVar16470_40811_in%></td>
										<td align="center"><%=SumVar1%></td>
										<td align="center">&nbsp;</td>
										<td align="center">&nbsp;</td>
										<td align="center">&nbsp;</td>
									</tr>
								</FooterTemplate>
							</asp:Repeater>
						</table>
					</td>
				</tr>
			</table>
	
			<DIV ID='CalendarDiv' STYLE='POSITION:absolute;VISIBILITY:hidden;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>
		</form>
			<form id="Form2" method="post"  name="Form2" target="_blank">
	
				<input type=hidden id="dateStartEx" name="dateStartEx">
				<input type=hidden id="dateEndEx" name="dateEndEx">
				<input type=hidden id="RadioTypeEx" name="RadioTypeEx">
				<input type=hidden id="seldepEx" name="seldepEx">
				<input type=hidden id="selUserEx" name="selUserEx">
				
		</form>
	</body>
</html>

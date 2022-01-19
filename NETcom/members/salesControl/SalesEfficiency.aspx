<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SalesEfficiency.aspx.vb" Inherits="bizpower_pegasus2018.SalesEfficiency" %>
<%@ OutputCache Duration="60" Location="Server" VaryByParam="*" %>
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
    width: 250px;
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
$(document).load(function()
{ 

parent.showIframeLoader(true);
});
		</script>
		<script>
$(document).ready(function()
{ 
parent.showIframeLoader(false);

$('#Button2').click(
function()
{
 // $(".loading_section").css("display", "block");
//alert ("click");
//-----
 var f = document.getElementById('Form3');
f.dateStartEx.value = document.getElementById("dateStart").value;
f.dateEndEx.value=document.getElementById("dateEnd").value;
if (document.getElementById('radio1').checked) {
  f.RadioTypeEx.value = document.getElementById('radio1').value;
  	if (!$("#seldep option:selected").length)
		{
		alert ("אנא בחר מחלקה")
		return false;
		}
  
  
}
if (document.getElementById('radio2').checked) {
  f.RadioTypeEx.value = document.getElementById('radio2').value;
}
//f.seldepEx.value=document.getElementById("seldep").value;

f.action="ExcelCountrySalesEfficiency.aspx"

var selectedValuesCountry=''
$("#selCountry option:selected").each(function() {
	selectedValuesCountry += $(this).val() + ",";
});

f.selCountryEx.value=selectedValuesCountry;
var selectedValuesDep=''
$("#seldep option:selected").each(function() {
	selectedValuesDep += $(this).val() + ",";
});

f.seldepEx.value=selectedValuesDep;
var selectedValuesUser=''
$("#selUser option:selected").each(function() {
	selectedValuesUser += $(this).val() + ",";
});
f.selUserEx.value=selectedValuesUser;


//alert(selectedValues)
//f.target="_blank"
//alert("form3")
//alert(f.action)
 //// window.open('ExcelWorkScreen.aspx', 'ExcelWork');
f.submit();
 return false;
 
///---
}
);

$('#Button1').click(
function()
{
 // if($('#radio1').is(':checked')) { alert("it's checked"); }
 
if($('#radio1')[0].checked)
{
		if (!$("#seldep option:selected").length)
		{
		alert ("אנא בחר מחלקה")
		return false;
		}
}


if($('#radio2')[0].checked) {
  
   	if (!$("#selUser option:selected").length)
		{
		alert ("אנא בחר נציג")
		return false;
		}
}

  $(".loading_section").css("display", "block");
//alert ("click");
}
);
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
		function OpenWindowWithPost(url, windowoption, name, params)
{
 var form = document.createElement("form");
 form.setAttribute("method", "post");
 form.setAttribute("action", url);
 form.setAttribute("target", name);
 for (var i in params)
 {
   if (params.hasOwnProperty(i))
   {
     var input = document.createElement('input');
     input.type = 'hidden';
     input.name = i;
     input.value = params[i];
  //alert(input.value)
     form.appendChild(input);
   }
 }
 document.body.appendChild(form);
 //note I am using a post.htm page since I did not want to make double request to the page
 //it might have some Page_Load call which might screw things up.
 window.open(url, name, windowoption);
 form.submit();
 document.body.removeChild(form);
}
function open16470_40105(uid)
{
		var param = { 'usname' : uid,'inpsearch40105' :'אין מענה','inpsearch40811' :'שיחה יוצאת' };

OpenWindowWithPost("ListAppeals.asp?prodId=16470&dStart="+document.getElementById("dateStart").value+"&dEnd="+document.getElementById("dateEnd").value, "width=1000, height=600, left=100, top=100, resizable=yes, scrollbars=yes", "NewFile", param);


}
function openOut(uid)
{
		var param = { 'usname' : uid,'inpsearch40105' :'תוכן שיחות','inpsearch40811' :'שיחה יוצאת' };

OpenWindowWithPost("ListAppeals.asp?prodId=16470&dStart="+document.getElementById("dateStart").value+"&dEnd="+document.getElementById("dateEnd").value, "width=1000, height=600, left=100, top=100, resizable=yes, scrollbars=yes", "NewFile", param);


}
function open16735Bitulim(uid)
		{
		var param = { 'usname' : uid,'inpsearch40661' :'בוטל' };

OpenWindowWithPost("ListAppeals.asp?prodId=16735&dStart="+document.getElementById("dateStart").value+"&dEnd="+document.getElementById("dateEnd").value, "width=1000, height=600, left=100, top=100, resizable=yes, scrollbars=yes", "NewFile", param);


}
function opennumberOf16735To16504(uid)
{
var param = { 'UserIdOwner' : uid,'inpsearch40661' :'','p16735_16470totalperUser':'3' };
OpenWindowWithPost("ListAppeals16504.asp?prodId=16504&dStart="+document.getElementById("dateStart").value+"&dEnd="+document.getElementById("dateEnd").value, "width=1000, height=600, left=100, top=100, resizable=yes, scrollbars=yes", "NewFile", param);
}

function opennumberNoOf16735_16470totalperUser(uid)
{
var param = { 'usname' : uid,'inpsearch40661' :'','p16735_16470totalperUser':'2' };
OpenWindowWithPost("ListAppeals.asp?prodId=16735&dStart="+document.getElementById("dateStart").value+"&dEnd="+document.getElementById("dateEnd").value, "width=1000, height=600, left=100, top=100, resizable=yes, scrollbars=yes", "NewFile", param);
}
function opennumberOf16735_16470totalperUser(uid)
{
		var param = { 'usname' : uid,'inpsearch40661' :'','p16735_16470totalperUser':'1' };

OpenWindowWithPost("ListAppeals.asp?prodId=16735&dStart="+document.getElementById("dateStart").value+"&dEnd="+document.getElementById("dateEnd").value, "width=1000, height=600, left=100, top=100, resizable=yes, scrollbars=yes", "NewFile", param);

}
function open16504(uid)
		{
		var param = { 'usname' : uid,'inpsearch40811' :'שיחה נכנסת' };

OpenWindowWithPost("ListAppeals.asp?prodId=16504&dStart="+document.getElementById("dateStart").value+"&dEnd="+document.getElementById("dateEnd").value, "width=1000, height=600, left=100, top=100, resizable=yes, scrollbars=yes", "NewFile", param);


}
function open16735(uid)
		{
		var param = { 'usname' : uid,'inpsearch40661' :'' };

OpenWindowWithPost("ListAppeals.asp?open16735=1&prodId=16735&dStart="+document.getElementById("dateStart").value+"&dEnd="+document.getElementById("dateEnd").value, "width=1000, height=600, left=100, top=100, resizable=yes, scrollbars=yes", "NewFile", param);


}
		function open16470_40811_in(uid)
		{
		var param = { 'usname' : uid,'inpsearch40811' :'שיחה נכנסת' };

OpenWindowWithPost("ListAppeals.asp?prodId=16470&dStart="+document.getElementById("dateStart").value+"&dEnd="+document.getElementById("dateEnd").value, "width=1000, height=600, left=100, top=100, resizable=yes, scrollbars=yes", "NewFile", param);


		
		//  h = 400;
        //  w = 1300;
        //  S_Wind = window.open("ListAppeals.asp?prodID=16470&uid=" + uid+"&dStart="+document.getElementById("dateStart").value+"&dEnd="+document.getElementById("dateEnd").value, "W_open16470", "scrollbars=1,toolbar=0,top=150,left=50,width=" + w + ",height=" + h + ",align=center,resizable=1");
       //   S_Wind.focus();
        //  return false;
		}
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
											<td>
												<table border="0" cellpadding="2" cellspacing="2" width="100%">
													<tr>
														<td align="right"><b>(חלוקה ליעדים (דו"ח אקסל</b><input type="checkbox" id="chkCountry"></td>
													</tr>
													<tr>
														<td align="center" height="2" width="300">&nbsp;</td>
													</tr>
													<tr>
														<td align="right">
															<div id="plhSelect_Country" style="display:none">
																<table border="0" cellpadding="0" cellspacing="0">
																	<tr>
																		<TD>
																			<select runat="server" id="selCountry" class="searchList" name="selUser" style="width: 220px;height:90px;direction:rtl;font-size:8pt"
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
																		<div id="Button2Div"><asp:Button runat="server" id="Button2"></asp:Button></div>
																	</td>
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
											<td><div id="Button1Div" style="display: block"><asp:Button runat="server" id="Button1" name="Button1"></asp:Button></div>
											</td>
										</tr>
										<tr>
											<td>
												<div class="loading_section" style="display: none;">
													<div class="loading_block">
														<img src="../../images/loading.gif">
													</div>
												</div>
											</td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			</td> </tr>
			<tr>
				<td height="20">&nbsp;</td>
			</tr>
			<tr>
				<td align="center">
					<table border="0" cellpadding="0" cellspacing="0" width="80%">
						<tr>
							<td><a href="#" onclick="openExcel();" class="button_small1" style="width:30px">Excel</a></td>
							<td align="center"><span style="COLOR: #6F6DA6;font-size:14pt"><%if RadioType=1 then%>מחלקת
									<%=depName%>
									<%else%>
									נציגים<%end if%>
									בין התאריכים
									<%=dateStart%>
									-
									<%=dateEnd%>
								</span>
							</td>
						</tr>
					</table>
				</td>
			</tr>
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
							<td class="title_sort" style="background-color:#BED49B"></td>
						</tr>
						<tr>
							<Td class="title_sort" align="center"><div class="tooltip">סגירה של מכירות ב"תהליך לא 
									מלא<span class="tooltiptext">אחוז מכירות הנציג ב"תהליך לא מלא" מתוך סך המכירות שלו 
										בתקופה המבוקשת</span></div>
							</Td>
							<Td class="title_sort" align="center"><div class="tooltip">כמות מכירות ב"תהליך לא מלא<span class="tooltiptext">כמות 
										האנשים בטפסי "טופס רישום חתום" אשר פתח המשתמש בתאריכים המבוקשים ואשר אינם 
										בסטטוס "מבוטל" תוך בדיקה שטפסי "מתעניין בטיול" אינם משויכים לאותו הנציג</span></div>
							</Td>
							<Td class="title_sort" align="center"><div class="tooltip">% סגירה של מכירות ב"תהליך 
									מלא"<span class="tooltiptext">אחוז מכירות הנציג ב"תהליך מלא" מתוך סך המכירות שלו 
										בתקופה המבוקשת</span></div>
							</Td>
							<Td class="title_sort" align="center"><div class="tooltip">כמות מכירות ב"תהליך מלא"<span class="tooltiptext">כמות 
										האנשים בטפסי "טופס רישום חתום" אשר פתח המשתמש בתאריכים המבוקשים ואשר אינם 
										בסטטוס "מבוטל" תוך בדיקה שגם טפסי "מתעניין בטיול" משויכים לאותו הנציג</span></div>
							</Td>
							<td class="title_sort" align="center">
								<div class="tooltip">כמות לקוחות שביטלו הרשמה<span class="tooltiptext">כמות האנשים 
										(לקוחות) בטפסי "טופס רישום חתום" של הנציג אשר נמצאים בסטטוס מבוטל</span></div>
							</td>
							<td class="title_sort" align="center">
								<div class="tooltip">כמות מכירות כללית של הנציג<span class="tooltiptext">כמות האנשים 
										בטפסי "טופס רישום חתום" אשר פתח המשתמש בתאריכים המבוקשים ואינם בסטטוס "בוטל"</span>
							</td>
							<td class="title_sort" align="center">
								<div class="tooltip">% סגירה של הנציג<span class="tooltiptext">אחוז הסגירה = כמות "טופס 
										רישום חתום" של הנציג ביחס לכמות "טפסי מתעניין בטיול" שפתח"</span></div>
							</td>
							<Td class="title_sort" align="center"><div class="tooltip">מתוכם בכמה בוצעה מכירה<span class="tooltiptext">כמה 
										מתוך סך טפסי המתעניין של הנציג בוצעה בגינם מכירה (טופס רישום חתום של הנציג). לא 
										כולל טפסים אשר נמצאים בסטטוס "בוטל"</span></div>
							</Td>
							<TD class="title_sort" align="center">כמות טפסי המתעניין</TD>
							<td class="title_sort" align="center">
								<div class="tooltip">ממוצע שיחות יומי<span class="tooltiptext">חלוקת הנתון מעמודת "כמות 
										כללית של סך השיחות" בימי העבודה שנכח הנציג</span></div>
							</td>
							<td class="title_sort" align="center">
								<div class="tooltip">כמות כללית של סך השיחות<span class="tooltiptext">פסי מתעניין בטיול 
										+ תיעודי שיחות נכנסות + תיעודי שיחות יוצאות אמיתיות (לא כולל "אין מענה")
										<%'if false then%>	
											<br>מציג את כל השיחות	אשר לא מכילות את הטקסט:
										<br>אין מענה ,תא קולי, ממתינה ,לא עונה,אין תשובה, תועד כפול, לא ענה,לא יכול לדבר, לא יכולה לדבר, ללא מענה, תפוס, ניתוק, לא זמין, לא זמינה
										<%'end if%>
										</span></div>
						
										</span></div>
							</td>
							<td class="title_sort" align="center">כמות טפסי המתעניין</td>
							<td class="title_sort" align="center">
								<div class="tooltip">כמות תיעודי שיחה יוצאת<span class="tooltiptext">כמות תיעודי השיחה 
										היוצאים עם חלוקה לתיעודים עם תוכן לאומת תיעודים של "אין מענה"
										<%'if false then%>
										<br>מציג את כל השיחות	אשר מכילות את הטקסט:
										<br>אין מענה ,תא קולי, ממתינה ,לא עונה,אין תשובה, תועד כפול, לא ענה,לא יכול לדבר, לא יכולה לדבר, ללא מענה, תפוס, ניתוק, לא זמין, לא זמינה<%'end if%><br>
										שיחות שאין מענה/שיחות אמיתיות
										</span></div>
							</td>
							<td class="title_sort" align="center">כמות תיעודי שיחה נכנסת</td>
							<td class="title_sort" align="center">
								<div class="tooltip">כמות ימים שלא נכח בעבודה <span class="tooltiptext">כמות הימים בהם 
										הנציג לא ביצע לוג אין למערכת </span>
								</div>
							</td>
							<td class="title_sort" align="center" nowrap>מחלקה</td>
							<td class="title_sort" align="center" nowrap>נציג</td>
							<td class="title_sort" align="center" nowrap>&nbsp;</td>
						</tr>
						<asp:Repeater ID="rptData" Runat="server">
							<ItemTemplate>
								<tr style="height:30px">
									<Td style="background-color:#ffffff;border-left:solid 1px #e1e1e1;border-right:solid 1px #e1e1e1"
										align="center"><%#IIF(cint(Container.dataItem("numberOf16735_40660"))-cint(Container.dataItem("numberOf16735_16470totalperUser"))>0,(Math.Round((cint(Container.dataItem("numberOf16735_40660"))-cint(Container.dataItem("numberOf16735_16470totalperUser")))/cint(Container.dataItem("numberOf16735_40660")),2))*100 &"%","")%>
										<!---15--></Td>
									<Td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%'#Container.dataItem("numberOf16735_40660")%><%'#Container.dataItem("numberOf16735_16470totalperUser")%><%#IIF (cint(Container.dataItem("numberOf16735_40660"))-cint(Container.dataItem("numberOf16735_16470totalperUser"))=0,"","<a href='#' style='font-color:#000000' onClick=""opennumberNoOf16735_16470totalperUser('" & func.altFix(trim(Container.Dataitem("User_Name"))) &"')"">" & cint(Container.dataItem("numberOf16735_40660"))-cint(Container.dataItem("numberOf16735_16470totalperUser")) &"</a>")%><!--14--></Td>
									<Td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#IIF (Container.DataItem("numberOf16735_40660")>0,(Math.Round(cint(Container.DataItem("numberOf16735_16470totalperUser"))/cint(Container.DataItem("numberOf16735_40660")),2))*100 &"%","")%></Td>
									<Td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#IIF (Container.dataItem("numberOf16735_16470totalperUser")=0,"","<a href='#' style='font-color:#000000' onClick=""opennumberOf16735_16470totalperUser('" & func.altFix(trim(Container.Dataitem("User_Name"))) &"')"">" & Container.DataItem("numberOf16735_16470totalperUser") &"</a>")%></Td>
									<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#IIF(Container.dataItem("numberOf16735_40660Bitul")=0,"","<a href='#' style='font-color:#000000' onClick=""open16735Bitulim('" & func.altFix(trim(Container.Dataitem("User_Name"))) &"')"">" &Container.Dataitem("numberOf16735_40660Bitul") &"</a>")%><%'#Container.Dataitem("numberOf16735_40660Bitul")%></td>
									<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#IIF(Container.DataItem("numberOf16735_40660")=0,"","<a href='#' style='font-color:#000000' onClick=""open16735('" & func.altFix(trim(Container.Dataitem("User_Name"))) &"')"">" & Container.DataItem("numberOf16735_40660") &"</a>")%><!--כמות מכירות כללית של הנציג--></td>
									<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><!--9--><%#IIF (Container.DataItem("numberOf16504")>0,(Math.Round(cint(Container.DataItem("numberOf16735To16504"))/cint(Container.DataItem("numberOf16504")),2))*100 &"%","")%></td>
									<Td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><!--8--><%#IIF(Container.DataItem("numberOf16735To16504")=0,"","<a href='#' style='font-color:#000000' onClick=""opennumberOf16735To16504('" & func.altFix(trim(Container.Dataitem("User_Id"))) &"')"">" & Container.DataItem("numberOf16735To16504") &"</a>")%>
									<%'#Container.DataItem("numberOf16735To16504")%></Td>
									<TD style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#IIF(Container.DataItem("numberOf16504")=0,"","<a href='#' style='font-color:#000000' onClick=""open16504('" & func.altFix(trim(Container.Dataitem("User_Name"))) &"')"">" & Container.DataItem("numberOf16504") &"</a>")%><%'#Container.DataItem("numberOf16504")%></TD>
									<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%'#func.fixNumeric(Container.dataItem("numberOfDays"))%><%#IIF (func.fixNumeric(Container.dataItem("numberOfDays"))=0,"",Math.Round((CDbl(Container.DataItem("numberOf16504"))+CDbl(Container.DataItem("numberOf16470_40811_in"))+CDbl(Container.DataItem("numberOf16470_40105")))/CDbl(Container.dataItem("numberOfDays")),2))%></td>
									<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center">
									<%#cint(Container.DataItem("numberOf16504"))+cint(Container.DataItem("numberOf16470_40811_in"))+(cint(Container.DataItem("numberOf16470_40811_out"))-cint(Container.DataItem("numberOf16470_40105"))) %></td>
									<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#IIF(Container.DataItem("numberOf16504")=0,"","<a href='#' style='font-color:#000000' onClick=""open16504('" & func.altFix(trim(Container.Dataitem("User_Name"))) &"')"">" & Container.DataItem("numberOf16504") &"</a>")%></td>
									<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center">
										<%#IIF (cint(Container.DataItem("numberOf16470_40811_out"))-cint(Container.DataItem("numberOf16470_40105"))=0,"0","<a href='#' style='font-color:#000000' onClick=""openOut('" & func.altFix(trim(Container.Dataitem("User_Name"))) &"')"">" & cint(Container.DataItem("numberOf16470_40811_out"))-cint(Container.DataItem("numberOf16470_40105")) &"</a>")%>
										/
									<%#IIF (Container.DataItem("numberOf16470_40105")=0,"0","<a href='#' style='font-color:#000000' onClick=""open16470_40105('" & func.altFix(trim(Container.Dataitem("User_Name"))) &"')"">" & Container.DataItem("numberOf16470_40105") &"</a>")%>
										<%'#cint(Container.DataItem("numberOf16470_40811_out"))-cint(Container.DataItem("numberOf16470_40105"))%>
										<%'#IIF(Container.DataItem("numberOf16470_40811_out")=0 ,"",Container.DataItem("numberOf16470_40811_out"))%>
									</td>
									<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%#IIF(Container.DataItem("numberOf16470_40811_in")=0 ,"","<a href='#' style='font-color:#000000' onClick=""open16470_40811_in('" & func.altFix(trim(Container.Dataitem("User_Name"))) &"')"">" & Container.DataItem("numberOf16470_40811_in") &"</a>")%></td>
									<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="center"><%'=WorkDays%><%'#Container.dataItem("numberOfDays")%><asp:Label id="pDays" Runat="server"></asp:Label></td>
									<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="right" nowrap><%#Container.DataItem("departmentName")%></td>
									<td style="background-color:#ffffff;border-right:solid 1px #e1e1e1" align="right" nowrap><%#Container.Dataitem("User_Name")%></td>
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
									<Td align="center"><%if SumVar16735_40660>0 then%><%=Math.Round((SumVar16735_40660-SumVarnumberOf16735_16470totalperUser)/SumVar16735_40660,2)*100%>%<%end if%><!--15--></Td>
									<Td align="center"><%=SumVar16735_40660-SumVarnumberOf16735_16470totalperUser%></Td>
									<Td align="center"><%if SumVar16735_40660>0 then%><%=Math.Round(SumVarnumberOf16735_16470totalperUser/SumVar16735_40660,2)*100%>%<%end if%><!--13"--></Td>
									<Td align="center"><%=SumVarnumberOf16735_16470totalperUser%><!---12--></Td>
									<td align="center"><%=SumVar16735_40660Bitul%></td>
									<td align="center"><%=SumVar16735_40660%></td>
									<td align="center"><!--9--><%if SumVar16504>0 then%><%=Math.Round(sum16735To16504/SumVar16504,2)*100%>%<%end if%></td>
									<Td align="center"><!---8--><%=sum16735To16504%></Td>
									<TD align="center"><%=SumVar16504%></TD>
									<td align="center"><%if func.fixNumeric(sumNumberDaysWork)>0 then%><%=Math.Round(sumColumn5/sumNumberDaysWork,2)%><%end if%></td>
									<td align="center"><%=sumColumn5%></td>
									<td align="center"><%=SumVar16504%></td>
									<td align="center"><%=SumVar16470_40811_out - SumVar16470_40105%>
									 / 
									 <%=SumVar16470_40105%></td>
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
		<form id="Form2" method="post" name="Form2" target="_blank">
			<input type="hidden" id="dateStartEx" name="dateStartEx"> <input type="hidden" id="dateEndEx" name="dateEndEx">
			<input type="hidden" id="RadioTypeEx" name="RadioTypeEx"> <input type="hidden" id="seldepEx" name="seldepEx">
			<input type="hidden" id="selUserEx" name="selUserEx">
		</form>
		<form id="Form3" method="post" name="Form3" target="DoSubmit" onsubmit="DoSubmit = window.open('about:blank','DoSubmit','width=400,height=350');">
			<input type="hidden" id="selCountryEx" name="selCountryEx"> 
			<input type="hidden" id="dateStartEx" name="dateStartEx"> 
			<input type="hidden" id="dateEndEx" name="dateEndEx">
			<input type="hidden" id="RadioTypeEx" name="RadioTypeEx"> 
			<input type="hidden" id="seldepEx" name="seldepEx">
			<input type="hidden" id="selUserEx" name="selUserEx">
		</form>
	</body>
</html>

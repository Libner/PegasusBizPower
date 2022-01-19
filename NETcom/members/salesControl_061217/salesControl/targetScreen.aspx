<%@ Page Language="vb" AutoEventWireup="false" Codebehind="targetScreen.aspx.vb" Inherits="bizpower_pegasus.targetScreen"%>
<!DOCTYPE html>
<html>
	<head>
		<title>screenSetting</title>
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
		<script type="text/javascript" src="//code.jquery.com/jquery-1.12.4.js"></script>
		<script src="../../include/CalendarPopup.js" charset="windows-1255"></script>
		<style>
.tooltip {
    position: relative;
    display: inline-block;
    /*border-bottom: 1px dotted black;*/
}

.tooltip .tooltiptext {
    visibility: hidden;
    width: 320px;
    background-color: #555;
    color: #fff;
    text-align: center;
    border-radius: 6px;
    padding: 5px 0;
    position: absolute;
    z-index: 1;
    bottom: 125%;
    left: 50%;
    margin-left: -115px;
    opacity: 0;
    transition: opacity 1s;
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
	function OpenInput(n,value)
{
//alert(value)
//$('form input').hide();
//$('div[id^="dvT"').show();

document.getElementById("dvT"+n+"_"+value).style.display="none";
document.getElementById("dv"+n+"_"+value).style.display="block";

document.getElementById(n+"_"+value).style.display = "block";

}
	
	function changeInput(n,value)
{
//alert("changeInput")
//update value-------
//alert(value)
	var  formData = "name="+n+"&dep="+<%=dep%>+"&Uid="+value+"&value="+document.getElementById(n+"_"+ value).value; 
//alert (formData);
 
$.ajax({
    url : "UpdateDataMonthly.aspx",
    type: "POST",
    data : formData,
    success: function(data, textStatus, jqXHR)
    {
   //alert(data)
   // var ResData new Array();
   // var ResData=data.split("_");
   //data - response from server
    },
    error: function (jqXHR, textStatus, errorThrown)
    {
 
    }
});
	

////
 document.getElementById("dv"+n+"_"+value).style.display="none";
		document.getElementById("dvT"+n+"_"+value).style.display="block";
 document.getElementById("dvT"+n+"_"+value).innerHTML=document.getElementById(n+"_"+ value).value;


}	
function validate(evt) {
  var theEvent = evt || window.event;
  var key = theEvent.keyCode || theEvent.which;
  key = String.fromCharCode( key );
  var regex = /[0-9]|\./;
  if( !regex.test(key) ) {
    theEvent.returnValue = false;
    if(theEvent.preventDefault) theEvent.preventDefault();
  }
}
function SendSub(pYear)
{
document.getElementById("y").value=pYear
document.Form1.submit()
//alert(pYear);
}
		</script>
	</head>
	<body>
		<form id="Form1" method="post" runat="server">
			<table border="0" cellpadding="2" cellspacing="0" align="center" width="100%">
				<tr>
					<td>
					<td width="10"></td>
				</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="left">
				<tr>
					<td align="left">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td height="30" align="center"><span style="COLOR: #6F6DA6;font-size:14pt">טבלת הגדרות יעדים שנתיים למכירות ולשירות וביצוע בפועל</span></td>
							</tr>
							<tr>
								<td width="100%" style="color:#000000;background-color:#e1e1e1e1;">
									<table border="0" cellpadding="0" cellpadding="0" width="100%" style="color:#000000;background-color:#E6E6E6;">
										<tr>
											<td align="center">
												<table border="0" cellpadding="2" cellspacing="2">
													<tr>
																<td><input type="submit" value="שנה קודמת" class="but_menu" style="width:90;cursor:pointer;BACKGROUND-COLOR: #ffd011;color:#000000"
																id="ButtonB" name="ButtonB" onclick="Javascript:SendSub(<%=Year(Now())-1%>)"></td>
												
													<td><input type="submit" value="שנה נוכחית" class="but_menu" style="width:90;cursor:pointer;BACKGROUND-COLOR: #ffd011;color:#000000"
																id="ButtonT" name="ButtonT" onclick="Javascript:SendSub(<%=Year(Now())%>)"></td>
																			<td><input type="submit" value="שנה הבאה" onclick="Javascript:SendSub(<%=Year(Now())+1%>)" class="but_menu" style="width:90;cursor:pointer;BACKGROUND-COLOR: #ffd011;color:#000000"
																id="ButtonN" name="ButtonN"></td>
												<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
															<td><input type="submit" value="הצג נתונים" class="but_menu" style="width:90;cursor:pointer"
																id="Button1" name="Button1"></td>
																
														<td>
															<select runat="server" id="seldep" class="searchList" dir="rtl" name="seldep" style="height:30px;direction:rtl;">
															</select>
															</select>
														</td>
														<td align="center"><span style="font-size:14pt"><!--onchange="this.form.submit()"--><select id="y" name="y" style="height:30px">
																	<%for yy=2014 to Year(Now())+1%>
																	<option value="<%=yy%>" <%if currYear=yy then%>selected<%end if%> ><%=yy%></option>
																	<%next%>
																</select>
															</span>
														</td>
														
													</tr>
												</table>
											</td>
										</tr>
										<tr>
										<td>
											<table cellpadding="0" cellspacing="0" width="100%" dir=rtl align="center" style="border:solid 1px #d3d3d3">
										
										<tr>
										<td valign=top width=22%>
											<table border=0 cellpadding=0 cellspacing=0 width=100%>
											<tr>
											<td align="right" valign=top height=40 class="title_sort">&nbsp;<td>
											</tr>
													<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
							
											<tr><td height=40 valign=middle align=right>&nbsp;ביצוע מכירות שנה קודמת&nbsp;</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
								
									<tr><td height=40 align=right>&nbsp;יעד מכירות שנה נוכחית&nbsp;</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=right>היעד הנוכחי ביחס לביצוע של השנה הקודמת (באחוזים)&nbsp;</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=right>ביצוע מכירות שנה נוכחית&nbsp;</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=right>הביצוע הנוכחי ביחס לביצוע בשנה קודמת (באחוזים)&nbsp;</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=right>
									<div class="tooltip">&nbsp;היחס שבוצע בפועל של קישוריות השירות של שנה קודמת אל מול שיחות השירות (באחוזים)<span class="tooltiptext">אחוז המייצג את היעד החודשי. האחוז הינו היחס הרצוי שאנו רוצים להשיג. האחוז הוא היחס של כמות קישוריות השירות אל מול כמות שיחות השירות הנכנסות לאותו חודש</span></div>
																&nbsp;</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
								<tr><td height=40 align=right>&nbsp;יעד קישוריות שירות של שנה נוכחית (אל מול שיחות שירות) (באחוזים)&nbsp;</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
							<tr><td height=40 align=right>&nbsp;ביצוע קישוריות שנה נוכחית&nbsp;</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
							<tr><td height=40 align=right>&nbsp;ביצוע זמן אקטיב לנציג שירות שנה קודמת&nbsp;</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
						<tr><td height=40 align=right>&nbsp;יעד זמן אקטיב לנציג שירות שנה נוכחית&nbsp;</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
					<tr><td height=40 align=right>&nbsp;ביצוע זמן אקטיב לנציג שירות שנה נוכחית&nbsp;</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
							
							
							
									</table></td>
										
										<asp:Repeater ID=rptMonth Runat=server>
<ItemTemplate>
										<td valign=top width=6%>
											<table border=0 cellpadding=0 cellspacing=0 width=100%>
											<tr>
											<td align="center" class="title_sort" height=40 valign=middle><%#MonthName(Container.DataItem)%><td>
											</tr>
												<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
							<asp:Repeater ID=rptData Runat=server>
<ItemTemplate>	
									<tr><td height=40 align=center><%#Container.DataItem("Sales")%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									
<tr><td height=40 align=center valign=bottom>
									<table border=0 cellpadding=1 cellspacing=1>
										<tr><td height=35 align=center>
									<table border=0 cellpadding=1 cellspacing=1 style="background-color:#ffffff">
										<tr><td valign=middle Id="TdSales_Point_<%#Container.DataItem("Dim_Month")%>_<%=dep%>_<%=currYear%>" name="TdSales_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>" <%if PEdit="1" then%> onclick="javascript:OpenInput('Sales_Point','<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>')"<%end if%>>
									<div id="dvSales_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>" style="display:none;background-color:#ffffff;border:0px solid #ff0000;height:100%">
																		<input type=text dir="ltr" name="Sales_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>" id="Sales_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>" value="<%#Container.DataItem("Sales_Point")%>"  onkeypress='validate(event)' style="width:40px;margin: 0em;font-family:arial;display: block;" maxlength=100 onblur="javascript:changeInput('Sales_Point','<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>');" >
																	</div>
									<div id="dvTSales_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>"  style="display:block;width:40px;text-align:center;vertical-align: middle;background-color:#ffffff;">&nbsp;<%#Container.DataItem("Sales_Point")%>&nbsp;</div>
                             	</td></tr></table></td></tr>
									
									</table>
	</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%#IIF (Container.DataItem("Sales_Point")<>0 ,String.Format("{0:0.00}", (Container.DataItem("Sales")/Container.DataItem("Sales_Point"))*100),"")%>%</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%#Container.DataItem("SalesYear")%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%#IIF (Container.DataItem("SalesYear")<>0 ,String.Format("{0:0.00}", (Container.DataItem("Sales")/Container.DataItem("SalesYear"))*100),"")%>%</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%#Container.DataItem("Kishuriot")%></td></tr>
									
									
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center valign=bottom>
									<table border=0 cellpadding=1 cellspacing=1>
										<tr><td height=35 align=center>
									<table border=0 cellpadding=1 cellspacing=1 style="background-color:#ffffff">
											<tr><td  valign=middle Id="TdKishuritService_Point_<%#Container.DataItem("Dim_Month")%>_<%=dep%>_<%=currYear%>" name="TdKishuritService_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>" <%if PEdit="1" then%> onclick="javascript:OpenInput('KishuritService_Point','<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>')"<%end if%>>
									<div id="dvKishuritService_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>" style="display:none;background-color:#ffffff;border:0px solid #ff0000;height:100%" >
																		<input type=text dir="ltr" name="KishuritService_Point_<%#Container.DataItem("Dim_Month")%>_<%=dep%>_<%=currYear%>" id="KishuritService_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>" value="<%#Container.DataItem("KishuritService_Point")%>" onkeypress='validate(event)' style="width:40px;margin: 0em;font-family:arial;display: block;" maxlength=100 onblur="javascript:changeInput('KishuritService_Point','<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>');" >
																	</div>
									<div id="dvTKishuritService_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>"  style="display:block;width:40px;text-align:center;vertical-align: middle;background-color:#ffffff;"><%#Container.DataItem("KishuritService_Point")%></div>

									</td></tr></table></td></tr>
									
									</table>
																		</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%#Container.DataItem("KishuriotYear")%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%#Container.DataItem("ActiveTimeBefore")%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center valign=bottom>
									<table border=0 cellpadding=1 cellspacing=1>
										<tr><td height=35 align=center>
									<table border=0 cellpadding=1 cellspacing=1 style="background-color:#ffffff">
									<tr><td  valign=middle Id="TdActiveTime_Point_<%#Container.DataItem("Dim_Month")%>_<%=dep%>_<%=currYear%>" name="TdActiveTime_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>" <%if PEdit="1" then%> onclick="javascript:OpenInput('ActiveTime_Point','<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>')"<%end if%>>
									<div id="dvActiveTime_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>" style="display:none;background-color:#ffffff;border:0px solid #ff0000;height:100%" >
																		<input type=text dir="ltr" name="ActiveTime_Point_<%#Container.DataItem("Dim_Month")%>_<%=dep%>_<%=currYear%>" id="ActiveTime_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>" value="<%#Container.DataItem("ActiveTime_Point")%>" onkeypress='validate(event)' style="width:40px;margin: 0em;font-family:arial;display: block;" maxlength=100 onblur="javascript:changeInput('ActiveTime_Point','<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>');" >
																	</div>
									<div id="dvTActiveTime_Point_<%#trim(Container.DataItem("Dim_Month"))%>_<%=dep%>_<%=currYear%>"  style="display:block;width:40px;text-align:center;vertical-align: middle;background-color:#ffffff;"><%#Container.DataItem("ActiveTime_Point")%></div>

									</td></tr></table></td></tr>
							
									
									</table>
																		</td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%#Container.DataItem("ActiveTime")%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									</ItemTemplate>
								</asp:Repeater>
											
									</table></td>
								
								</ItemTemplate>
								</asp:Repeater>
											
										<td align="center" valign=top width=6% ><table border=0 cellpadding=0 cellspacing=0 width=100%>
											<tr>
											<td align="center" valign=middle height=40 class="title_sort">מצתבר שנתי<td>
											</tr>
													<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
							
									<tr><td height=40 align=center><%=SumSales%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%=SumSales_Point%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%=SumRow3%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%=String.Format("{0:0.00}",SumSalesYear)%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%if SumSales>0 then%><%=String.Format("{0:0.00}",(SumSalesYear/SumSales)*100)%><%end if%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%=SumKishuriot%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%=SumKishuritService_Point%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%=SumKishuriotYear%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%=SumActiveTime%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%=SumActiveTime_Point%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									<tr><td height=40 align=center><%=SumActiveTimeYear%></td></tr>
									<tr><td height=1 style="width:100%;background-color:#ffffff"></td></tr>
									</table></td>
										</TR>
								</table>
									</td></tr></table>
		</form>
	</body>
</html>

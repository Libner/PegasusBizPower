<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ViewScreen.aspx.vb" Inherits="bizpower_pegasus2018.ViewScreen"  validateRequest="false"%>
<%@ Register TagPrefix="DS" TagName="TopIn" Src="../../top_in.ascx" %>
<%@ Register TagPrefix="DS" TagName="LogoTop" Src="../../logo_top.ascx" %>
<%@ Register TagPrefix="DS" TagName="metaInc" Src="../../title_meta_inc.ascx" %>
<html>
	<head>
		<title>screenSetting</title>
		<DS:metaInc id="metaInc" runat="server"></DS:metaInc>
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
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
    width: 120px;
    background-color: #555;
    color: #fff;
    text-align: center;
    border-radius: 6px;
    padding: 5px 0;
    position: absolute;
    z-index: 1;
    bottom: 125%;
    left: 50%;
    margin-left: -60px;
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
		<script charset="utf-8">
		
function OpenData16719(dt,st)
{
var dP=format(dt)
//var dP=dt

if (st=="Sales")
{appeal_CallStatus=0
}
if (st=="Service")
{
appeal_CallStatus=1
}
/*inpsearch40729='Sales' //'לקוח לפני רישום'
alert(inpsearch40729)*/
//var param = { 'inpsearch40579' : dP,'inpsearch40729':inpsearch40729};
var param = { 'inpsearch40579' : dP};

OpenWindowWithPost("appealsDetails.asp?prodId=16719&appeal_CallStatus="+appeal_CallStatus+"&dt="+dP, "width=1000, height=600, left=100, top=100, resizable=yes, scrollbars=yes", "NewFile", param);

}
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
		</script>
		<script>

		
function OpenInputService(n,value)
{
//alert(n+'/'+value)
//$('form input').hide();
//$('div[id^="dvT"').show();

document.getElementById("dvT"+n+"_"+value).style.display="none";
document.getElementById("dv"+n+"_"+value).style.display="block";
document.getElementById(n+"_"+value).style.display = "block";
}
function format(d) {
var formatItems = d.split('/');
var dp
if (formatItems[0].length==1)
{dp="0"+formatItems[0]}
else
{dp=formatItems[0]}
if (formatItems[1].length==1)
{ dm="0"+formatItems[1]}
else
{dm=formatItems[1]}

return dp+"/"+dm+"/"+formatItems[2]

    }

function ChangeInputService(n,value,dep)
{
//alert(n)
//alert(value)
//update value-------
var  formData = "name="+n+"&Uid="+value+"&value="+document.getElementById(n+"_"+ value).value+"&dep="+dep; ; 
//alert (formData);
 
$.ajax({
    url : "UpdateDataAvrgDep.aspx",
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
	

//alert(n)

 document.getElementById("dv"+n+"_"+value).style.display="none";
 document.getElementById("dvT"+n+"_"+value).style.display="block";
 if (n=='ProcSalesCallMerkaz')
  document.getElementById("dvT"+n+"_"+value).innerHTML=document.getElementById(n+"_"+ value).value +'%';
 else
 document.getElementById("dvT"+n+"_"+value).innerHTML=document.getElementById(n+"_"+ value).value;

}


function OpenInput(n,value)
{
//alert(value)
//$('form input').hide();
//$('div[id^="dvT"').show();

document.getElementById("dvT"+n+"_"+value).style.display="none";
document.getElementById("dv"+n+"_"+value).style.display="block";

document.getElementById(n+"_"+value).style.display = "block";

}
function changeInput(n,value,dep)
{
//alert("changeInput")
//update value-------
//var dep=<%=dep%>;
//alert(dep)

	var  formData = "name="+n+"&Uid="+value+"&value="+document.getElementById(n+"_"+ value).value+"&dep="+dep; 
//alert (formData);
 
$.ajax({
    url : "UpdateData.aspx",
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
/*appeal_CallStatus="1"--פניית שירות 
	else פניית מכירה
*/
		</script>
	</head>
	<body style="margin:0px">
		<form id="Form1" method="post" runat="server" name="Form1">
			<DS:LOGOTOP id="logotop" runat="server"></DS:LOGOTOP>
			<DS:TOPIN id="topin" numOfLink="0" numOfTab="108" toplevel2="111" runat="server"></DS:TOPIN>
	
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
								<td height="30" align="center">
<span style="COLOR: #6F6DA6;font-size:14pt">מעקב ביצוע 
										מכירות ושירות כתמונת מצב יומית/חודשית</span></td>
							</tr>
							<tr>
								<td width="100%" style="color:#000000;background-color:#e1e1e1e1;">
									<table border="0" cellpadding="0" cellpadding="0" width="100%" style="color:#000000;background-color:#E6E6E6;">
										<tr>
											<td align="center">
												<table border="0" cellpadding="2" cellspacing="2">
													<tr>
														<td><input type="submit" value="הצג נתונים" class="but_menu" style="width:90;cursor:pointer"
																id="Button1" name="Button1"></td>
														<td>
															<select runat="server" id="seldep" class="searchList" dir="rtl" name="seldep" style="height:30px;direction:rtl;">
															</select>
															
														</td>
														<td align="center"><!--onchange="this.form.submit()"--><select class="searchList" id="y" name="y" style="height:30px">
																	<%for yy=2014 to Year(Now())+1%>
																	<option value="<%=yy%>" <%if currYear=yy then%>selected<%end if%> ><%=yy%></option>
																	<%next%>
																</select>
														</td>
														<td>
														<!--onchange="this.form.submit()"--><select class="searchList" id="m" name="m" style="height:30px;direction:rtl;">
																	<%for mm=1 to 12%>
																	<option value="<%=mm%>" <%if currMonth=mm then%>selected<%end if%> ><%=MonthName(mm)%></option>
																	<%next%>
																</select>
														
														</td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td valign="top" align="left">
									<table cellpadding="1" cellspacing="1" width="100%" align="left" style="border:solid 1px #d3d3d3">
										<tr>
											<td height="30" align="center"><span style="COLOR: #6F6DA6;font-size:14pt">נתונים 
													חודשיים</span></td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td valign="top" align="left">
									<table cellpadding="1" cellspacing="1" width="100%" align="left" style="border:solid 1px #d3d3d3">
										<tr>
											<td align="center"><span style="COLOR: #6F6DA6;font-size:14pt"> מכירות
													<%=monthName(currMonth)%>
													&nbsp;<%=currYear%>
												</span>
											</td>
										</tr>
										<tr>
											<td><table cellpadding="1" cellspacing="1" width="50%" align="center" style="border:solid 0px #d3d3d3">
												
													<Tr>
														<TD class="title_sort" align="center">זמן ממוצע של שיחת מכירה נכנסת</TD>
														<TD class="title_sort" align="center">אחוז המייצג את כמות זמן השיחה של צוות המכירות 
															ביחס לכמות הזמן בתוך מערכת המרכזיה</TD>
														<td  class="title_sort" align="center" dir=rtl><div  class="tooltip">אחוז טפסי המתעניין ביחס לפוטנציאל המכירות (הצגת אחוז)<span class="tooltiptext">שדה זה מציג את סך טפסי המתעניין שנפתחו עד כה ביחס לסך הפוטנציאל המכירתי הקיים עד כה</span></div></td>
														<TD class="title_sort" align="center"><div  class="tooltip">ממוצע יומי של טפסי מתעניין<span class="tooltiptext">דה זה מציג את ממוצע טפסי המתעניין היומי של חודש זה – החישוב מתבצע עד יום האתמול אלא אם החודש הסתיים</span></div></TD>
														<TD class="title_sort" align="center"><div class="tooltip">כמות מכירות חודשית<span class="tooltiptext">שדה זה מציג את הסכימה של השדה היומי "נטו רישום יומי" מתחילת החודש ועד ליום האתמול</span></div></TD>
														<TD class="title_sort" align="center"><div class="tooltip">ממוצע יומי של שיחות מכירה נכנסות <span class="tooltiptext">שדה זה מציג את ממוצע השיחות היומי של חודש זה – החישוב מתבצע עד יום האתמול אלא אם החודש הסתיים</span></div></TD>
														<TD class="title_sort" align="center">יעד מכירות חודשי</TD>
														<td class="title_sort" align="center" nowrap>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
													</Tr>
													
												<asp:Repeater id="rptTableSales" runat="server">
														<ItemTemplate>	
													<tr style="background-color: rgb(201, 201, 201);height:30px;">
																<td align=center style="background-color:#ffffff;border-top:solid 1px #c9c9c9;border-right:solid 1px #c9c9c9;border-left:solid 1px #c9c9c9" valign=middle id="TdTimeZmanSales_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" <%if PEdit="1" then%> onclick="javascript:OpenInputService('TimeZmanSales','<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>')"<%end if%>>
														<div id="dvTimeZmanSales_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" style="display:none;background-color:#ffffff">
															<input type=text dir="ltr" name="TimeZmanSales_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" id="TimeZmanSales_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" value="<%#Container.DataItem("TimeZmanSales")%>"  style="width:40px;font-family:arial" maxlength=100 onblur="javascript:ChangeInputService('TimeZmanSales','<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>','<%=dep%>');" >
														</div>
														<div id="dvTTimeZmanSales_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>"  style="display:block;padding-top:0px;background-color:#ffffff"><%#Container.Dataitem("TimeZmanSales")%>
														</div>
														</TD>
														<td align=center style="background-color:#ffffff;border-top:solid 1px #c9c9c9;border-right:solid 1px #c9c9c9;border-left:solid 1px #c9c9c9" valign=middle id="TdProcSalesCallMerkaz_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" <%if PEdit="1" then%> onclick="javascript:OpenInputService('ProcSalesCallMerkaz','<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>')"<%end if%>>
														<div id="dvProcSalesCallMerkaz_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" style="display:none;background-color:#ffffff">
															<input type=text dir="ltr" name="ProcSalesCallMerkaz_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" id="ProcSalesCallMerkaz_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" value="<%#Container.DataItem("ProcSalesCallMerkaz")%>"  style="width:40px;font-family:arial" maxlength=100 onblur="javascript:ChangeInputService('ProcSalesCallMerkaz','<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>','<%=dep%>');" >
														</div>
														<div id="dvTProcSalesCallMerkaz_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>"  style="display:block;padding-top:0px;background-color:#ffffff"><%#Container.Dataitem("ProcSalesCallMerkaz")%>%
														</div>
														</TD>
														<TD align="center">&nbsp;<%#String.Format("{0:0.0}",func.PotencialTo16504Sales(Container.Dataitem("SMonth"),Container.Dataitem("SYear"),dep))%><%'if sumPotencialSales>0 then%><%'=String.Format("{0:0}",(sumP16504/sumPotencialSales)*100)%>%<%'end if%>&nbsp;</TD>
														<TD align="center">&nbsp;<%#func.Neto16504(Container.Dataitem("SMonth"),Container.Dataitem("SYear"))%><%'=sumP16504/MonthDays%>&nbsp;</TD>
														<TD align="center">&nbsp;<%#func.Neto16735(Container.Dataitem("SMonth"),Container.Dataitem("SYear"),dep)%><%'=sum16735Days%>&nbsp;</TD>
														<TD align="center">&nbsp;<%#func.AvrgCallSales(Container.Dataitem("SMonth"),Container.Dataitem("SYear"),dep)%><%'=String.Format("{0:0}",sumcall_Sales/MonthDays)%>&nbsp;</TD>
														<TD align="center">&nbsp;<%#func.Sales_Point(Container.Dataitem("SMonth"),Container.Dataitem("SYear"),dep)%>&nbsp;</TD>
														<td align="center" nowrap>&nbsp;<%#Container.Dataitem("SMonth")%>/<%#Container.Dataitem("SYear")%>&nbsp;</td>
													</tr>
													</ItemTemplate>
													</asp:Repeater>
												
												</table>
											</td>
										</tr>
										<tr>
											<td><table cellpadding="0" cellspacing="0" width="50%" align="center" style="border:solid 1px #d3d3d3">
													<tr><td></td></tr></table></td>
									</tr>
										<tr>
											<td height="30"></td>
										</tr>
										
										<tr>
											<td><table cellpadding="1" cellspacing="1" width="50%" align="center" style="border:solid 0px #d3d3d3">
													<tr>
														<td class="title_sort" align="center"><div class="tooltip">ממוצע קצב מכירות יומי לחודש זה<span class="tooltiptext">ממוצע המכירות היומי עד אתמול כולל </span></div></td>
														<td class="title_sort" align="center"><div class="tooltip">צפי ביצוע עד לסוף החודש<span class="tooltiptext">צפי הביצוע בסיום חודש זה בהסתמך על הקצב עד כה ועל סך ימי העבודה הנותרים</span></div></td>
														<td class="title_sort" align="center">כמות ימי עבודה חודשית שבוצעו עד כה</td>
														<td class="title_sort" align="center"><div class="tooltip">אחוז ביצוע חודשי אל מול יעד חודשי<span class="tooltiptext">שדה זה מציג את היחס בין היעד לבין הביצוע בפועל של חודש זה – החישוב מתבצע עד יום האתמול אלא אם החודש הסתיים</span></div></td>
													</tr>
													<tr style="background-color:#E6E6E6">
														<td align="center">&nbsp;<%If MonthDays>0 then %><%=String.Format("{0:0.0}",(sum16735Days/MonthDays))%><%end if%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
														<td align="center">&nbsp;<%If MonthDays>0 then %><%=String.Format("{0:0.0}",(sum16735Days/MonthDays)*sumDays)%><%end if%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
														<td align="center">&nbsp;<%If MonthDays>0 then %><%=String.Format("{0:0.0}",MonthDays)%><%end if%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
														<td align="center">&nbsp;<%'=func.Sales_Point(currMonth,currYear,dep)%><%if func.Sales_Point(currMonth,currYear,dep)>0 then%><%=String.Format("{0:0.0}",(sum16735Days/func.Sales_Point(currMonth,currYear,dep))*100)%>%<%end if%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
													</tr>
												</table>
											</td>
										</tr>
										<TR>
											<td height="10"></td>
										</TR>
										<tr>
											<td align="center"><span style="COLOR: #6F6DA6;font-size:14pt"> שירות
													<%=monthName(currMonth)%>
													&nbsp;<%=currYear%>
												</span>
											</td>
										</tr>
										<tr>
											<td><table cellpadding="1" cellspacing="1" width="50%" align="center" style="border:solid 0px #d3d3d3">
													<Tr>
														<TD class="title_sort" align="center">ממוצע זמן שיחת שירות נכנסת</TD>
														<TD class="title_sort" align="center">ממוצע זמן שיחה בו נציג מדבר באופן אקטיבי עם 
															לקוח</TD>
														<TD class="title_sort" align="center">
															<div class="tooltip">(%)היחס בין כמות קישוריות השירות לעומת כמות שיחות השירות 
																שנכנסה <span class="tooltiptext">אחוז אשר מראה את היחס הממוצע בחודש זה בין קישוריות 
																	שירות לבין שיחות שירות</span>
															</div>
														</TD>
														<TD class="title_sort" align="center">ממוצע יומי של קישוריות שירות</TD>
														<TD class="title_sort" align="center">ממוצע יומי של כמות שיחות שירות נכנסות</TD>
														<td class="title_sort" align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
													</Tr>
													<asp:Repeater id="rptTableService" runat="server">
														<ItemTemplate>
															<tr style="background-color: rgb(201, 201, 201);height:30px;">
																<td align=center style="background-color:#ffffff;border-top:solid 1px #c9c9c9;border-right:solid 1px #c9c9c9;border-left:solid 1px #c9c9c9" valign=middle id="TdAvrg_Call_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" <%if PEdit="1" then%> onclick="javascript:OpenInputService('Avrg_Call','<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>')"<%end if%>>
																	<div id="dvAvrg_Call_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" style="display:none;background-color:#ffffff">
																		<input type=text dir="ltr" name="Avrg_Call_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" id="Avrg_Call_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" value="<%#Container.Dataitem("Avrg_Call")%>"  style="width:40px;font-family:arial" maxlength=100 onblur="javascript:ChangeInputService('Avrg_Call','<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>','<%=dep%>');" >
																	</div>
																	<div id="dvTAvrg_Call_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>"  style="display:block;padding-top:0px;background-color:#ffffff"><%#String.Format("{0:0}",Container.Dataitem("Avrg_Call"))%></div>
																</td>
																<td align=center style="background-color:#ffffff;border-top:solid 1px #c9c9c9" valign=middle id="TdAvrg_CallProc_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>"  <%if PEdit="1" then%> onclick="javascript:OpenInputService('Avrg_CallProc','<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>')"<%end if%>>
																	<div id="dvAvrg_CallProc_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" style="display:none;background-color:#ffffff">
																		<input type=text dir="ltr" name="Avrg_CallProc_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" id="Avrg_CallProc_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>" value="<%#Container.Dataitem("Avrg_CallProc")%>"  style="width:40px;font-family:arial" maxlength=100 onblur="javascript:ChangeInputService('Avrg_CallProc','<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>','<%=dep%>');" >
																	</div>
																	<div id="dvTAvrg_CallProc_<%#Container.Dataitem("SMonth")%>_<%#Container.Dataitem("SYear")%>"  style="display:block;padding-top:0px;background-color:#ffffff"><%#String.Format("{0:0}",Container.Dataitem("Avrg_CallProc"))%>
																	</div>
																</td>
																<TD align="center"><%'#String.Format("{0:0.0}",func.AvrgService(Container.Dataitem("SMonth"),Container.Dataitem("SYear"),dep))%>
																<%'''#func.AvrgCallService(Container.Dataitem("SMonth"),Container.Dataitem("SYear"),dep)%>
																<%'#func.Kishuriot16719Service(Container.Dataitem("SMonth"),Container.Dataitem("SYear"))%><%'#func.AvrgCallService(Container.Dataitem("SMonth"),Container.Dataitem("SYear"),dep)%>
																<%'#func.Kishuriot16719Service(Container.Dataitem("SMonth"),Container.Dataitem("SYear"))%>
																<%'#func.AvrgCallService(Container.Dataitem("SMonth"),Container.Dataitem("SYear"),dep)%>
																
																<%#func.ProcCallService(func.Kishuriot16719Service(Container.Dataitem("SMonth"),Container.Dataitem("SYear")),func.AvrgCallService(Container.Dataitem("SMonth"),Container.Dataitem("SYear"),dep))%>
																
																<%'#IIF (cint(func.AvrgCallService(Container.Dataitem("SMonth"),Container.Dataitem("SYear"),dep))>0,(func.Kishuriot16719Service(Container.Dataitem("SMonth"),Container.Dataitem("SYear"))/func.AvrgCallService(Container.Dataitem("SMonth"),Container.Dataitem("SYear"),dep))*100,"")%>




																<%'if sumCallService>0 then%>
																<%'=String.Format("{0:0.0}",(sumP16719Kishurit_Service/sumCallService)*100)%>
																%<%'end if%></TD>
																
																<TD align="center">&nbsp;<%#func.Kishuriot16719Service(Container.Dataitem("SMonth"),Container.Dataitem("SYear"))%><asp:label id="LblKishurit" runat=server ></asp:label><%'=resSP16719Kishurit_Service%>
																<%'#IIF (datediff("d","01/"& Container.Dataitem("SMonth")&"/"& Container.Dataitem("SYear"),Now())<MonthDays ,Container.Dataitem("SP16719Kishurit_Service")/datediff("d","01/"& Container.Dataitem("SMonth")&"/"& Container.Dataitem("SYear"),Now()),Container.Dataitem("SP16719Kishurit_Service")/DateTime.DaysInMonth(currYear, currMonth))%>&nbsp;</TD>
																<TD align="center"><%#func.AvrgCallService(Container.Dataitem("SMonth"),Container.Dataitem("SYear"),dep)%>
																<%'#String.Format("{0:0.0}",sumCallService/MonthDays)%>
																<%'#IIF (datediff("d","01/"& Container.Dataitem("SMonth")&"/"& Container.Dataitem("SYear"),Now())<MonthDays ,String.Format("{0:0.0}",Container.Dataitem("Scall_Service")/MonthDaysNow),String.Format("{0:0.0}",Container.Dataitem("Scall_Service")/MonthDays))%>
																&nbsp;</TD>
																<td align="center" nowrap>&nbsp;&nbsp;<%#Container.Dataitem("SMonth")%>/<%#Container.Dataitem("SYear")%>&nbsp;&nbsp;</td>
															</tr>
														</ItemTemplate>
													</asp:Repeater>
												</table>
											</td>
										</tr>
									</table>
								</td>
							</tr>
							<!--  <div id="dvCustomers">-->
							<!--paging-->
							<tr>
								<td valign="top" align="left" height="20">
								<td>
							</tr>
							<tr>
								<td valign="top" align="left">
									<table cellpadding="1" cellspacing="1" width="100%" align="left" style="border:solid 0px #d3d3d3">
										<tr>
											<td height="30" align="center"><span style="COLOR: #6F6DA6;font-size:14pt">טבלת מעקב ושליטה יומית</span></td>
										</tr>
<tr>					<td>
	<!--navigation-->
		<table align="right" cellpadding="1" cellspacing="1" width="100%" style="border:solid 1px #d3d3d3">
													<tr>
														<td class="title_sort" align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<div class="tooltip"><B>סה"כ
																	<BR>
																	צור קשר</B> <span class="tooltiptext">כמות יומית של סך טפסי צור קשר + טפסי צור 
																	קשר מקוצר שנפתחו ב CRM</span></div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip"><B>סה"כ קישוריות<BR>
																	מכירה + שירות</B> <span class="tooltiptext">כמות יומית של סך טפסי הקישוריות 
																	שנפתחו ב CRM</span></div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip"><B>אחוז קישוריות
																	<BR>
																	יומי שירות</B> <span class="tooltiptext">היחס שבין כמות הקישוריות שירות שהתקבלה 
																	ביום מסויים אל מול כמות שיחות השירות הנכנסות לאותו היום</span>
															</div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip"><B>סה"כ פוטנציאל<BR>
																	שירות</B> <span class="tooltiptext">כמות הפוטנציאל השירותי באגף השירות כמות זו 
																	מחושבת על בסיס קישוריות שירות, שיחות שירות, צור קשר שירות במערכת</span>
															</div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip"><B>שיחות שירות
																	<BR>
																	נכנסות</B> <span class="tooltiptext">כמות שיחות המכירה שנכנסו לנתב השיחות של 
																	המרכזיה במהלך היום</span></div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip"><B>צור קשר
																	<BR>
																	שירות</B> <span class="tooltiptext">כמות יומית של טפסי צור קשר + טפסי צור קשר 
																	מקוצר שהתקבלה ב CRM נכון ליום הרלוונטי העוסקים בשירות</span>
															</div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip"><B>קישוריות<BR>
																	שירות</B> <span class="tooltiptext">כמות יומית של סך טפסי קישוריות השירות אשר 
																	נסכם לפי השדה לקוח אחרי רישום</span>
															</div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip"><B>אחוז קישוריות
																	<BR>
																	יומי מכירה</B> <span class="tooltiptext">היחס שבין כמות הקישוריות מכירה שהתקבלה 
																	ביום מסויים אל מול כמות שיחות המכירה הנכנסות לאותו היום</span>
															</div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip"><B>נטו<BR>
																	רישום יומי</B> <span class="tooltiptext">כמות יומית של טפסי הרישום החתום לאחר 
																	קיזוז הטפסים של אותו יום שבוטלו- נכון ליום הרלוונטי</span>
															</div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip"><B>אחוז
																	<BR>
																	הצלחה יומי</B> <span class="tooltiptext">"חישוב של סך כמות המטיילים בטפסי 
																	הרישום החתום חלקי סך כמות טפסי המתעניין שנפתחה ביום מסויים כפול 1.8"</span>
															</div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip"><B>מבוטלים</B> <span class="tooltiptext">
																	כמות יומית של טפסי הרישום החתום שבוטלו ב CRM נכון ליום הרלוונטי</span></div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip"><B>טפסי
																	<BR>
																	רישום חתום</B> <span class="tooltiptext">כמות יומית של טפסי הרישום החתום שנפתחה 
																	ב CRM נכון ליום הרלוונטי</span></div>
														</td>
														<td class="title_sort" align="center"><B>סה"כ טפסי
																<BR>
																מתעניין</B></td>
														<td class="title_sort" align="center"><div class="tooltip"><B>סה"כ פוטנציאל
																	<BR>
																	מכירות</B> <span class="tooltiptext">כמות הפוטנציאל המכירתי באגף המכירות כמות 
																	זו מחושבת על בסיס קישוריות המכירה, שיחות המכירה, צור קשר המכירה במערכת</span>
															</div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip"><B>שיחות<BR>
																	מכירה נכנסות</B> <span class="tooltiptext">כמות שיחות המכירה שנכנסו לנתב השיחות 
																	של המרכזיה במהלך היום</span></div>
														</td>
														<td class="title_sort" align="center"><div class="tooltip">
																<B>צור קשר<BR>
																	מכירה</B> <span class="tooltiptext">כמות יומית של טפסי צור קשר + טפסי צור קשר 
																	מקוצר שהתקבלה ב CRM נכון ליום הרלוונטי העוסקים במכירה</span></div>
														</td>
														<td class="title_sort" align="center">
															<div class="tooltip"><B>קישוריות<BR>
																	מכירה</B> <span class="tooltiptext">כמות יומית של סך טפסי קישוריות המכירה אשר 
																	נסכם לפי השדה "לקוח לפני רישום</span>
															</div>
														</td>
														<td class="title_sort" align="center"><B>תאריך</B></td>
														<td class="title_sort" align="center"><B>יום<BR>
																עבודה</B></td>
														<td class="title_sort" align="center"><B>מספר<BR>
																שורה</B></td>
													</tr>
													<asp:Repeater ID="rptDays" Runat="server">
														<ItemTemplate>
															<tr style="background-color: rgb(201, 201, 201);height:30px;" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';" style="background-color: rgb(201, 201, 201);">
																<td align="center"><%#Container.DataItem("P16724ContactUs_Service")+Container.DataItem("P16724ContactUs_Sales")+Container.DataItem("P17012ContactUs_Service")+Container.DataItem("P17012ContactUs_Sales")%></td>
																<td align=center style="cursor: pointer;" onclick="javascript:OpenData16719('<%#day(Container.DataItem("DateKey"))%>/<%#month(Container.DataItem("DateKey"))%>/<%#year(Container.DataItem("DateKey"))%>','Service')"><%#Container.DataItem("P16719Kishurit_Service")+Container.DataItem("P16719Kishurit_Sales")%></td>
																<td align="center"><%#IIF(Container.DataItem("call_Service")>0 ,String.Format("{0:0.0}",Container.DataItem("P16719Kishurit_Service")/Container.DataItem("call_Service")*100) &"%","")%></td>
																<td align="center"><%#Container.DataItem("P16719Kishurit_Service")+Container.DataItem("P16724ContactUs_Service")+Container.DataItem("P17012ContactUs_Service")+Container.DataItem("call_Service")%></td>
																<td align=center style="background-color:#ffffff;border-top:solid 1px #c9c9c9" valign=middle id="Tdcall_Service_<%#Container.DataItem("DateKey")%>"  <%if PEdit="1" then%> onclick="javascript:OpenInput('call_Service','<%#Container.DataItem("DateKey")%>')"<%end if%>>
																	<div id="dvcall_Service_<%#Container.DataItem("DateKey")%>" style="display:none;background-color:#ffffff">
																		<input type=text dir="ltr" name="call_Service_<%#Container.DataItem("DateKey")%>" id="call_Service_<%#Container.DataItem("DateKey")%>" value="<%#Container.DataItem("call_Service")%>"  style="width:40px;font-family:arial" maxlength=100 onblur="javascript:changeInput('call_Service','<%#Container.DataItem("DateKey")%>','<%=dep%>');" >
																	</div>
																	<div id="dvTcall_Service_<%#Container.DataItem("DateKey")%>"  style="display:block;padding-top:0px;background-color:#ffffff"><%#Container.DataItem("call_Service")%></div>
																</td>
																	<td align="center"><%#Container.DataItem("P16724ContactUs_Service")+Container.DataItem("P17012ContactUs_Service")%></td>
																<td align="center" style="cursor: pointer;" onclick="javascript:OpenData16719('<%#day(Container.DataItem("DateKey"))%>/<%#month(Container.DataItem("DateKey"))%>/<%#year(Container.DataItem("DateKey"))%>','Service')"><%#Container.DataItem("P16719Kishurit_Service")%></td>
																<td align="center"><%#IIF(Container.DataItem("call_Sales")>0 ,String.Format("{0:0.0}",Container.DataItem("P16719Kishurit_Sales")/Container.DataItem("call_Sales")*100)&"%","")%></td>
																<td align=center style="cursor: pointer;" onclick="window.open('appealsDetails16735.asp?prodId=16735&dep=<%=dep%>&dt=<%#day(Container.DataItem("DateKey"))%>/<%#month(Container.DataItem("DateKey"))%>/<%#year(Container.DataItem("DateKey"))%>','winApp','top=20, left=10, width=1800, height=700, scrollbars=1');"><%#Container.DataItem("P16735")-Container.DataItem("P16735_Bitulim")%></td>
																<td align="center"><%#IIF (Container.DataItem("P16504")=0 ,"", String.Format("{0:0.0}",100*cint(Container.DataItem("P16735"))/(cint(Container.DataItem("P16504"))*1.8))&"%")%></td>
																<td align=center style="cursor: pointer;" onclick="window.open('appealsDetailsBitulim16735.asp?prodId=16735&dep=<%=dep%>&dt=<%#day(Container.DataItem("DateKey"))%>/<%#month(Container.DataItem("DateKey"))%>/<%#year(Container.DataItem("DateKey"))%>','winApp','top=20, left=10, width=1800, height=700, scrollbars=1');"><%#Container.DataItem("P16735_Bitulim")%></td>
																<td align=center style="cursor: pointer;" onclick="window.open('appealsDetails16735.asp?prodId=16735&dep=<%=dep%>&dt=<%#day(Container.DataItem("DateKey"))%>/<%#month(Container.DataItem("DateKey"))%>/<%#year(Container.DataItem("DateKey"))%>','winApp','top=20, left=10, width=1800, height=700, scrollbars=1');"><%#Container.DataItem("P16735")%></td>
																<td align="center"><%#Container.DataItem("P16504")%></td>
																<td align="center"><%#Container.DataItem("P16719Kishurit_Sales")+Container.DataItem("P16724ContactUs_Sales")+Container.DataItem("P17012ContactUs_Sales")+Container.DataItem("call_Sales")%></td>
																<td align=center style="background-color:#ffffff;border-top:solid 1px #c9c9c9" valign=middle id="Tdcall_Sales_<%#Container.DataItem("DateKey")%>"  <%if PEdit="1" then%> onclick="javascript:OpenInput('call_Sales','<%#Container.DataItem("DateKey")%>')"<%end if%>>
																	<div id="dvcall_Sales_<%#Container.DataItem("DateKey")%>" style="display:none;background-color:#ffffff">
																		<input type=text dir="ltr" name="call_Sales_<%#Container.DataItem("DateKey")%>" id="call_Sales_<%#Container.DataItem("DateKey")%>" value="<%#Container.DataItem("call_Sales")%>"  style="width:40px;font-family:arial" maxlength=100 onblur="javascript:changeInput('call_Sales','<%#Container.DataItem("DateKey")%>','<%=dep%>');" >
																	</div>
																	<div id="dvTcall_Sales_<%#Container.DataItem("DateKey")%>"  style="display:block;padding-top:0px;background-color:#ffffff"><%#Container.DataItem("call_Sales")%></div>
																</td>
																<td align="center"><%#Container.DataItem("P16724ContactUs_Sales")+Container.DataItem("P17012ContactUs_Sales")%></td>
																<td align="center" style="cursor: pointer;" onclick="javascript:OpenData16719('<%#day(Container.DataItem("DateKey"))%>/<%#month(Container.DataItem("DateKey"))%>/<%#year(Container.DataItem("DateKey"))%>','Sales')"><%#Container.DataItem("P16719Kishurit_Sales")%></td>
																<td align="center"><%#day(Container.DataItem("DateKey"))%>/<%#month(Container.DataItem("DateKey"))%>/<%#year(Container.DataItem("DateKey"))%></td>
																<td align="center"><%#Container.DataItem("DayTypeId")%></td>
																<td align="center"><%#Container.ItemIndex+1%></td>
															</tr>
														</ItemTemplate>
														<FooterTemplate>
															<tr style="background-color: #6F6DA6;height:30px;">
																<td align=center><b><%=sumP16724ContactUs%></b></td>
																<td align=center><b><%=sumP16719Kishurit%></b></td>
																<td align=center><b><%if sumCallService>0 then%><%=String.Format("{0:0.0}",(sumP16719Kishurit_Service / sumCallService) * 100)%>%<%end if%><%'=String.Format("{0:0.0}",sumProcKishurit_Service)%></b></td>
																<td align=center><b><%=sumPotencialSevice%></b></td>
																<td align=center><b><%=sumCallService%></b></td>
																<td align=center><b><%=sumContactUs_Service%></b></td>
																<td align=center><b><%=sumP16719Kishurit_Service%></b></td>
																<td align=center><b><%if sumcall_Sales>0 then%><%=String.Format("{0:0.0}",(sumP16719Kishurit_Sales/sumcall_Sales) * 100 )%>%<%end if%><%'sumProcKishurit_Sales%></b></td>
																<td align=center><b><%=sum16735Days%></b></td>
																<td align=center><b><%if sumP16504>0 then%><%=String.Format("{0:0.0}", 100 * CInt(sum16735Days) / (CInt(sumP16504) * 1.8))%>%<%end if%><%'%aclahaYomi=String.Format("{0:0.0}",sumProcAzlaha)%></b></td>
																<td align=center><b><%=sumP16735_Bitulim%></b></td>
																<td align=center><b><%=sumP16735%></b></td>
																<td align=center><b><%=sumP16504%></b></td>
																<td align=center><b><%=sumPotencialSales%></b></td>
																<td align =center><b><%=sumcall_Sales%></b></td>
																<td align=center><b><%=sumP16724ContactUs_Sales%></b></td>
																<td align=center><b><%=sumP16719Kishurit_Sales%></b></td>
																<td>&nbsp;&nbsp;</td>
																<td align=center>&nbsp;<b><%=sumDays%></b>&nbsp;</td>
																<td align="center"><b>סכימה<br>
																		חודשית</b></td>
															</tr>
														</FooterTemplate>
													</asp:Repeater>
												</table>
											</td>
										</tr>
									</table>
									<SCRIPT type="text/javascript">
            <!--
            var cal1xx = new CalendarPopup('CalendarDiv');
                cal1xx.setYearSelectStartOffset(120);
	cal1xx.setYearSelectEndOffset(2);
	//cal1xx.setYearSelectStart(1910); //Added by Mila
                cal1xx.showNavigationDropdowns();
                cal1xx.offsetX = -50;
                cal1xx.offsetY = -40;
          
                //-->
									</SCRIPT>
									<DIV ID='CalendarDiv' STYLE='POSITION:absolute;VISIBILITY:hidden;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>
		</form>
	</body>
</html>

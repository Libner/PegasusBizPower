<%@ Page Language="vb" AutoEventWireup="false" Codebehind="WorkScreen.aspx.vb" Inherits="bizpower_pegasus.WorkScreen" %>
<!DOCTYPE html>

<html >
  <head>
    <title>screenSetting</title>
     <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
	<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
	<script src="../../include/CalendarPopup.js" charset="windows-1255"></script>
<script>
function ClearAll()
{
/*document.getElementById("sUsers").value=0;
document.getElementById("Departments").value=0;
document.getElementById("Series").value=0;
sFromD=$("#sPayFromDate").val()
sToD=$("#sPayToDate").val()
sPage=$("#pageList").val()
sStatus=$("#sStatus").val()*/
window.location ="WorkScreen.aspx"

} function replaceAll(find, replace, str) 
    {
      while( str.indexOf(find) > -1)
      {
        str = str.replace(find, replace);
      }
      return str;
    }
function edit_row(no)
{
 document.getElementById("edit_button"+no).style.display="none";
document.getElementById("save_button"+no).style.display="block";
//document.getElementById("cancel_button"+no).style.display="block";
	
 //var name=document.getElementById("FinancialEuro_row"+no);
 
// var name_data=name.innerHTML;
//var n=document.getElementById("Departure_GuideTelphone_row"+no).innerHTML
//var y=replaceAll("<br>"," ",n)
//alert(n)
//alert(y)
 document.getElementById("TimeMeetingAfterTrip_row"+no).innerHTML="<input style='width:35px' type='text' id='TimeMeetingAfterTrip_text"+no+"' value='"+document.getElementById("TimeMeetingAfterTrip_row"+no).innerHTML+"'>";
 document.getElementById("FinancialEuro_row"+no).innerHTML="<input style='width:65px' type='text' dir='ltr' id='FinancialEuro_text"+no+"' value='"+document.getElementById("FinancialEuro_row"+no).innerHTML+"'>";
 document.getElementById("Departure_GuideTelphone_row"+no).innerHTML="<textarea type='text' dir='ltr' id='Departure_GuideTelphone_text"+no+"' style='width:100px;height:60px'>" + replaceAll("<br>","\n",document.getElementById("Departure_GuideTelphone_row"+no).innerHTML)+"</textarea>";
 
 
 document.getElementById("FinancialDolar_row"+no).innerHTML="<input style='width:65px' type='text' dir='ltr' id='FinancialDolar_text"+no+"' value='"+document.getElementById("FinancialDolar_row"+no).innerHTML+"'>";
 document.getElementById("DepartureTimeBrief_row"+no).innerHTML="<input style='width:35px' dir='ltr' type='text' id='DepartureTimeBrief_text"+no+"' value='"+document.getElementById("DepartureTimeBrief_row"+no).innerHTML+"'>";
 document.getElementById("Departure_TimeGroupMeeting_row"+no).innerHTML="<input style='width:35px' dir='ltr' type='text' id='Departure_TimeGroupMeeting_text"+no+"' value='"+document.getElementById("Departure_TimeGroupMeeting_row"+no).innerHTML+"'>";
 //document.getElementById("Departure_DateBrief_row"+no).innerHTML="<input style='width:50px' dir='ltr' type='text' id='Departure_DateBrief_text"+no+"' value='"+document.getElementById("Departure_DateBrief_row"+no).innerHTML+"'>";
 // document.getElementById("GilboaHotel_row"+no).innerHTML="<select id='selGilboaHotel"+no+"'><option value=''>בחר</option><option value='כן' >כן</option></select>";

 //alert (document.getElementById("GilboaHotel_row"+no).innerHTML)
if (document.getElementById("GilboaHotel_row"+no).innerHTML=='כן')
 {
 document.getElementById("GilboaHotel_row"+no).innerHTML="<select id='selGilboaHotel"+no+"' style='width:45px'><option value=''>בחר</option><option value='כן' selected >כן</option></select>";
 }
 else
 {
   document.getElementById("GilboaHotel_row"+no).innerHTML="<select id='selGilboaHotel"+no+"' style='width:45px'><option value=''>בחר</option><option value='כן' >כן</option></select>";
 }
 
 
  if (document.getElementById("Vouchers_Provider_row"+no).innerHTML=='כן')
 {
 document.getElementById("Vouchers_Provider_row"+no).innerHTML="<select id='selVouchers_Provider"+no+"' style='width:60px'><option value=''>בחר</option><option value='כן' selected>כן</option><option value='מותאם'>מותאם</option></select>";
 }
   if (document.getElementById("Vouchers_Provider_row"+no).innerHTML=='מותאם')
 {
 document.getElementById("Vouchers_Provider_row"+no).innerHTML="<select id='selVouchers_Provider"+no+"' style='width:60px'><option value=''>בחר</option><option value='כן' >כן</option><option value='מותאם' selected>מותאם</option></select>";
 }
 else
 {
   document.getElementById("Vouchers_Provider_row"+no).innerHTML="<select id='selVouchers_Provider"+no+"' style='width:60px'><option value=''>בחר</option><option value='כן' >כן</option><option value='מותאם'>מותאם</option></select>";
 }
 
 
 if (document.getElementById("Voucher_Simultaneous_row"+no).innerHTML=='כן')
 {
 document.getElementById("Voucher_Simultaneous_row"+no).innerHTML="<select id='selVoucher_Simultaneous"+no+"' style='width:45px'><option value=''>בחר</option><option value='כן' selected>כן</option></select>";
 }
 else
 {
   document.getElementById("Voucher_Simultaneous_row"+no).innerHTML="<select id='selVoucher_Simultaneous"+no+"' style='width:45px'><option value=''>בחר</option><option value='כן' >כן</option></select>";
 }
if (document.getElementById("Departure_Costing_row"+no).innerHTML=='כן')
 {
 document.getElementById("Departure_Costing_row"+no).innerHTML="<select id='selDeparture_Costing"+no+"' style='width:45px'><option value=''>בחר</option><option value='כן' selected>כן</option></select>";
 }
 else
 {
   document.getElementById("Departure_Costing_row"+no).innerHTML="<select id='selDeparture_Costing"+no+"' style='width:45px'><option value=''>בחר</option><option value='כן' >כן</option></select>";
 }
 if (document.getElementById("Voucher_Group_row"+no).innerHTML=='כן')
 {
 document.getElementById("Voucher_Group_row"+no).innerHTML="<select id='selVoucher_Group"+no+"' style='width:45px'><option value=''>בחר</option><option value='כן' selected>כן</option></select>";
 }
 else
 {
   document.getElementById("Voucher_Group_row"+no).innerHTML="<select id='selVoucher_Group"+no+"' style='width:45px'><option value=''>בחר</option><option value='כן' >כן</option></select>";
 }
 
 
 //$("#GilboaHotel"+no).val('כן');

//alert(document.getElementById("GilboaHotel_row"+no).innerHTML)
}
function save_row(no)
{
	
 $.ajax
 ({
  type:'post',
  url:'modify_records.aspx',
  data:{
   edit_row:'edit_row',
   row_id:no,
   Departure_GuideTelphone_val:document.getElementById("Departure_GuideTelphone_text"+no).value,
   TimeMeetingAfterTrip_val:document.getElementById("TimeMeetingAfterTrip_text"+no).value,
   FinancialEuro_val:document.getElementById("FinancialEuro_text"+no).value,
   FinancialDolar_val:document.getElementById("FinancialDolar_text"+no).value,
   Departure_TimeGroupMeeting_val:document.getElementById("Departure_TimeGroupMeeting_text"+no).value,
   DepartureTimeBrief_val:document.getElementById("DepartureTimeBrief_text"+no).value,
   GilboaHotel_val:document.getElementById("selGilboaHotel"+no).value,
   Departure_Costing_val:document.getElementById("selDeparture_Costing"+no).value,
   Voucher_Group_val:document.getElementById("selVoucher_Group"+no).value,
   Vouchers_Provider_val:document.getElementById("selVouchers_Provider"+no).value,
   Voucher_Simultaneous_val:document.getElementById("selVoucher_Simultaneous"+no).value
   
     },
  success:function(response) {
  //alert (response)
   if(response=="success")
   {

   }
  }
 });
 document.getElementById("TimeMeetingAfterTrip_row"+no).innerHTML=document.getElementById("TimeMeetingAfterTrip_text"+no).value;
 document.getElementById("FinancialEuro_row"+no).innerHTML=document.getElementById("FinancialEuro_text"+no).value;
 document.getElementById("FinancialDolar_row"+no).innerHTML=document.getElementById("FinancialDolar_text"+no).value;
 document.getElementById("DepartureTimeBrief_row"+no).innerHTML=document.getElementById("DepartureTimeBrief_text"+no).value;
 document.getElementById("Departure_TimeGroupMeeting_row"+no).innerHTML=document.getElementById("Departure_TimeGroupMeeting_text"+no).value;
 document.getElementById("GilboaHotel_row"+no).innerHTML=document.getElementById("selGilboaHotel"+no).value;
 document.getElementById("Voucher_Simultaneous_row"+no).innerHTML=document.getElementById("selVoucher_Simultaneous"+no).value;
 document.getElementById("Departure_Costing_row"+no).innerHTML=document.getElementById("selDeparture_Costing"+no).value;
 document.getElementById("Voucher_Group_row"+no).innerHTML=document.getElementById("selVoucher_Group"+no).value;
 document.getElementById("Vouchers_Provider_row"+no).innerHTML=document.getElementById("selVouchers_Provider"+no).value;
 
 document.getElementById("Departure_GuideTelphone_row"+no).innerHTML=document.getElementById("Departure_GuideTelphone_text"+no).value;




 document.getElementById("edit_button"+no).style.display="block";
 document.getElementById("save_button"+no).style.display="none";
}
/*function cancel_row(no)
{
document.getElementById("edit_button"+no).style.display="block";
document.getElementById("save_button"+no).style.display="none";
document.getElementById("cancel_button"+no).style.display="none";
document.getElementById("FinancialEuro_text"+no).innerHTML=document.getElementById("FinancialEuro_text"+no).value;

}*/

	function openDeparture(ID)
	{
		h = 200;
		w = 800;
		S_Wind = window.open("UpdDep.asp?ID=" + ID, "S_Wind" ,"scrollbars=1,toolbar=0,titlebar=0,menubar=0;top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
	}	
function SelectSupplier(departureId,DepartureCode)
{
//alert(event.clientX)
	newWin=window.open("editSelectSupplier.aspx?DepartureId=" + escape(departureId) +"&DepartureCode="+escape(DepartureCode), "Supp", "toolbar=0,menubar=0,width=300,height=150,top=" + event.clientY +",left="+event.clientX +",scrollbars=auto");
	newWin.opener=window
}
function GuideTel(departureId,DepartureCode)
{
//alert(event.clientX)
	newWin=window.open("editGuideTel.aspx?DepartureId=" + escape(departureId) +"&DepartureCode="+escape(DepartureCode), "GuideTel", "toolbar=0,menubar=0,width=300,height=150,top=" + event.clientY +",left="+event.clientX +",scrollbars=auto");
	newWin.opener=window
}
function Financial_Dolar(departureId,DepartureCode)
{
//alert(event.clientX)
	newWin=window.open("editFinancial_Dolar.aspx?DepartureId=" + escape(departureId) +"&DepartureCode="+escape(DepartureCode), "FDolar", "toolbar=0,menubar=0,width=300,height=150,top=" + event.clientY +",left="+event.clientX +",scrollbars=auto");
	newWin.opener=window
}
function Financial_Euro(departureId,DepartureCode)
{
//alert(event.clientX)
	newWin=window.open("editFinancial_Euro.aspx?DepartureId=" + escape(departureId) +"&DepartureCode="+escape(DepartureCode), "FEuro", "toolbar=0,menubar=0,width=300,height=150,top=" + event.clientY +",left="+event.clientX +",scrollbars=auto");
	newWin.opener=window
}
function DepartureTimeBrief(departureId,DepartureCode)
{
//alert(event.clientX)
	newWin=window.open("editDepartureTimeBrief.aspx?DepartureId=" + escape(departureId) +"&DepartureCode="+escape(DepartureCode), "DepartureTimeBrief", "toolbar=0,menubar=0,width=300,height=150,top=" + event.clientY +",left="+event.clientX +",scrollbars=auto");
	newWin.opener=window
}

function TimeMeetingAfterTrip(departureId,DepartureCode)
{
//alert(event.clientX)
	newWin=window.open("editTimeMeetingAfterTrip.aspx?DepartureId=" + escape(departureId) +"&DepartureCode="+escape(DepartureCode), "TimeMeetingAfterTrip", "toolbar=0,menubar=0,width=300,height=150,top=" + event.clientY +",left="+event.clientX +",scrollbars=auto");
	newWin.opener=window
}
function Departure_TimeGroupMeeting(departureId,DepartureCode)
{
//alert(event.clientX)
	newWin=window.open("editDeparture_TimeGroupMeeting.aspx?DepartureId=" + escape(departureId) +"&DepartureCode="+escape(DepartureCode), "Departure_TimeGroupMeeting", "toolbar=0,menubar=0,width=300,height=150,top=" + event.clientY +",left="+event.clientX +",scrollbars=auto");
	newWin.opener=window
}
function EditCalendar (departureId,fieldName,DepartureCode)
{
//alert(event.clientX)
	newWin=window.open("editCalendar.aspx?DepartureId=" + escape(departureId)+"&fName="+escape(fieldName)+"&DepartureCode="+escape(DepartureCode), "EditCalendar", "toolbar=0,menubar=0,width=360,height=300,top=" + event.clientY +",left="+event.clientX +",scrollbars=auto");
	newWin.opener=window
}
function EditSelect (departureId,fieldName,DepartureCode)
{
//alert(event.clientX)
	newWin=window.open("editSelect.aspx?DepartureId=" + escape(departureId)+"&fName="+escape(fieldName) +"&DepartureCode="+escape(DepartureCode), "EditSelect", "toolbar=0,menubar=0,width=250,height=160,top=" + event.clientY +",left="+event.clientX +",scrollbars=auto");
	newWin.opener=window
}

//-->

</script>
<style>
.div1
{height:100%;position:relative;width:100%}
.div2
{position: relative;top:30%;right:0px;-webkit-transform: translate(0%, -30%);transform : translate(0%, -30%);color:#000000}
a.link1
{cursor:hand;text-decoration:none;height:100%;width:100%;display:block;}
a.link1:hover
{text-decoration:none;}
</style>
  </head>
  <body style="margin:0px">
      <form id="Form1" method="post" runat="server">
<table border=0 cellpadding=0 cellspacing=0 width=100% align=left>
<tr ><td align=left>
<table border=0 cellpadding=0 cellspacing=0 width=100%>
<tr >
<td height=30></td></tr></table>
</td></tr>

<tr><td valign=top align=left>
<table  cellpadding=1 cellspacing=1 width=100% align=left style="border:solid 1px #d3d3d3">
<tr bgcolor=#d8d8d8 style="height:25px" >

	<td class="td_admin_5" align="center" width="50" valign=bottom><a href="#" OnClick ="javascript:ClearAll();" Class="button_small1">הצג הכל</a>
	</td>
<td class="td_admin_5" align="center"  valign=bottom><asp:LinkButton Runat="server" ID="btnSearch" CssClass="button_small1">&nbsp;חפש&nbsp;</asp:LinkButton></td>
<Td><!--?הסתייםטיפול--></Td>
<Td>&nbsp;</Td>
<Td class="title_sort" style="width:60px"><a href="" onclick="cal1xx.select(document.getElementById('sMeetingAfterTripDate'),'AsMeetingAfterTripDate','dd/MM/yy'); return false;" id="AsMeetingAfterTripDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sMeetingAfterTripDate" class="searchList" style="WIDTH:55px"
NAME="sMeetingAfterTripDate"></Td>
<Td><!--Euro--></Td>
<Td><!--$--></Td>
<Td class="title_sort" style="width:60px"  valign=bottom><select id=sVouchers_Provider name="sVouchers_Provider" dir=rtl><option value="בחר">בחר</option><option value="כן">כן</option><option value="לא">לא</option><option value="מותאם">מותאם</option></select></Td>

<Td class="title_sort" style="width:45px"  valign=bottom><select id=sDeparture_Costing name="sDeparture_Costing" dir=rtl><option value="בחר">בחר</option><option value="כן">כן</option><option value="לא">לא</option></select></Td>
<Td class="title_sort" style="width:45px"  valign=bottom><select id=sVoucher_Group name="sVoucher_Group" dir=rtl><option value="בחר">בחר</option><option value="כן">כן</option><option value="לא">לא</option></select></Td>
<Td class="title_sort" style="width:45px"  valign=bottom><select id=sVoucher_Simultaneous name="sVoucher_Simultaneous" dir=rtl><option value="בחר">בחר</option><option value="כן">כן</option><option value="לא">לא</option></select></Td>
<Td class="title_sort" style="width:45px"  valign=bottom><select id=sGilboaHotel name="sGilboaHotel" dir=rtl><option value="בחר">בחר</option><option value="כן">כן</option><option value="לא">לא</option></select></Td>
<Td>&nbsp;</Td>
<Td class="title_sort" style="width:60px"><a href="" onclick="cal1xx.select(document.getElementById('sGroupMeetingDate'),'AsGroupMeetingDate','dd/MM/yy'); return false;" id="AsGroupMeetingDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sGroupMeetingDate" class="searchList" style="WIDTH:55px"
											NAME="sGroupMeetingDate"></Td><Td>&nbsp;</Td>
<Td class="title_sort" style="width:60px"><a href="" onclick="cal1xx.select(document.getElementById('sBrifDate'),'AsBrifDate','dd/MM/yy'); return false;" id="AsBrifDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sBrifDate" class="searchList" style="WIDTH:55px"
											NAME="sBrifDate"></Td>
<Td class="title_sort" style="width:60px"><a href="" onclick="cal1xx.select(document.getElementById('sItineraryDate'),'AsItineraryDate','dd/MM/yy'); return false;" id="AsItineraryDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sItineraryDate" class="searchList" style="WIDTH:55px"
											NAME="sItineraryDate"></Td>
<td class="title_sort" style="width:60px" valign=bottom nowrap><select runat="server" id="sSuppliers" class="searchList" dir="rtl" name="sSuppliers" style="width:75px"></select></td>
<Td>&nbsp;<!--tel--></Td>
<td class="title_sort" style="width:60px" valign=bottom nowrap><select runat="server" id="sGuides" class="searchList" dir="rtl" name="sGuides" style="width:75px"></select></td>

<Td class="title_sort" style="width:60px"  valign=bottom><select id="sStatus" name="sStatus" dir=rtl><option value="בחר">בחר</option><option value="כן">כרגע בחול</option><option value="חזר">חזר</option><option value="עתידי">טיול עתידי</option></select></Td>

<Td class="title_sort" style="width:60px"><a href="" onclick="cal1xx.select(document.getElementById('sPayToDate'),'AsPayToDate','dd/MM/yy'); return false;"
											id="AsPayToDate"><img src="../../images/calendar.gif" border="0" align="center"></a>
										<input runat="server" type="text" id="sPayToDate" class="searchList" style="WIDTH:65px"
											NAME="sPayToDate"></Td>
<Td class="title_sort" style="width:60px"><a href="" onclick="cal1xx.select(document.getElementById('sPayFromDate'),'AsPayFromDate','dd/MM/yy'); return false;" id="AsPayFromDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sPayFromDate" class="searchList" style="WIDTH:55px"
											NAME="sPayFromDate"></Td>
<td><!--קוד טיול--></td>
<Td>&nbsp;</Td>
<td class="title_sort" width="20px" valign=bottom><select runat="server" id="sSeries" class="searchList" dir="rtl" name="sSeries" ></select></td>
<Td class="title_sort" width="75px" valign=bottom><%If Request.Cookies("bizpegasus")("Chief") = "1" Then%><select runat="server" id="sUsers" class="searchList" dir="rtl" name="sUsers" style="width:75px"></select><%end if%></Td>
<Td>&nbsp;</Td>
</tr>
<tr bgcolor=#d8d8d8 style="height:45px" >
		<td align=center>הוסף <br>שיחה</td>
	<td class="title_sort" align=center>כמות <BR>שיחות</td>
	<td class="title_sort" align=center>?הסתיים<br>טיפול</td>
	<td class="title_sort" align=center>שעת<BR>פגישה <br>לאחר טיול</td>
	<td class="title_sort" align=center>תאריך פגישה <br>לאחר טיול</td>
	<td class="title_sort" align=center>הערכות<BR>כספית<br>-€</td>
	<td class="title_sort" align=center>הערכות<BR>כספית<br>-$</td>
	<td class="title_sort" align=center dir=rtl>שוברי <br>ספק<BR>קרקע?</td>
	<td class="title_sort" align=center>תמחיר</td>
	<td class="title_sort" align=center dir=rtl>שוברי<BR>הוצאות<BR>קבוצה<BR>ומדריך?</td>
	<td class="title_sort" align=center>שובר<br>?לסימולטני</td>
	<td class="title_sort" align=center>?הוקלדו<br>מלונות<br> בגלבוע</td>
	<td class="title_sort" align=center>שעת<br>מפגש<BR>קבוצה</td>
	<td class="title_sort" align=center>תאריך <br>מפגש<BR>קבוצה</td>
	<td class="title_sort" style="padding:3px" align=center>שעת<BR>בריף</td>
	<td class="title_sort" align=center>תאריך<BR>בריף</td>
	<td class="title_sort" align=center>תאריך<BR>קבלת<br> Itinerary</td>
	<td class="title_sort" align=center  style="padding:3px">ספק</td>
	<td class="title_sort" align=center>טלפון<BR>של מדריך</td>
	<td class="title_sort" align=center>שם<br>המדריך</td>
	<td class="title_sort"  align="center">?כרגע בחול</td>	
	<td class="title_sort"  align="center" nowrap>תאריך<BR>חזרה</td>	
	<td class="title_sort" align="center" nowrap>תאריך<BR>יציאה</td>
	<td class="title_sort"  align="center" >קוד<br>טיול</td>
	<td class="title_sort"  align="center" nowrap>תאריך</td>
	<td class="title_sort" align=center>טיול</td>
	<td class="title_sort" align=right>אופרטור&nbsp;</td>
	<td></td>
</tr>
<asp:Repeater ID=rptData Runat=server>
<ItemTemplate>
<tr style="background-color: rgb(201, 201, 201);height:45px">
	<td align=center>	<a href="#" onclick="window.open('AddGuideMessages.asp?dID=<%#Container.DataItem("Departure_Id")%>','winCA','top=20, left=10, width=1000, height=350, scrollbars=1');">
<img src="../../images/copy_icon.gif" border=0 alt="הוסף שיחה"></a></td>
	<td  align=center  style="height:45px">
	<div class="div1">
	<a href="#"   class="link1" onclick="window.open('GuideMessages.asp?dID=<%#Container.DataItem("Departure_Id")%>','winCA','top=20, left=10, width=1200, height=800, scrollbars=1');">
	<div id="GuideMessages_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#func.GetCountGuideMessages(Container.DataItem("Departure_Id"))%></div>
</a>
</div>
	</td><td  align=center><%#func.StatusOperation(Container.DataItem("Status"))%><%'IIF (IsDBNull(Container.DataItem("Status")),"",func.StatusOperation(Container.DataItem("Status")))%></td>
	<td  align=center  style="height:45px">
<div class="div1">
<div id="TimeMeetingAfterTrip_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="TimeMeetingAfterTrip_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("TimeMeetingAfterTrip")),"", Container.DataItem("TimeMeetingAfterTrip"))%></div></div>
</div>
	</td>
		<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','DateMeetingAfterTrip','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
<div id="DateMeetingAfterTrip_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("DateMeetingAfterTrip")),"", DataBinder.Eval(Container.DataItem, "DateMeetingAfterTrip", "{0:dd/MM/yy}"))%></div>
</a>
</div>
	</td>
	

	<td  align=center  style="height:45px">
<div class="div1">
<div id="Financial_Euro_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="FinancialEuro_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Financial_Euro")),"" , Container.DataItem("Financial_Euro"))%></div></div>
</div>
	</td>	
	<td  align=center  style="height:45px">
<div class="div1">
<div id="Financial_Dolar_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="FinancialDolar_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Financial_Dolar")),"", Container.DataItem("Financial_Dolar"))%></div></div>
</div>
	</td>
	<td  align=center  style="height:45px">
<div class="div1">
<div id="Vouchers_Provider_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="Vouchers_Provider_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Vouchers_Provider")),"" , Container.DataItem("Vouchers_Provider"))%></div></div>
</div>

	</td>
			<td  align=center  style="height:45px;padding-left:15px;">
<div class="div1">
<div id="Departure_Costing_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="Departure_Costing_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_Costing")),"" , Container.DataItem("Departure_Costing"))%></div></div>
</div>

</td>
    			<td  align=center  style="height:45px;padding-left:15px;">

<div class="div1">
<div id="Voucher_Group_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="Voucher_Group_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Voucher_Group")),"" , Container.DataItem("Voucher_Group"))%></div></div>
</div>

</td>
			<td  align=center  style="height:45px;padding-left:15px;">
<div class="div1">
<div id="Voucher_Simultaneous_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="Voucher_Simultaneous_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Voucher_Simultaneous")),"" , Container.DataItem("Voucher_Simultaneous"))%></div></div>
</div>
	</td>
	<td  align=center  style="height:45px;padding-left:15px;">
<div class="div1">
<div id="GilboaHotel_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="GilboaHotel_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("GilboaHotel")),"" , Container.DataItem("GilboaHotel"))%></div></div>
</div>
	</td>
		<td  align=center  style="height:45px">
<div class="div1">
<div id="Departure_TimeGroupMeeting_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="Departure_TimeGroupMeeting_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_TimeGroupMeeting")),"&nbsp;","" & Container.DataItem("Departure_TimeGroupMeeting"))%></div></div>
</div>
	</td>
	<td  align=center  style="height:45px">
	<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateGroupMeeting','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
<div id="Departure_DateGroupMeeting_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_DateGroupMeeting")),"", DataBinder.Eval(Container.DataItem, "Departure_DateGroupMeeting", "{0:dd/MM/yy}"))%></div>
</a>
</div>
</td>
	<td  align=center  style="height:45px">
<div class="div1">
<div id="DepartureTimeBrief_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="DepartureTimeBrief_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_TimeBrief")),"", Container.DataItem("Departure_TimeBrief"))%></div></div>
</div>
	</td>	
				<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateBrief','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
<div id="Departure_DateBrief_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_DateBrief")),"" , DataBinder.Eval(Container.DataItem, "Departure_DateBrief", "{0:dd/MM/yy}"))%></div>
</a>
</div>

	</td>

				<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_Itinerary','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
<div id="Departure_Itinerary_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_Itinerary")),"", DataBinder.Eval(Container.DataItem, "Departure_Itinerary", "{0:dd/MM/yy}"))%></div>
</a>
</div>
	</td>
			<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="SelectSupplier('<%#Container.DataItem("Departure_Id")%>','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
<div id="supplier_Id_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("supplier_Name")),"&nbsp;","&nbsp;" & Container.DataItem("supplier_Name"))%></div>
</a>
</div>
	</td>
<td  align=center  style="height:45px;width:100px">
<div class="div1">
<div id="Departure_GuideTelphone_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div style="border:0px solid #ff0000" id="Departure_GuideTelphone_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_GuideTelphone")),"", func.breaks(Container.DataItem("Departure_GuideTelphone")))%></div></div>
</div>

	</td>
	<td  align=center><%#Container.DataItem("GuideName")%></td>
<td  align="center"><%#func.TourStatusForOperation(Container.DataItem("Departure_Id"),Container.DataItem("Departure_Date"),func.dbNullFix(Container.DataItem("Departure_Date_End")))%></td>	
<td  align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date_End", "{0:dd/MM/yy}")%></td>	
	<td  align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date", "{0:dd/MM/yy}")%></td>
	<td  align="center" width=5%><%#Container.DataItem("Departure_Code")%></td>
	<td  align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date", "{0:MMdd}")%></td>
	<td  align=center><%#Container.DataItem("Series_Name")%></td>
	<td  align=right><%#Container.DataItem("User_Name")%></td>
	<td align=center><input type="button" id="edit_button<%#Container.DataItem("Departure_Id")%>" value="עדכן" class="edit" onclick="edit_row('<%#Container.DataItem("Departure_Id")%>')" style="display: block;">
<input type="button" id="save_button<%#Container.DataItem("Departure_Id")%>" value="שמור" class="save" onclick="save_row('<%#Container.DataItem("Departure_Id")%>')" style="display: none;">
<!--input type="button" id="cancel_button<%#Container.DataItem("Departure_Id")%>" value="ביטול" class="delete" onclick="cancel_row('<%#Container.DataItem("Departure_Id")%>')" style="display: none;"--></td>
	</tr>
</ItemTemplate>
<AlternatingItemTemplate>
	<tr style="background-color: rgb(230, 230, 230);height:45px" id="row<%#Container.DataItem("Departure_Id")%>">
	<td align=center>	<a href="#" onclick="window.open('AddGuideMessages.asp?dID=<%#Container.DataItem("Departure_Id")%>','winCA','top=20, left=10, width=1000, height=350, scrollbars=1');">
<img src="../../images/copy_icon.gif" border=0 alt="הוסף שיחה"></a></td>
	<td  align=center  style="height:45px">
	<div class="div1">
	<a href="#"   class="link1" onclick="window.open('GuideMessages.asp?dID=<%#Container.DataItem("Departure_Id")%>','winCA','top=20, left=10, width=1200, height=800, scrollbars=1');">
	<div id="GuideMessages_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#func.GetCountGuideMessages(Container.DataItem("Departure_Id"))%></div>
</a>
</div>
	</td>	<td  align=center><%#func.StatusOperation(Container.DataItem("Status"))%></td>
	<td  align=center  style="height:45px">
<div class="div1">
<div id="TimeMeetingAfterTrip_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="TimeMeetingAfterTrip_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("TimeMeetingAfterTrip")),"" , Container.DataItem("TimeMeetingAfterTrip"))%></div></div>
</div>
	</td>
			<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','DateMeetingAfterTrip','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
<div id="DateMeetingAfterTrip_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("DateMeetingAfterTrip")),"", DataBinder.Eval(Container.DataItem, "DateMeetingAfterTrip", "{0:dd/MM/yy}"))%></div>
</a>
</div>
	</td>
	
		<td  align=center  style="height:45px" >
<div class="div1">
<div id="Financial_Euro_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="FinancialEuro_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Financial_Euro")),"" , Container.DataItem("Financial_Euro"))%></div></div>

</div>
	</td>
	<td  align=center  style="height:45px">
<div class="div1">
<div id="Financial_Dolar_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="FinancialDolar_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Financial_Dolar")),"", Container.DataItem("Financial_Dolar"))%></div></div>
</div>
	</td>


<td  align=center  style="height:45px">
<div class="div1">
<div id="Vouchers_Provider_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="Vouchers_Provider_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Vouchers_Provider")),"" , Container.DataItem("Vouchers_Provider"))%></div></div>
</div>

	</td>		
		<td  align=center  style="height:45px;padding-left:15px;">

<div class="div1">
<div id="Departure_Costing_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="Departure_Costing_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_Costing")),"" , Container.DataItem("Departure_Costing"))%></div></div>
</div>
	</td>
    			<td  align=center  style="height:45px;padding-left:15px;">

<div class="div1">
<div id="Voucher_Group_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="Voucher_Group_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Voucher_Group")),"" , Container.DataItem("Voucher_Group"))%></div></div>
</div>

	</td>
			<td  align=center  style="height:45px;padding-left:15px;">
<div class="div1">
<div id="Voucher_Simultaneous_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="Voucher_Simultaneous_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Voucher_Simultaneous")),"" , Container.DataItem("Voucher_Simultaneous"))%></div></div>
</div>
	</td>
			<td  align=center  style="height:45px;padding-left:15px;">
<div class="div1">
<div id="GilboaHotel_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="GilboaHotel_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("GilboaHotel")),"", Container.DataItem("GilboaHotel"))%></div></div>
</div>
	</td>
	<td  align=center  style="height:45px">
<div class="div1">
<div id="Departure_TimeGroupMeeting_<%#Container.DataItem("Departure_Id")%>"  class="div2">
<div id="Departure_TimeGroupMeeting_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_TimeGroupMeeting")),"" , Container.DataItem("Departure_TimeGroupMeeting"))%></div></div>
</div>
	</td>
	<td  align=center  style="height:45px">
	<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateGroupMeeting','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
<div id="Departure_DateGroupMeeting_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_DateGroupMeeting")),"" , DataBinder.Eval(Container.DataItem, "Departure_DateGroupMeeting", "{0:dd/MM/yy}"))%></div>
</a>
</div>
</td>
	<td  align=center  style="height:45px">
<div class="div1">
<div id="DepartureTimeBrief_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="DepartureTimeBrief_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_TimeBrief")),"" , Container.DataItem("Departure_TimeBrief"))%></div></div>
</div>
	</td>	
				<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateBrief','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
<div id="Departure_DateBrief_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_DateBrief")),"" , DataBinder.Eval(Container.DataItem, "Departure_DateBrief", "{0:dd/MM/yy}"))%></div>
</a>
</div>
	</td>

				<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_Itinerary','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
<div id="Departure_Itinerary_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_Itinerary")),"" , DataBinder.Eval(Container.DataItem, "Departure_Itinerary", "{0:dd/MM/yy}"))%></div>
</a>
</div>
	</td>

			<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="SelectSupplier('<%#Container.DataItem("Departure_Id")%>','supplier_Id','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
<div id="supplier_Id_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("supplier_Name")),"" , Container.DataItem("supplier_Name"))%></div>
</a>
</div>
	</td>

<td  align=center  style="height:45px">
<div class="div1">
<div id="Departure_GuideTelphone_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="Departure_GuideTelphone_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_GuideTelphone")),"", func.breaks(Container.DataItem("Departure_GuideTelphone")))%></div></div>
</div>

	</td>


	<td  align=center><%'#iif(container.DataItem("GuideName") is DBNULL.value,"",container.DataItem("GuideName") )%><%#Container.DataItem("GuideName")%></td>
	<td  align="center"><%#func.TourStatusForOperation(Container.DataItem("Departure_Id"),Container.DataItem("Departure_Date"),func.dbNullFix(Container.DataItem("Departure_Date_End")))%></td>	
	<td  align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date_End", "{0:dd/MM/yy}")%></td>	
	<td  align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date", "{0:dd/MM/yy}")%></td>
	<td  align="center"><%#Container.DataItem("Departure_Code")%></td>
	<td  align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date", "{0:MMdd}")%></td>
	<td  align=center><%#Container.DataItem("Series_Name")%></td>
	<td  align=right><%#Container.DataItem("User_Name")%></td>
		<td align=center><input type="button" id="edit_button<%#Container.DataItem("Departure_Id")%>" value="עדכן" class="edit" onclick="edit_row('<%#Container.DataItem("Departure_Id")%>')" style="display: block;">
<input type="button" id="save_button<%#Container.DataItem("Departure_Id")%>" value="שמור" class="save" onclick="save_row('<%#Container.DataItem("Departure_Id")%>')" style="display: none;">
</td>
</tr>
</AlternatingItemTemplate>

</asp:Repeater>

				
</table>

</td>
	</tr>
	<!--  <div id="dvCustomers">-->
 		<!--paging-->
						<asp:PlaceHolder id="pnlPages" Runat="server">
						<tr><td height="2"></td></tr>
						<tr>
							<td class="plata_paging" vAlign="top" align="center" colspan="7" bgcolor=#D8D8D8>
								<table dir="ltr" cellSpacing="0" cellPadding="2" width="100%" border="0" >
									<tr>
										<td class="plata_paging" vAlign="baseline" noWrap align="left" width="150">&nbsp;הצג
											<asp:DropDownList id="PageSize" CssClass="PageSelect" Runat="server" AutoPostBack="true">
													<asp:ListItem Value="10">10</asp:ListItem>
												<asp:ListItem Value="20">20</asp:ListItem>
													<asp:ListItem Value="50" Selected="True">50</asp:ListItem>
											</asp:DropDownList>&nbsp;בדף&nbsp;
										</td>
										<td vAlign="baseline" noWrap align="right" >
											<asp:LinkButton id="cmdNext" runat="server" CssClass="page_link" text="«עמוד הבא"></asp:LinkButton></TD>
										<td class="plata_paging" vAlign="baseline" noWrap align="center" width="160">
											<asp:Label id="lblTotalPages" Runat="server"></asp:Label>&nbsp;דף&nbsp;
											<asp:DropDownList id="pageList" CssClass="PageSelect" Runat="server" AutoPostBack="true"></asp:DropDownList>&nbsp;מתוך&nbsp;
										</td>
										<td vAlign="baseline" noWrap align="left" >
											<asp:LinkButton id="cmdPrev" runat="server" CssClass="page_link" text="עמוד קודם»"></asp:LinkButton></TD>
										<td class="plata_paging" vAlign="baseline" align="right">&nbsp;נמצאו&nbsp;&nbsp;&nbsp;
											<asp:Label CssClass="plata_paging" id="lblCount" Runat="server"></asp:Label>&nbsp;יציאות
										</td>
										</tr>				
									</table>
								</td>
							</tr>
						</asp:PlaceHolder>
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

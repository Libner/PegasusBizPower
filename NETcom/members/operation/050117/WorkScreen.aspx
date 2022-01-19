<%@ Page Language="vb" AutoEventWireup="false" Codebehind="WorkScreen.aspx.vb" Inherits="bizpower_pegasus.WorkScreen"%>
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
		<script src="../../include/CalendarPopup.js" charset="windows-1255"></script>
	<style>
	.wrapper1, .wrapper2 { width: 100%; overflow-x: scroll; overflow-y: hidden; }
.wrapper1 { height: 20px; }
.wrapper2 {}
.div11 { height: 20px; }
.div21 { overflow: none; }
	</style>
	<script>
	$(function () {
    $('.wrapper1').on('scroll', function (e) {
        $('.wrapper2').scrollLeft($('.wrapper1').scrollLeft());
    }); 
    $('.wrapper2').on('scroll', function (e) {
        $('.wrapper1').scrollLeft($('.wrapper2').scrollLeft());
    });
});
$(window).on('load', function (e) {
    $('.div11').width($('table').width());
    $('.div21').width($('table').width());
});
	</script>
	
	
		<script>
		
function isValidTime(time) {

    return time.match(/(^([0-9]|[0-1][0-9]|[2][0-3]):([0-5][0-9])$)|(^([0-9]|[1][0-9]|[2][0-3])$)/);
  }
   
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

}
 function replaceAll(find, replace, str) 
    {
      while( str.indexOf(find) > -1)
      {
        str = str.replace(find, replace);
      }
      return str;
    }
 function checkNumbers(evt, objInput) {
    /* var theEvent = evt || window.event;
    // var key = theEvent.keyCode || theEvent.which;
    var key = (theEvent.which) ? theEvent.which :
    ((theEvent.charCode) ? theEvent.charCode :
    ((theEvent.keyCode) ? theEvent.keyCode : 0));
    // Don't validate the input if below arrow, delete and backspace keys were pressed 
    if (key == 37 || key == 38 || key == 39 || key == 40 || key == 8 || key == 46) { // Left / Up / Right / Down Arrow, Backspace, Delete keys
    return;
    }*/
    if (objInput.value.length > 0) {
        keys = objInput.value.charAt(objInput.value.length - 1);
        var regex = /[0-9]/;
        if (!regex.test(keys)) {
            // alert(' checkNumbers-' + objInput.value)

            objInput.value = objInput.value.substring(0, objInput.value.length - 1)
            // alert(' checkNumbers-'+objInput.value)
            //return false;
        }
    }
    else
        return;
}
function validate(evt) {
  var theEvent = evt || window.event;
  var key = theEvent.keyCode || theEvent.which;
  key = String.fromCharCode( key );
  var regex = /[0-9]|\:/;
  if( !regex.test(key) ) {
    theEvent.returnValue = false;
    if(theEvent.preventDefault) theEvent.preventDefault();
  }
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
 document.getElementById("TimeMeetingAfterTrip_row"+no).innerHTML="<input style='width:35px' type='text' onkeypress='validate(event)' id='TimeMeetingAfterTrip_text"+no+"' value='"+document.getElementById("TimeMeetingAfterTrip_row"+no).innerHTML+"'>";
 document.getElementById("FinancialEuro_row"+no).innerHTML="<input style='width:65px' type='text' dir='ltr' id='FinancialEuro_text"+no+"' value='"+document.getElementById("FinancialEuro_row"+no).innerHTML+"'>";
 document.getElementById("Departure_GuideTelphone_row"+no).innerHTML="<textarea type='text' dir='ltr' id='Departure_GuideTelphone_text"+no+"' style='width:100px;height:20px'>" + replaceAll("<br>","\n",document.getElementById("Departure_GuideTelphone_row"+no).innerHTML)+"</textarea>";
// document.getElementById("DateMeetingAfterTrip_row"+no).innerHTML="<a href='#'  onclick=EditCalendar('"+no +"','DateMeetingAfterTrip','"+no+"');return false class='link1'><img src=../../images/more.png></a>";
 document.getElementById("DateMeetingAfterTrip_row"+no).style.display="block";
 document.getElementById("Departure_DateBrief_row"+no).style.display="block";
 document.getElementById("Departure_DateGroupMeeting_row"+no).style.display="block";
 document.getElementById("Departure_Itinerary_row"+no).style.display="block";
 document.getElementById("supplier_Id_row"+no).style.display="block";

 
 document.getElementById("FinancialDolar_row"+no).innerHTML="<input style='width:65px' type='text' dir='ltr' id='FinancialDolar_text"+no+"' value='"+document.getElementById("FinancialDolar_row"+no).innerHTML+"'>";
 document.getElementById("DepartureTimeBrief_row"+no).innerHTML="<input style='width:35px' onkeypress='validate(event)' dir='ltr' type='text' id='DepartureTimeBrief_text"+no+"' value='"+document.getElementById("DepartureTimeBrief_row"+no).innerHTML+"'>";
 document.getElementById("Departure_TimeGroupMeeting_row"+no).innerHTML="<input style='width:35px' dir='ltr' onkeypress='validate(event)' type='text' id='Departure_TimeGroupMeeting_text"+no+"' value='"+document.getElementById("Departure_TimeGroupMeeting_row"+no).innerHTML+"'>";
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
	if (document.getElementById("TimeMeetingAfterTrip_text"+no).value!='')
		{
			if (!isValidTime(document.getElementById("TimeMeetingAfterTrip_text"+no).value))
			{
				alert ("פורמט שעה אינו נכון\nפורמט הנכון HH:MM");
				document.getElementById("TimeMeetingAfterTrip_text"+no).focus()
				return false;
			}
		}
		if (document.getElementById("DepartureTimeBrief_text"+no).value!='')
		{
			if (!isValidTime(document.getElementById("DepartureTimeBrief_text"+no).value))
			{
				alert ("פורמט שעה אינו נכון\nפורמט הנכון HH:MM");
				document.getElementById("DepartureTimeBrief_text"+no).focus()
				return false;
			}
		}
			if (document.getElementById("Departure_TimeGroupMeeting_text"+no).value!='')
		{
			if (!isValidTime(document.getElementById("Departure_TimeGroupMeeting_text"+no).value))
			{
				alert ("פורמט שעה אינו נכון\nפורמט הנכון HH:MM");
				document.getElementById("Departure_TimeGroupMeeting_text"+no).focus()
				return false;
			}
		}
		
		
if (document.getElementById("Departure_TimeGroupMeeting_text"+no).value!='')
		{
			if (!isValidTime(document.getElementById("Departure_TimeGroupMeeting_text"+no).value))
			{
				alert ("פורמט שעה אינו נכון\nפורמט הנכון HH:MM");
				document.getElementById("Departure_TimeGroupMeeting_text"+no).focus()
				return false;
			}
		}
		if (document.getElementById("DepartureTimeBrief_text"+no).value!='')
		{
			if (!isValidTime(document.getElementById("DepartureTimeBrief_text"+no).value))
			{
				alert ("פורמט שעה אינו נכון\nפורמט הנכון HH:MM");
				document.getElementById("DepartureTimeBrief_text"+no).focus()
				return false;
			}
		}
		
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
  document.getElementById("Departure_GuideTelphone_row"+no).innerHTML=document.getElementById("Departure_GuideTelphone_text"+no).value;

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
//document.getElementById("DateMeetingAfterTrip_row"+no).innerHTML=document.getElementById("DateMeetingAfterTrip_"+no).value;
 document.getElementById("DateMeetingAfterTrip_row"+no).style.display="none";
 document.getElementById("Departure_DateBrief_row"+no).style.display="none";
 document.getElementById("Departure_DateGroupMeeting_row"+no).style.display="none";
 document.getElementById("Departure_Itinerary_row"+no).style.display="none";
 document.getElementById("supplier_Id_row"+no).style.display="none";

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
function FuncSort(query,par,srt)
{

var sUser,sSeries,  sGuides, sSuppliers
sUser=$("#sUsers").val()
sSeries=$("#sSeries").val()
sGuides =$("#sGuides").val()
sSuppliers=$("#sSuppliers").val()
sFromD=$("#sPayFromDate").val()
sToD=$("#sPayToDate").val()
sPage=$("#pageList").val()
//sStatus=$("#sStatus").val()

if (query!="")
{
var r  ;
r="&"+par +"=ASC"
query = query.replace(r, "");
r="&"+par +"=DESC"
query = query.replace(r, "");

window.location ="WorkScreen.aspx?" +query +"&"+par +"="+ srt
}
else
{ 
window.location ="WorkScreen.aspx?"+"sUser="+ escape(sUser) +"&sFromD="+ escape(sFromD) +"&sToD="+ escape(sToD) +"&sPage=" + escape(sPage)+"&"+par +"="+ srt
}
}

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
{/*height:100%;*/position:relative;width:100%}
.div2
{position: relative;/*top:30%;right:0px;-webkit-transform: translate(0%, -30%);transform : translate(0%, -30%);*/color:#000000}
a.link1
{cursor:hand;text-decoration:none;height:100%;width:100%;display:block;}
a.link1:hover
{text-decoration:none;}
		</style>
	</head>
	<body style="margin:0px">
		<form id="Form1" method="post" runat="server">
	<div class="wrapper1">
	<div class="div11"></div>
</div>
<div class="wrapper2">
	<div class="div21">
			<table border="0" cellpadding="0" cellspacing="0" width="100%" align="left">
				<tr>
					<td align="left">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td height="30" align="center"><span style="COLOR: #6F6DA6;font-size:14pt"><%=tabName%></span></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td valign="top" align="left">
						<table cellpadding="1" cellspacing="1" width="100%" align="left" style="border:solid 1px #d3d3d3">
							<tr bgcolor="#d8d8d8" style="height:25px">
								<td class="td_admin_5" align="center" width="50" valign="bottom"><a href="#" OnClick="javascript:ClearAll();" Class="button_small1">הצג 
										הכל</a>
								</td>
								<td class="td_admin_5" align="center" valign="bottom"><asp:LinkButton Runat="server" ID="btnSearch" CssClass="button_small1">&nbsp;חפש&nbsp;</asp:LinkButton></td>
								<Td><!--?הסתייםטיפול--></Td>
								<Td>&nbsp;</Td>
								<Td class="title_sort" style="width:89px" nowrap valign="bottom"><a href="" onclick="cal1xx.select(document.getElementById('sFromMeetingAfterTripDate'),'AsFromMeetingAfterTripDate','dd/MM/yy'); return false;"
										id="AsFromMeetingAfterTripDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sFromMeetingAfterTripDate" class="searchList" style="WIDTH:50px"
										NAME="sFromMeetingAfterTripDate" readonly><br>
									<a href="" onclick="cal1xx.select(document.getElementById('sToMeetingAfterTripDate'),'AsToMeetingAfterTripDate','dd/MM/yy'); return false;"
										id="AsToMeetingAfterTripDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sToMeetingAfterTripDate" class="searchList" style="WIDTH:50px"
										NAME="sToMeetingAfterTripDate" readonly></Td>
								<Td><!--Euro--></Td>
								<Td><!--$--></Td>
								<Td class="title_sort" style="width:60px" valign="bottom"><select id="sVouchers_Provider" name="sVouchers_Provider" dir="rtl"><option value="בחר">בחר</option>
										<option value="כן" <%if pVouchers_Provider="כן" then%>selected<%end if%>>כן</option>
										<option value="לא" <%if pVouchers_Provider="לא" then%>selected<%end if%>>לא</option>
										<option value="מותאם" <%if pVouchers_Provider="מותאם" then%>selected<%end if%>>מותאם</option>
									</select></Td>
								<Td class="title_sort" style="width:45px" valign="bottom"><select id="sDeparture_Costing" name="sDeparture_Costing" dir="rtl"><option value="בחר">בחר</option>
										<option value="כן" <%if pDeparture_Costing="כן" then%>selected<%end if%>>כן</option>
										<option value="לא" <%if pDeparture_Costing="לא" then%>selected<%end if%>>לא</option>
									</select></Td>
								<Td class="title_sort" style="width:45px" valign="bottom"><select id="sVoucher_Group" name="sVoucher_Group" dir="rtl"><option value="בחר">בחר</option>
										<option value="כן" <%if pVoucher_Group="כן" then%>selected<%end if%>>כן</option>
										<option value="לא" <%if pVoucher_Group="לא" then%>selected<%end if%>>לא</option>
									</select></Td>
								<Td class="title_sort" style="width:45px" valign="bottom"><select id="sVoucher_Simultaneous" name="sVoucher_Simultaneous" dir="rtl"><option value="בחר">בחר</option>
										<option value="כן" <%if pVoucher_Simultaneous="כן" then%>selected<%end if%>>כן</option>
										<option value="לא" <%if pVoucher_Simultaneous="לא" then%>selected<%end if%>>לא</option>
									</select></Td>
								<Td class="title_sort" style="width:45px" valign="bottom"><select id="sGilboaHotel" name="sGilboaHotel" dir="rtl"><option value="בחר">בחר</option>
										<option value="כן" <%if pGilboaHotel="כן" then%>selected<%end if%>>כן</option>
										<option value="לא" <%if pGilboaHotel="לא" then%>selected<%end if%>>לא</option>
									</select></Td>
								<Td>&nbsp;</Td>
								<Td class="title_sort" style="width:89px" nowrap valign="bottom"><a href="" onclick="cal1xx.select(document.getElementById('sFromGroupMeetingDate'),'AsFromGroupMeetingDate','dd/MM/yy'); return false;"
										id="AsFromGroupMeetingDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sFromGroupMeetingDate" class="searchList" style="WIDTH:50px"
										NAME="sFromGroupMeetingDate" readonly><br>
									<a href="" onclick="cal1xx.select(document.getElementById('sToGroupMeetingDate'),'AsToGroupMeetingDate','dd/MM/yy'); return false;"
										id="AsToGroupMeetingDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sToGroupMeetingDate" class="searchList" style="WIDTH:50px"
										NAME="sToGroupMeetingDate" readonly></Td>
								<Td>&nbsp;</Td>
								<Td class="title_sort" style="width:89px" nowrap valign="bottom"><a href="" onclick="cal1xx.select(document.getElementById('sFromBrifDate'),'AsFromBrifDate','dd/MM/yy'); return false;"
										id="AsFromBrifDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sFromBrifDate" class="searchList" style="WIDTH:50px"
										NAME="sFromBrifDate" readonly><br>
									<a href="" onclick="cal1xx.select(document.getElementById('sToBrifDate'),'AsToBrifDate','dd/MM/yy'); return false;"
										id="AsToBrifDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sToBrifDate" class="searchList" style="WIDTH:50px"
										NAME="sToBrifDate" readonly></Td>
								<Td class="title_sort" style="width:89px" nowrap valign="bottom"><a href="" onclick="cal1xx.select(document.getElementById('sFromItineraryDate'),'AsFromItineraryDate','dd/MM/yy'); return false;"
										id="AsFromItineraryDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sFromItineraryDate" class="searchList" style="WIDTH:50px"
										NAME="sFromItineraryDate" readonly><br>
									<a href="" onclick="cal1xx.select(document.getElementById('sToItineraryDate'),'AsToItineraryDate','dd/MM/yy'); return false;"
										id="AsToItineraryDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sToItineraryDate" class="searchList" style="WIDTH:50px"
										NAME="sToItineraryDate" readonly></Td>
								<td class="title_sort" style="width:30px" valign="bottom" nowrap><select runat="server" id="sSuppliers" class="searchList" dir="rtl" name="sSuppliers" style="width:75px"></select></td>
								<Td>&nbsp;<!--tel--></Td>
								<td class="title_sort" style="width:60px" valign="bottom" nowrap><select runat="server" id="sGuides" class="searchList" dir="rtl" name="sGuides" style="width:75px"></select></td>
								<Td class="title_sort" style="width:30px" valign="bottom"><select id="sStatus" name="sStatus" dir="rtl"><option value="בחר">בחר</option>
										<option value="כן" <%if pStatus="כן" or tab="1" or tab="3" then%>selected<%end if%>>בחול</option>
										<option value="חזר" <%if pStatus="חזר" then%>selected<%end if%>>חזר</option>
										<option value="עתידי" <%if pStatus="עתידי" then%>selected<%end if%>>עתידי</option>
									</select></Td>
								<Td class="title_sort" valign="bottom" style="width:89px" nowrap><a href="" onclick="cal1xx.select(document.getElementById('sFromDateEnd'),'AsFromDateEnd','dd/MM/yy'); return false;"
										id="AsFromDateEnd"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sFromDateEnd" class="searchList" style="WIDTH:50px"
										NAME="sFromDateEnd" readonly>
									<br>
									<a href="" onclick="cal1xx.select(document.getElementById('sToDateEnd'),'AsToDateEnd','dd/MM/yy'); return false;"
										id="AsToDateEnd"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sToDateEnd" class="searchList" style="WIDTH:50px"
										NAME="sToDateEnd" readonly></Td>
								<Td class="title_sort" style="width:89px" valign="bottom" nowrap><a href="" onclick="cal1xx.select(document.getElementById('sFromDate'),'AsFromDate','dd/MM/yy'); return false;"
										id="AsFromDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sFromDate" class="searchList" style="WIDTH:50px"
										NAME="sFromDate" readonly><br>
									<a href="" onclick="cal1xx.select(document.getElementById('sToDate'),'AsToDate','dd/MM/yy'); return false;"
										id="AsToDate"><img src="../../images/calendar.gif" border="0" align="center"></a><input runat="server" type="text" id="sToDate" class="searchList" style="WIDTH:50px" readonly
										NAME="sToDate"></Td>
								<td class="title_sort" valign="bottom"><input type="text" id="codeTiul" name="codeTiul" style="width:82px"></td>
								<Td>&nbsp;</Td>
								<td class="title_sort" width="20px" valign="bottom"><select runat="server" id="sSeries" class="searchList" dir="rtl" name="sSeries" style="width:40px"></select></td>
								<Td class="title_sort" width="45px" valign="bottom"><%If Request.Cookies("bizpegasus")("Chief") = "1" Then%><select runat="server" id="sUsers" class="searchList" dir="rtl" name="sUsers" style="width:65px"></select><%end if%></Td>
								<Td>&nbsp;</Td>
							</tr>
							<tr bgcolor="#d8d8d8" style="height:60px">
								<td align="center">הוסף
									<br>
									שיחה</td>
								<td class="title_sort" align="center">כמות
									<BR>
									שיחות</td>
								<td class="title_sort" align="center" dir="rtl" nowrap>טיפול<br>
									הסתיים?</td>
								<td class="title_sort" align="center">שעת<BR>
									פגישה
									<br>
									לאחר<BR>
									טיול</td>
								<td class="title_sort" align="center">תאריך פגישה
									<br>
									לאחר טיול<br>
									<a href="javascript:FuncSort('<%=qrystring%>','sort_9','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_9")="ASC" then%>_act<%end if%>.gif"  title="למיין לפי תאריך מפגש לאחר טיול" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_9','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_9")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי תאריך מפגש לאחר טיול" border=0></a></td>
								<td class="title_sort" align="center">הערכות<BR>
									כספית<br>
									-€</td>
								<td class="title_sort" align="center">הערכות<BR>
									כספית<br>
									-$</td>
								<td class="title_sort" align="center" dir="rtl">שוברי
									<br>
									ספק<BR>
									קרקע?</td>
								<td class="title_sort" align="center">תמחיר</td>
								<td class="title_sort" align="center" dir="rtl">שוברי<BR>
									הוצאות<BR>
									קבוצה<BR>
									ומדריך?</td>
								<td class="title_sort" align="center" dir="rtl">שובר<br>
									לסימולטני?</td>
								<td class="title_sort" align="center" dir="rtl">הוקלדו<br>
									מלונות<br>
									בגלבוע?</td>
								<td class="title_sort" align="center">שעת<br>
									מפגש<BR>
									קבוצה</td>
								<td class="title_sort" align="center">תאריך<br>
									מפגש<BR>
									קבוצה<br>
									<a href="javascript:FuncSort('<%=qrystring%>','sort_8','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_8")="ASC" then%>_act<%end if%>.gif"  title="למיין לפי תאריך מפגש" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_8','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_8")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי תאריך מפגש" border=0></a></td>
								<td class="title_sort" style="padding:3px" align="center">שעת<BR>
									בריף</td>
								<td class="title_sort" align="center">תאריך<BR>
									בריף<br>
									<a href="javascript:FuncSort('<%=qrystring%>','sort_7','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_7")="ASC" then%>_act<%end if%>.gif"  title="למיין לפי תאריך בריף" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_7','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_7")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי תאריך בריף" border=0></a></td>
								<td class="title_sort" align="center">תאריך<BR>
									קבלת<br>
									Itinerary<br>
									<a href="javascript:FuncSort('<%=qrystring%>','sort_6','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_6")="ASC" then%>_act<%end if%>.gif"  title="למיין לפי Itinerary" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_6','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_6")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי Itinerary" border=0></a></td>
								<td class="title_sort" align="center" style="padding:3px">ספק<br>
									<a href="javascript:FuncSort('<%=qrystring%>','sort_10','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_10")="ASC" then%>_act<%end if%>.gif"  title="למיין לפי ספק" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_10','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_10")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי ספק" border=0></a></td>
								<td class="title_sort" align="center">טלפון<BR>
									של מדריך</td>
								<td class="title_sort" align="center">שם<br>
									המדריך<br>
									<a href="javascript:FuncSort('<%=qrystring%>','sort_5','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_5")="ASC" then%>_act<%end if%>.gif"  title="למיין לפי המדריך" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_5','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_5")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי המדריך" border=0></a></td>
								<td class="title_sort" align="center">?כרגע בחול</td>
								<td class="title_sort" align="center">תאריך<BR>
									חזרה<br>
									<a href="javascript:FuncSort('<%=qrystring%>','sort_4','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_4")="ASC" then%>_act<%end if%>.gif"  title="למיין לפי תאריך חזרה" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_4','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_4")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי תאריך חזרה" border=0></a></td>
								<td class="title_sort" align="center">&nbsp;תאריך<BR>
									יציאה<br>
									<a href="javascript:FuncSort('<%=qrystring%>','sort_3','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_3")="ASC" then%>_act<%end if%>.gif"  title="למיין לפי תאריך יציאה" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_3','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_3")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי תאריך יציאה" border=0></a></td>
								<td class="title_sort" align="center">קוד<br>
									טיול</td>
								<td class="title_sort" align="center" nowrap>תאריך</td>
								<td class="title_sort" align="center">&nbsp;טיול<br>
									<a href="javascript:FuncSort('<%=qrystring%>','sort_2','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_2")="ASC" then%>_act<%end if%>.gif"  title="למיין לפי קוד סדרה" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_2','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_2")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי קוד סדרה" border=0></a></td>
								<td class="title_sort" align="center">&nbsp;אופרטור&nbsp;<br>
									<a href="javascript:FuncSort('<%=qrystring%>','sort_1','ASC')"><img src="../../../images/arrow_top<%If  Request.Querystring("sort_1")="ASC" then%>_act<%end if%>.gif"  title="למיין לפי אופרטור" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_1','DESC')"><img src="../../../images/arrow_bot<%If  Request.Querystring("sort_1")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי אופרטור" border=0></a></td>
								<td></td>
							</tr>
							<asp:Repeater ID="rptData" Runat="server">
								<ItemTemplate>
									<tr style="background-color: rgb(201, 201, 201);height:60px">
										<td align="center">
											<a href="#" onclick="window.open('AddGuideMessages.asp?dID=<%#Container.DataItem("Departure_Id")%>','winCA','top=20, left=10, width=1000, height=350, scrollbars=1');">
												<img src="../../images/copy_icon.gif" border="0" alt="הוסף שיחה"></a></td>
										<td align="center" style="height:60px">
											<div class="div1">
												<a href="#"   class="link1" onclick="window.open('GuideMessages.asp?dID=<%#Container.DataItem("Departure_Id")%>','winCA','top=20, left=10, width=1200, height=800, scrollbars=1');">
													<div id="GuideMessages_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#func.GetCountGuideMessages(Container.DataItem("Departure_Id"))%></div>
												</a>
											</div>
										</td>
										<td align="center"><%#func.StatusOperation(Container.DataItem("Status"))%><%'IIF (IsDBNull(Container.DataItem("Status")),"",func.StatusOperation(Container.DataItem("Status")))%></td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="TimeMeetingAfterTrip_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="TimeMeetingAfterTrip_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("TimeMeetingAfterTrip")),"", Container.DataItem("TimeMeetingAfterTrip"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="DateMeetingAfterTrip_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','DateMeetingAfterTrip','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<%#IIF(IsDBNull(Container.DataItem("DateMeetingAfterTrip")),"", DataBinder.Eval(Container.DataItem, "DateMeetingAfterTrip", "{0:dd/MM/yy}"))%>
													</a>
												</div>
												<div id="DateMeetingAfterTrip_row<%#Container.DataItem("Departure_Id")%>" style="display:none;">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','DateMeetingAfterTrip','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<img src="../../images/more.png" border="0" width="20"></a>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Financial_Euro_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="FinancialEuro_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Financial_Euro")),"&nbsp;" , Container.DataItem("Financial_Euro"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Financial_Dolar_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="FinancialDolar_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Financial_Dolar")),"&nbsp;", Container.DataItem("Financial_Dolar"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Vouchers_Provider_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="Vouchers_Provider_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Vouchers_Provider")),"&nbsp;" , Container.DataItem("Vouchers_Provider"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px;padding-left:15px;">
											<div class="div1">
												<div id="Departure_Costing_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="Departure_Costing_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_Costing")),"&nbsp;" , Container.DataItem("Departure_Costing"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px;padding-left:15px;">
											<div class="div1">
												<div id="Voucher_Group_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="Voucher_Group_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Voucher_Group")),"&nbsp;" , Container.DataItem("Voucher_Group"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px;padding-left:15px;">
											<div class="div1">
												<div id="Voucher_Simultaneous_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="Voucher_Simultaneous_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Voucher_Simultaneous")),"&nbsp;" , Container.DataItem("Voucher_Simultaneous"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px;padding-left:15px;">
											<div class="div1">
												<div id="GilboaHotel_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="GilboaHotel_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("GilboaHotel")),"" , Container.DataItem("GilboaHotel"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Departure_TimeGroupMeeting_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="Departure_TimeGroupMeeting_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_TimeGroupMeeting")),"&nbsp;","" & Container.DataItem("Departure_TimeGroupMeeting"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Departure_DateGroupMeeting_<%#Container.DataItem("Departure_Id")%>"  class="div2" style="border:0px solid #ff0000">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateGroupMeeting','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<%#IIF(IsDBNull(Container.DataItem("Departure_DateGroupMeeting")),"&nbsp;", DataBinder.Eval(Container.DataItem, "Departure_DateGroupMeeting", "{0:dd/MM/yy}"))%>
													</a>
												</div>
												<div id="Departure_DateGroupMeeting_row<%#Container.DataItem("Departure_Id")%>" style="display:none;">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateGroupMeeting','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<img src="../../images/more.png" border="0" width="20"></a>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="DepartureTimeBrief_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="DepartureTimeBrief_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_TimeBrief")),"&nbsp;", Container.DataItem("Departure_TimeBrief"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Departure_DateBrief_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateBrief','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<%#IIF(IsDBNull(Container.DataItem("Departure_DateBrief")),"&nbsp;", DataBinder.Eval(Container.DataItem, "Departure_DateBrief", "{0:dd/MM/yy}"))%>
													</a>
												</div>
												<div id="Departure_DateBrief_row<%#Container.DataItem("Departure_Id")%>" style="display:none;">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateBrief','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<img src="../../images/more.png" border="0" width="20"></a>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Departure_Itinerary_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_Itinerary','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<%#IIF(IsDBNull(Container.DataItem("Departure_Itinerary")),"&nbsp;", DataBinder.Eval(Container.DataItem, "Departure_Itinerary", "{0:dd/MM/yy}"))%>
													</a>
												</div>
												<div id="Departure_Itinerary_row<%#Container.DataItem("Departure_Id")%>" style="display:none;">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_Itinerary','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<img src="../../images/more.png" border="0" width="20"></a>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="supplier_Id_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<a href="#"  onclick="SelectSupplier('<%#Container.DataItem("Departure_Id")%>','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<%#IIF(IsDBNull(Container.DataItem("supplier_Name")),"&nbsp;", Container.DataItem("supplier_Name"))%>
													</a>
												</div>
												<div id="supplier_Id_row<%#Container.DataItem("Departure_Id")%>" style="display:none;">
													<a href="#"  onclick="SelectSupplier('<%#Container.DataItem("Departure_Id")%>','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<img src="../../images/select.png" border="0" width="20"></a>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px;width:100px">
											<div class="div1">
												<div id="Departure_GuideTelphone_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div style="border:0px solid #ff0000" id="Departure_GuideTelphone_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_GuideTelphone")),"&nbsp;", func.breaks(Container.DataItem("Departure_GuideTelphone")))%></div>
												</div>
											</div>
										</td>
										<td align="center"><%#Container.DataItem("GuideName")%></td>
										<td align="center"><%#func.TourStatusForOperation(Container.DataItem("Departure_Id"),Container.DataItem("Departure_Date"),func.dbNullFix(Container.DataItem("Departure_Date_End")))%></td>
										<td align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date_End", "{0:dd/MM/yy}")%></td>
										<td align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date", "{0:dd/MM/yy}")%></td>
										<td align="center" width="5%"><%#Container.DataItem("Departure_Code")%></td>
										<td align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date", "{0:MMdd}")%></td>
										<td align="center"><%#Container.DataItem("Series_Name")%></td>
										<td align="right"><%#Container.DataItem("User_Name")%></td>
										<td align="center" style="width:25px"><input type="button" id="edit_button<%#Container.DataItem("Departure_Id")%>" value="עדכן" class="edit" onclick="edit_row('<%#Container.DataItem("Departure_Id")%>')" style="display: block;">
											<input type="button" id="save_button<%#Container.DataItem("Departure_Id")%>" value="שמור" class="save" onclick="save_row('<%#Container.DataItem("Departure_Id")%>')" style="display: none;">
											<!--input type="button" id="cancel_button<%#Container.DataItem("Departure_Id")%>" value="ביטול" class="delete" onclick="cancel_row('<%#Container.DataItem("Departure_Id")%>')" style="display: none;"--></td>
									</tr>
								</ItemTemplate>
								<AlternatingItemTemplate>
									<tr style="background-color: rgb(230, 230, 230);height:60px" id="row<%#Container.DataItem("Departure_Id")%>">
										<td align="center">
											<a href="#" onclick="window.open('AddGuideMessages.asp?dID=<%#Container.DataItem("Departure_Id")%>','winCA','top=20, left=10, width=1000, height=350, scrollbars=1');">
												<img src="../../images/copy_icon.gif" border="0" alt="הוסף שיחה"></a></td>
										<td align="center" style="height:60px">
											<div class="div1">
												<a href="#"   class="link1" onclick="window.open('GuideMessages.asp?dID=<%#Container.DataItem("Departure_Id")%>','winCA','top=20, left=10, width=1200, height=800, scrollbars=1');">
													<div id="GuideMessages_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#func.GetCountGuideMessages(Container.DataItem("Departure_Id"))%></div>
												</a>
											</div>
										</td>
										<td align="center"><%#func.StatusOperation(Container.DataItem("Status"))%></td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="TimeMeetingAfterTrip_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="TimeMeetingAfterTrip_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("TimeMeetingAfterTrip")),"" , Container.DataItem("TimeMeetingAfterTrip"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="DateMeetingAfterTrip_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','DateMeetingAfterTrip','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<%#IIF(IsDBNull(Container.DataItem("DateMeetingAfterTrip")),"&nbsp;", DataBinder.Eval(Container.DataItem, "DateMeetingAfterTrip", "{0:dd/MM/yy}"))%>
													</a>
												</div>
												<div id="DateMeetingAfterTrip_row<%#Container.DataItem("Departure_Id")%>" style="display:none;" >
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','DateMeetingAfterTrip','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<img src="../../images/more.png" border="0" width="20" height="100%"></a>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Financial_Euro_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="FinancialEuro_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Financial_Euro")),"" , Container.DataItem("Financial_Euro"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Financial_Dolar_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="FinancialDolar_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Financial_Dolar")),"", Container.DataItem("Financial_Dolar"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Vouchers_Provider_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="Vouchers_Provider_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Vouchers_Provider")),"" , Container.DataItem("Vouchers_Provider"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px;padding-left:15px;">
											<div class="div1">
												<div id="Departure_Costing_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="Departure_Costing_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_Costing")),"" , Container.DataItem("Departure_Costing"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px;padding-left:15px;">
											<div class="div1">
												<div id="Voucher_Group_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="Voucher_Group_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Voucher_Group")),"" , Container.DataItem("Voucher_Group"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px;padding-left:15px;">
											<div class="div1">
												<div id="Voucher_Simultaneous_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="Voucher_Simultaneous_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Voucher_Simultaneous")),"" , Container.DataItem("Voucher_Simultaneous"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px;padding-left:15px;">
											<div class="div1">
												<div id="GilboaHotel_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="GilboaHotel_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("GilboaHotel")),"", Container.DataItem("GilboaHotel"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Departure_TimeGroupMeeting_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<div id="Departure_TimeGroupMeeting_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_TimeGroupMeeting")),"&nbsp;" , Container.DataItem("Departure_TimeGroupMeeting"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Departure_DateGroupMeeting_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateGroupMeeting','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<%#IIF(IsDBNull(Container.DataItem("Departure_DateGroupMeeting")),"&nbsp;", DataBinder.Eval(Container.DataItem, "Departure_DateGroupMeeting", "{0:dd/MM/yy}"))%>
													</a>
												</div>
												<div id="Departure_DateGroupMeeting_row<%#Container.DataItem("Departure_Id")%>" style="display:none;">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateGroupMeeting','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<img src="../../images/more.png" border="0" width="20"></a>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="DepartureTimeBrief_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="DepartureTimeBrief_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_TimeBrief")),"" , Container.DataItem("Departure_TimeBrief"))%></div>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Departure_DateBrief_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateBrief','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<%#IIF(IsDBNull(Container.DataItem("Departure_DateBrief")),"&nbsp;", DataBinder.Eval(Container.DataItem, "Departure_DateBrief", "{0:dd/MM/yy}"))%>
													</a>
												</div>
												<div id="Departure_DateBrief_row<%#Container.DataItem("Departure_Id")%>" style="display:none;">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateBrief','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<img src="../../images/more.png" border="0" width="20"></a>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Departure_Itinerary_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_Itinerary','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<%#IIF(IsDBNull(Container.DataItem("Departure_Itinerary")),"&nbsp;", DataBinder.Eval(Container.DataItem, "Departure_Itinerary", "{0:dd/MM/yy}"))%>
													</a>
												</div>
												<div id="Departure_Itinerary_row<%#Container.DataItem("Departure_Id")%>" style="display:none;">
													<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_Itinerary','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<img src="../../images/more.png" border="0" width="20"></a>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="supplier_Id_<%#Container.DataItem("Departure_Id")%>"  class="div2">
													<a href="#"  onclick="SelectSupplier('<%#Container.DataItem("Departure_Id")%>','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<%#IIF(IsDBNull(Container.DataItem("supplier_Name")),"&nbsp;", Container.DataItem("supplier_Name"))%>
													</a>
												</div>
												<div id="supplier_Id_row<%#Container.DataItem("Departure_Id")%>" style="display:none;">
													<a href="#"  onclick="SelectSupplier('<%#Container.DataItem("Departure_Id")%>','<%#func.vfix(Container.DataItem("Departure_Code"))%>');return false" class="link1">
														<img src="../../images/select.png" border="0" width="20"></a>
												</div>
											</div>
										</td>
										<td align="center" style="height:60px">
											<div class="div1">
												<div id="Departure_GuideTelphone_<%#Container.DataItem("Departure_Id")%>"  class="div2"><div id="Departure_GuideTelphone_row<%#Container.DataItem("Departure_Id")%>"><%#IIF(IsDBNull(Container.DataItem("Departure_GuideTelphone")),"", func.breaks(Container.DataItem("Departure_GuideTelphone")))%></div>
												</div>
											</div>
										</td>
										<td align="center"><%'#iif(container.DataItem("GuideName") is DBNULL.value,"",container.DataItem("GuideName") )%><%#Container.DataItem("GuideName")%></td>
										<td align="center"><%#func.TourStatusForOperation(Container.DataItem("Departure_Id"),Container.DataItem("Departure_Date"),func.dbNullFix(Container.DataItem("Departure_Date_End")))%></td>
										<td align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date_End", "{0:dd/MM/yy}")%></td>
										<td align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date", "{0:dd/MM/yy}")%></td>
										<td align="center"><%#Container.DataItem("Departure_Code")%></td>
										<td align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date", "{0:MMdd}")%></td>
										<td align="center"><%#Container.DataItem("Series_Name")%></td>
										<td align="right"><%#Container.DataItem("User_Name")%></td>
										<td align="center" style="width:25px"><input type="button" id="edit_button<%#Container.DataItem("Departure_Id")%>" value="עדכן" class="edit" onclick="edit_row('<%#Container.DataItem("Departure_Id")%>')" style="display: block;">
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
					<tr>
						<td height="2"></td>
					</tr>
					<tr>
						<td class="plata_paging" vAlign="top" align="center" colspan="7" bgcolor="#D8D8D8">
							<table dir="ltr" cellSpacing="0" cellPadding="2" width="100%" border="0">
								<tr>
									<td class="plata_paging" vAlign="baseline" noWrap align="left" width="150">&nbsp;הצג
										<asp:DropDownList id="PageSize" CssClass="PageSelect" Runat="server" AutoPostBack="true">
											<asp:ListItem Value="10">10</asp:ListItem>
											<asp:ListItem Value="20">20</asp:ListItem>
											<asp:ListItem Value="50" Selected="True">50</asp:ListItem>
										</asp:DropDownList>&nbsp;בדף&nbsp;
									</td>
									<td vAlign="baseline" noWrap align="right">
										<asp:LinkButton id="cmdNext" runat="server" CssClass="page_link" text="«עמוד הבא"></asp:LinkButton></td>
									<td class="plata_paging" vAlign="baseline" noWrap align="center" width="160">
										<asp:Label id="lblTotalPages" Runat="server"></asp:Label>&nbsp;דף&nbsp;
										<asp:DropDownList id="pageList" CssClass="PageSelect" Runat="server" AutoPostBack="true"></asp:DropDownList>&nbsp;מתוך&nbsp;
									</td>
									<td vAlign="baseline" noWrap align="left">
										<asp:LinkButton id="cmdPrev" runat="server" CssClass="page_link" text="עמוד קודם»"></asp:LinkButton></td>
									<td class="plata_paging" vAlign="baseline" align="right">&nbsp;נמצאו&nbsp;&nbsp;&nbsp;
										<asp:Label CssClass="plata_paging" id="lblCount" Runat="server"></asp:Label>&nbsp;יציאות
									</td>
								</tr>
							</table>
						</td>
					</tr>
				</asp:PlaceHolder>
			</table>
			</div></div>
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

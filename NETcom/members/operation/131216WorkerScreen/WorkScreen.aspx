<%@ Page Language="vb" AutoEventWireup="false" Codebehind="WorkScreen.aspx.vb" Inherits="bizpower_pegasus.WorkScreen"%>
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
{position: relative;top:50%;right:0px;-webkit-transform: translate(0%, -50%);transform : translate(0%, -50%);color:#000000}
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
<tr bgcolor=#d8d8d8 style="height:45px" >
		<td>הוסף <br>שיחה</td>
	<td class="title_sort" align=center>כמות <BR>שיחות</td>
	<td class="title_sort" align=center>?הסתיים<br>טיפול</td>
	<td class="title_sort" align=center>שעת פגישה <br>לאחר טיול</td>
	<td class="title_sort" align=center>תאריך פגישה <br>לאחר טיול</td>
	<td class="title_sort" align=center>הערכות כספית<br>-€</td>
	<td class="title_sort" align=center>הערכות כספית<br>-$</td>
	<td class="title_sort" align=center dir=rtl>שוברי <br>ספק קרקע?</td>
	<td class="title_sort" align=center>תמחיר</td>
	<td class="title_sort" align=center dir=rtl>שוברי<BR>הוצאות קבוצה<BR>ומדריך?</td>
	<td class="title_sort" align=center>שובר<br>?לסימולטני</td>
	<td class="title_sort" align=center>?הוקלדו<br>מלונות בגלבוע</td>
	<td class="title_sort" align=center>שעת<br>מפגש קבוצה</td>
	<td class="title_sort" align=center>תאריך <br>מפגש קבוצה</td>
	<td class="title_sort" style="padding:3px" align=center>שעת<BR>בריף</td>
	<td class="title_sort" align=center>תאריך<BR>בריף</td>
	<td class="title_sort" align=center>תאריך קבלת<br> Itinerary</td>
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
<a href="#"  onclick="TimeMeetingAfterTrip('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="TimeMeetingAfterTrip_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("TimeMeetingAfterTrip")),"&nbsp;","&nbsp;" & Container.DataItem("TimeMeetingAfterTrip"))%></div>
</a>
</div>
	</td>
		<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','DateMeetingAfterTrip','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="DateMeetingAfterTrip_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("DateMeetingAfterTrip")),"&nbsp;","&nbsp;" & Container.DataItem("DateMeetingAfterTrip"))%></div>
</a>
</div>
	</td>
	

	<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="Financial_Euro('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Financial_Euro_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Financial_Euro")),"&nbsp;","&nbsp;" & Container.DataItem("Financial_Euro"))%></div>
</a>
</div>
	</td>	
	<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="Financial_Dolar('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Financial_Dolar_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Financial_Dolar")),"&nbsp;","&nbsp;" & Container.DataItem("Financial_Dolar"))%></div>
</a>
</div>
	</td>
	<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditSelect('<%#Container.DataItem("Departure_Id")%>','Vouchers_Provider','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Vouchers_Provider_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Vouchers_Provider")),"&nbsp;","&nbsp;" & Container.DataItem("Vouchers_Provider"))%></div>
</a>
</div>
	</td>
			<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditSelect('<%#Container.DataItem("Departure_Id")%>','Departure_Costing','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Departure_Costing_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_Costing")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_Costing"))%></div>
</a>
</div>
	</td>
    			<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditSelect('<%#Container.DataItem("Departure_Id")%>','Voucher_Group','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Voucher_Group_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Voucher_Group")),"&nbsp;","&nbsp;" & Container.DataItem("Voucher_Group"))%></div>
</a>
</div>
	</td>
			<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditSelect('<%#Container.DataItem("Departure_Id")%>','Voucher_Simultaneous','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Voucher_Simultaneous_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Voucher_Simultaneous")),"&nbsp;","&nbsp;" & Container.DataItem("Voucher_Simultaneous"))%></div>
</a>
</div>
	</td>
	<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditSelect('<%#Container.DataItem("Departure_Id")%>','GilboaHotel','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="GilboaHotel_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("GilboaHotel")),"&nbsp;","&nbsp;" & Container.DataItem("GilboaHotel"))%></div>
</a>
</div>
	</td>
		<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="Departure_TimeGroupMeeting('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Departure_TimeGroupMeeting_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_TimeGroupMeeting")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_TimeGroupMeeting"))%></div>
</a>
</div>
	</td>
	<td  align=center  style="height:45px">
	<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateGroupMeeting','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Departure_DateGroupMeeting_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_DateGroupMeeting")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_DateGroupMeeting"))%></div>
</a>
</div>
</td>
	<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="DepartureTimeBrief('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="DepartureTimeBrief_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_TimeBrief")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_TimeBrief"))%></div>
</a>
</div>
	</td>	
				<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateBrief','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Departure_DateBrief_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_DateBrief")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_DateBrief"))%></div>
</a>
</div>
	</td>

				<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_Itinerary','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Departure_Itinerary_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_Itinerary")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_Itinerary"))%></div>
</a>
</div>
	</td>
			<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="SelectSupplier('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="supplier_Id_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("supplier_Name")),"&nbsp;","&nbsp;" & Container.DataItem("supplier_Name"))%></div>
</a>
</div>
	</td>
<td  align=center  style="height:45px;">
<div class="div1">
<a href="#"  onclick="GuideTel('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false"  class="link1">
<div id="GuideTel_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_GuideTelphone")),"&nbsp;","&nbsp;" & func.breaks(Container.DataItem("Departure_GuideTelphone")))%></div>
</a>
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
	</tr>
</ItemTemplate>
<AlternatingItemTemplate>
	<tr style="background-color: rgb(230, 230, 230);height:45px">
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
<a href="#"  onclick="TimeMeetingAfterTrip('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="TimeMeetingAfterTrip_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("TimeMeetingAfterTrip")),"&nbsp;","&nbsp;" & Container.DataItem("TimeMeetingAfterTrip"))%></div>
</a>
</div>
	</td>
			<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','DateMeetingAfterTrip','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="DateMeetingAfterTrip_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("DateMeetingAfterTrip")),"&nbsp;","&nbsp;" & Container.DataItem("DateMeetingAfterTrip"))%></div>
</a>
</div>
	</td>
	
		<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="Financial_Euro('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Financial_Euro_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Financial_Euro")),"&nbsp;","&nbsp;" & Container.DataItem("Financial_Euro"))%></div>
</a>
</div>
	</td>
	<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="Financial_Dolar('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Financial_Dolar_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Financial_Dolar")),"&nbsp;","&nbsp;" & Container.DataItem("Financial_Dolar"))%></div>
</a>
</div>
	</td>


<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditSelect('<%#Container.DataItem("Departure_Id")%>','Vouchers_Provider','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Vouchers_Provider_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Vouchers_Provider")),"&nbsp;","&nbsp;" & Container.DataItem("Vouchers_Provider"))%></div>
</a>
</div>
	</td>		
		<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditSelect('<%#Container.DataItem("Departure_Id")%>','Departure_Costing','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Departure_Costing_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_Costing")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_Costing"))%></div>
</a>
</div>
	</td>
    			<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditSelect('<%#Container.DataItem("Departure_Id")%>','Voucher_Group','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Voucher_Group_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Voucher_Group")),"&nbsp;","&nbsp;" & Container.DataItem("Voucher_Group"))%></div>
</a>
</div>
	</td>
			<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditSelect('<%#Container.DataItem("Departure_Id")%>','Voucher_Simultaneous','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Voucher_Simultaneous_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Voucher_Simultaneous")),"&nbsp;","&nbsp;" & Container.DataItem("Voucher_Simultaneous"))%></div>
</a>
</div>
	</td>
			<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditSelect('<%#Container.DataItem("Departure_Id")%>','GilboaHotel','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="GilboaHotel_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("GilboaHotel")),"&nbsp;","&nbsp;" & Container.DataItem("GilboaHotel"))%></div>
</a>
</div>
	</td>
	<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="Departure_TimeGroupMeeting('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Departure_TimeGroupMeeting_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_TimeGroupMeeting")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_TimeGroupMeeting"))%></div>
</a>
</div>
	</td>
	<td  align=center  style="height:45px">
	<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateGroupMeeting','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Departure_DateGroupMeeting_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_DateGroupMeeting")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_DateGroupMeeting"))%></div>
</a>
</div>
</td>
	<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="DepartureTimeBrief('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="DepartureTimeBrief_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_TimeBrief")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_TimeBrief"))%></div>
</a>
</div>
	</td>	
				<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_DateBrief','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Departure_DateBrief_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_DateBrief")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_DateBrief"))%></div>
</a>
</div>
	</td>

				<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="EditCalendar('<%#Container.DataItem("Departure_Id")%>','Departure_Itinerary','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="Departure_Itinerary_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_Itinerary")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_Itinerary"))%></div>
</a>
</div>
	</td>

			<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="SelectSupplier('<%#Container.DataItem("Departure_Id")%>','supplier_Id','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="supplier_Id_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("supplier_Name")),"&nbsp;","&nbsp;" & Container.DataItem("supplier_Name"))%></div>
</a>
</div>
	</td>

<td  align=center  style="height:45px">
<div class="div1">
<a href="#"  onclick="GuideTel('<%#Container.DataItem("Departure_Id")%>','<%#Container.DataItem("Departure_Code")%>');return false" class="link1">
<div id="GuideTel_<%#Container.DataItem("Departure_Id")%>"  class="div2"><%#IIF(IsDBNull(Container.DataItem("Departure_GuideTelphone")),"&nbsp;","&nbsp;" & Container.DataItem("Departure_GuideTelphone"))%></div>
</a>
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

    </form>

  </body>
</html>

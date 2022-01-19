<html>
<head>
<title>Calendar</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1255">
<link href="IE4.css" rel="STYLESHEET" type="text/css">
<%	obName = trim(Request.QueryString("obName"))
	obDate = Request.QueryString("obValue")%>
<style>
TD 
{ border : 1px solid #999999 }
</style>	
</head>
<script LANGUAGE="JavaScript">
<!--
function check_Date(objDate){
	var strtmp = new String(objDate);
	if (strtmp != '') {
		var arrmemb = new Array;
		arrmemb[0] = "";
		arrmemb[1] = "";
		arrmemb[2] = "";
		myarr = strtmp.split("/");
		for (cnt=0; cnt < myarr.length ; cnt++)
		{	arrmemb[cnt] = myarr[cnt]; }
	
		if (isNaN(arrmemb[0]) || isNaN(arrmemb[1]) || isNaN(arrmemb[2]) )
		{	alert("Incorrect Date !");
			objDate.select();	
			return false;}
		else
		{	var dd = arrmemb[0];
			var mm = arrmemb[1] - 1 ;
			var yy = arrmemb[2];
			var ValidDate = new Date(yy, mm, dd);
			if ((ValidDate.getDate()==dd) && (ValidDate.getMonth()==mm))
			{	return true; }
			else
			{	
				objDate.select();	
				return false;}
		}
	}
	return false;	
}

//************* start of calendar ******************
var today=new Date()
var currDay=today.getDate()
var currMonth=today.getMonth()
var currYear=today.getYear()

var currObjectDate = window.opener.document.all['<%=obName%>'];

var monames= new array("ינואר","פברואר","מרץ","אפריל","מאי",
                       "יוני","יולי","אוגוסט","ספטמבר","אוקטובר",
                       "נובמבר","דצמבר")

var days= new array(31,28,31,30,31,30,31,31,30,31,30,31)
var daysname= new array("'א","'ב","'ג","'ד","'ה","'ו","תבש")
var currColor="#FFCC33"

function visForm(objDate)
{
	var strDate = new String(objDate)
	if (check_Date(objDate)) 
	{	
		arrDate = strDate.split("/")
		currDay = arrDate[0]
		currMonth = arrDate[1] - 1
		currYear = arrDate[2]
	}
	else
	{
		currDay=today.getDate()
		currMonth=today.getMonth()
		currYear=today.getYear()
	}		
	showCalendar(currMonth,currYear);
}
function setDateObj(sDay,sMonth,sYear)
{
	currObjectDate.value=sDay+'/'+sMonth+'/'+sYear;
	this.close();
}

function daysBetween(dd,mmCur)
{
  var diff=Date.UTC(currYear,mmCur,dd,0,0,0)-Date.UTC(currYear,currMonth,currDay,0,0,0)
  return diff;
}

function array(m0,m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11)
{
  this[0]=m0; this[1]=m1; this[2]=m2; this[3]=m3; this[4]=m4;
  this[5]=m5; this[6]=m6; this[7]=m7; this[8]=m8; this[9]=m9;
  this[10]=m10; this[11]=m11; 
}

function showCalendar(MM,YY){
    document.tablemonth.namemonth.value =YY  + " " + monames[MM];
    firstDay=new Date(YY,MM,01);
    startDay=firstDay.getDay();
    if (((YY %4==0)&&(YY %100!=0))||(YY %400==0))
        days[1]=29;
    else
        days[1]=28;
	var yyyy = YY; 
	var mmmm = MM;
	var mrealy = ++mmmm;       
	for(i=0; i<6; i++)
    {	for(j=0; j<7; j++)
		{
			document.tablemonth.elements['iday'+i+j].value='';
			document.all['aday'+i+j].href='javascript:void(0);';
			document.tablemonth.elements['iday'+i+j].style.backgroundColor='#DBDBDB';
			document.all('tdday'+i+j).style.backgroundColor='#DBDBDB';
			document.tablemonth.elements['iday'+i+j].style.color='#000000';
		}	
    } 
    document.all['nextyear'].href='javascript:showCalendar(' + MM + ',' + (++yyyy) + ')';
    document.all['prevyear'].href='javascript:showCalendar(' + MM + ',' + (yyyy - 2) + ')';       
    yyyy = YY;
    var mmmm = MM;
    if (MM == 11) {
		mmmm = 0;
		yyyy++; }
	else 
		mmmm++;
    document.all['nextmonth'].href='javascript:showCalendar(' + mmmm + ',' + yyyy + ')';
    if (MM == 0) {
		mmmm = 11;
		yyyy = YY - 1;}
	else
		{mmmm = MM - 1;	
		yyyy = YY; }	
    document.all['prevmonth'].href='javascript:showCalendar(' + mmmm + ',' + yyyy + ')';       
    var row = 0;
    var column=0;
    for(i=0; i<startDay; i++)
    {
    document.tablemonth.elements['iday' + row + column].value="";
    column++;
    }
    for(i=1; i<=days[MM]; i++)
    {
    if (column==6)
    {
		if ((i==currDay)&&(MM==currMonth)&&(YY==currYear))
		{	document.tablemonth.elements['iday'+row+column].value=i;
			document.tablemonth.elements['iday'+row+column].style.backgroundColor='#FFCC33';
			document.all('tdday'+row+column).style.backgroundColor='#FFCC33';
			document.tablemonth.elements['iday'+row+column].style.color='#FFFFFF';
			document.all['aday'+row+column].href='javascript:setDateObj('+i+','+mrealy+','+YY+')';
		}	
        else
		{	document.tablemonth.elements['iday'+row+column].value=i;
			document.all['aday'+row+column].href='javascript:setDateObj('+i+','+mrealy+','+YY+')';
		}
		column=-1;
		row++; 
	}
    else
    {
         if ((i==currDay)&&(MM==currMonth)&&(YY==currYear))
		{	document.tablemonth.elements['iday'+row+column].value=i;
			document.tablemonth.elements['iday'+row+column].style.backgroundColor='#FFCC33';
			document.all('tdday'+row+column).style.backgroundColor='#FFCC33';
			document.tablemonth.elements['iday'+row+column].style.color='#000000';
			document.all['aday'+row+column].href='javascript:setDateObj('+i+','+mrealy+','+YY+')';
		}	
        else
		{	document.tablemonth.elements['iday'+row+column].value=i;
			document.all['aday'+row+column].href='javascript:setDateObj('+i+','+mrealy+','+YY+')';
		}  
     }
     column++;
	 }
} 

// *********** end of calendar **********
//-->
</script> 
<body style="margin:10px;background-color:#FFFFFF;margin-top:15px;">
 
<!--  LAYER  1-----    -->
<div name="formCal" id="formCal" style="POSITION: absolute;  WIDTH: 150px; Z-INDEX:500; CLEAR: both;">                  
	<table bgcolor="#DBDBDB" cellpadding="0" cellspacing="0"  style="border-collapse:collapse;border : 1px solid #999999"  width="150" align=center>
		<form action method="POST" id="tablemonth" name="tablemonth">
		<tr>
			<td valign="middle" style="height:25" align=center><a HREF="javascript:window.close();"><img SRC="images/x.gif" BORDER="0" HSPACE="3" alt="Close" align="middle" WIDTH="15" HEIGHT="15"></a><a id="nextyear" name="nextyear" HREF="javascript:void(0)"><img SRC="images/ar_left.gif" BORDER="0" HSPACE="2" alt="Next Year" WIDTH="8" HEIGHT="8"></a><a id="nextmonth" name="nextmonth" HREF="javascript:void(0)"><img SRC="images/ar_left1.gif" BORDER="0" HSPACE="2" alt="Next Month" WIDTH="6" HEIGHT="8"></a><input class="taarih_month" type="text" id="namemonth" name="namemonth" size="12" readonly><a id="prevmonth" name="prevmonth" HREF="javascript:void(0)"><img SRC="images/ar_right1.gif" BORDER="0" HSPACE="2" alt="Previous Month" WIDTH="6" HEIGHT="8"></a><a id="prevyear" name="prevyear" HREF="javascript:void(0)"><img SRC="images/ar_right.gif" BORDER="0" HSPACE="2" alt="Previous Year" WIDTH="8" HEIGHT="8"></a></td>
		</tr>
		<tr>
			<td valign="middle" align="center">	
				<table style="border-collapse:collapse;border : 1px solid #999999" cellpadding="0" cellspacing="0" style="width:100%">
					<tr>
						<td bgcolor="#B0B0B0" align="center"><font style="font-size:8pt;FONT-FAMILY: Arial;" color="white">'א</font></td>
						<td bgcolor="#B0B0B0" align="center"><font style="font-size:8pt;FONT-FAMILY: Arial;" color="white">'ב</font></td>
						<td bgcolor="#B0B0B0" align="center"><font style="font-size:8pt;FONT-FAMILY: Arial;" color="white">'ג</font></td>
						<td bgcolor="#B0B0B0" align="center"><font style="font-size:8pt;FONT-FAMILY: Arial;" color="white">'ד</font></td>
						<td bgcolor="#B0B0B0" align="center"><font style="font-size:8pt;FONT-FAMILY: Arial;" color="white">'ה</font></td>
						<td bgcolor="#B0B0B0" align="center"><font style="font-size:8pt;FONT-FAMILY: Arial;" color="white">'ו</font></td>
						<td bgcolor="#FFCC33" align="center"><font style="font-size:8pt;FONT-FAMILY: Arial;" color="#000000">שבת</font></td>
					</tr>
					<tr>
						<td align="center" id="tdday00" name="tdday00"><a width=100% id="aday00" name="aday00" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday00" name="iday00" size="2" readonly></a></td>
						<td align="center" id="tdday01" name="tdday01"><a width=100% id="aday01" name="aday01" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday01" name="iday01" size="2" readonly></a></td>
						<td align="center" id="tdday02" name="tdday02"><a width=100% id="aday02" name="aday02" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday02" name="iday02" size="2" readonly></a></td>
						<td align="center" id="tdday03" name="tdday03"><a width=100% id="aday03" name="aday03" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday03" name="iday03" size="2" readonly></a></td>
						<td align="center" id="tdday04" name="tdday04"><a width=100% id="aday04" name="aday04" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday04" name="iday04" size="2" readonly></a></td>
						<td align="center" id="tdday05" name="tdday05"><a width=100% id="aday05" name="aday05&quot;" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday05" name="iday05" size="2" readonly></a></td>
						<td align="center" id="tdday06" name="tdday06"><a width=100% id="aday06" name="aday06&quot;" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday06" name="iday06" size="2" readonly></a></td>
					</tr>
					<tr>
						<td align="center" id="tdday10" name="tdday10"><a width=100% id="aday10" name="aday10" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday10" name="iday10" size="2" readonly></a></td>
						<td align="center" id="tdday11" name="tdday11"><a width=100% id="aday11" name="aday11" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday11" name="iday11" size="2" readonly></a></td>
						<td align="center" id="tdday12" name="tdday12"><a width=100% id="aday12" name="aday12" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday12" name="iday12" size="2" readonly></a></td>
						<td align="center" id="tdday13" name="tdday13"><a width=100% id="aday13" name="aday13" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday13" name="iday13" size="2" readonly></a></td>
						<td align="center" id="tdday14" name="tdday14"><a width=100% id="aday14" name="aday14" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday14" name="iday14" size="2" readonly></a></td>
						<td align="center" id="tdday15" name="tdday15"><a width=100% id="aday15" name="aday15" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday15" name="iday15" size="2" readonly></a></td>
						<td align="center" id="tdday16" name="tdday16"><a width=100% id="aday16" name="aday16" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday16" name="iday16" size="2" readonly></a></td>
					</tr>
					<tr>
						<td align="center" id="tdday20" name="tdday20"><a width=100% id="aday20" name="aday20" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday20" name="iday20" size="2" readonly></a></td>
						<td align="center" id="tdday21" name="tdday21"><a width=100% id="aday21" name="aday21" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday21" name="iday21" size="2" readonly></a></td>
						<td align="center" id="tdday22" name="tdday22"><a width=100% id="aday22" name="aday22" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday22" name="iday22" size="2" readonly></a></td>
						<td align="center" id="tdday23" name="tdday23"><a width=100% id="aday23" name="aday23" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday23" name="iday23" size="2" readonly></a></td>
						<td align="center" id="tdday24" name="tdday24"><a width=100% id="aday24" name="aday24" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday24" name="iday24" size="2" readonly></a></td>
						<td align="center" id="tdday25" name="tdday25"><a width=100% id="aday25" name="aday25" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday25" name="iday25" size="2" readonly></a></td>
						<td align="center" id="tdday26" name="tdday26"><a width=100% id="aday26" name="aday26" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday26" name="iday26" size="2" readonly></a></td>
					</tr>
					<tr>
						<td align="center" id="tdday30" name="tdday30"><a width=100% id="aday30" name="aday30" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday30" name="iday30" size="2" readonly></a></td>
						<td align="center" id="tdday31" name="tdday31"><a width=100% id="aday31" name="aday31" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday31" name="iday31" size="2" readonly></a></td>
						<td align="center" id="tdday32" name="tdday32"><a width=100% id="aday32" name="aday32" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday32" name="iday32" size="2" readonly></a></td>
						<td align="center" id="tdday33" name="tdday33"><a width=100% id="aday33" name="aday33" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday33" name="iday33" size="2" readonly></a></td>
						<td align="center" id="tdday34" name="tdday34"><a width=100% id="aday34" name="aday34" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday34" name="iday34" size="2" readonly></a></td>
						<td align="center" id="tdday35" name="tdday35"><a width=100% id="aday35" name="aday35" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday35" name="iday35" size="2" readonly></a></td>
						<td align="center" id="tdday36" name="tdday36"><a width=100% id="aday36" name="aday36" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday36" name="iday36" size="2" readonly></a></td>
					</tr>
					<tr>
						<td align="center" id="tdday40" name="tdday40"><a width=100% id="aday40" name="aday40" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday40" name="iday40" size="2" readonly></a></td>
						<td align="center" id="tdday41" name="tdday41"><a width=100% id="aday41" name="aday41" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday41" name="iday41" size="2" readonly></a></td>
						<td align="center" id="tdday42" name="tdday42"><a width=100% id="aday42" name="aday42" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday42" name="iday42" size="2" readonly></a></td>
						<td align="center" id="tdday43" name="tdday43"><a width=100% id="aday43" name="aday43" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday43" name="iday43" size="2" readonly></a></td>
						<td align="center" id="tdday44" name="tdday44"><a width=100% id="aday44" name="aday44" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday44" name="iday44" size="2" readonly></a></td>
						<td align="center" id="tdday45" name="tdday45"><a width=100% id="aday45" name="aday45" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday45" name="iday45" size="2" readonly></a></td>
						<td align="center" id="tdday46" name="tdday46"><a width=100% id="aday46" name="aday46" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday46" name="iday46" size="2" readonly></a></td>
					</tr>
					<tr>
						<td align="center" id="tdday50" name="tdday50"><a width=100% id="aday50" name="aday50" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday50" name="iday50" size="2" readonly></a></td>
						<td align="center" id="tdday51" name="tdday51"><a width=100% id="aday51" name="aday51" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday51" name="iday51" size="2" readonly></a></td>
						<td align="center" id="tdday52" name="tdday52"><a width=100% id="aday52" name="aday52" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday52" name="iday52" size="2" readonly></a></td>
						<td align="center" id="tdday53" name="tdday53"><a width=100% id="aday53" name="aday53" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday53" name="iday53" size="2" readonly></a></td>
						<td align="center" id="tdday54" name="tdday54"><a width=100% id="aday54" name="aday54" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday54" name="iday54" size="2" readonly></a></td>
						<td align="center" id="tdday55" name="tdday55"><a width=100% id="aday55" name="aday55" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday55" name="iday55" size="2" readonly></a></td>
						<td align="center" id="tdday56" name="tdday56"><a width=100% id="aday56" name="aday56" href="javascript:void(0);"><input class="taarih_day" type="text" id="iday56" name="iday56" size="2" readonly></a></td>
					</tr>
			    </table>
			</td>
		</tr>
		</form>
	</table>
</div>
<!--end layer-->    
<script LANGUAGE="javascript">
<!--
	visForm(<%=obDate%>);
//-->
</script>
</body>
</html>

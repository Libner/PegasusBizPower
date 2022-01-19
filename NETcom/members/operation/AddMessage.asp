<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<script>
function selectToursDate(comboSeriaId) 
{ 
   SeriaId = comboSeriaId.value;
 var xmlhttp;
if (window.XMLHttpRequest)
  {// code for IE7+, Firefox, Chrome, Opera, Safari
  xmlhttp=new XMLHttpRequest();
  }
else
  {// code for IE6, IE5
  xmlhttp=new ActiveXObject("Microsoft.XMLHTTP");
  }
  
      if(isNaN(SeriaId) == false)
   { // var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");

	 xmlhttp.open("POST", 'selectToursDate.asp', false);
	 xmlhttp.send('<?xml version="1.0" encoding="UTF-8"?><request><SeriaId>' + SeriaId + '</SeriaId></request>');	 
	 
			result = new String(xmlhttp.responseText);						
	
		arr_ToursDate= result.split(";");		
		window.document.all("ToursDate_Id").length = 1;
	
	
			for (i=0; i < arr_ToursDate.length-1; i++)
		{			
			arr_TourDate = new String(arr_ToursDate[i]);
			arr_TourDate = arr_TourDate.split(",")			
			window.document.all("ToursDate_Id").options[window.document.all("ToursDate_Id").options.length]=
            new Option(arr_TourDate[1],arr_TourDate[0]); 
	    }
	    
	 
		
	  //document.getElementById("ToursDate_Id").style="dispaly:block"
	 }
}
</script>
<script LANGUAGE="JavaScript">
<!--
function CheckFields()
{	
if ( document.frmMain.selSeries.value=='')		   
	{
				alert("'*' אנא מלאו כל השדות הצוינות בסימן")
				return false;				
		}
		if ( document.frmMain.ToursDate_Id.value=='')		   
	{
				alert("'*' אנא מלאו כל השדות הצוינות בסימן")
				return false;				
		}
	if ( document.frmMain.Messages_Content.value==''
	     )		   
		{
				alert("'*' אנא מלאו כל השדות הצוינות בסימן")
				return false;				
		}
	else
		{
			document.frmMain.submit();
			return true;
		}		
}

function exit_()
	{
		this.close();	
		return false;
	}

//-->
</script>  
<%if request.form("Messages_Content")<>nil then 'after form filling
 UserID=trim(Request.Cookies("bizpegasus")("UserID"))
	 Messages_Content = sFix(request.form("Messages_Content"))
	 departureId= sFix(request.form("ToursDate_Id"))
		con.executeQuery("SET DATEFORMAT dmy")
		sqlStr = "insert into GuideMessages (Departure_Id,User_Id,Messages_Content) "
		sqlStr=sqlStr& " values (" & departureId &",'"& UserID & "','" & Messages_Content &"')"
		'Response.Write sqlStr
		'response.end
		con.GetRecordSet(sqlStr)	%>				

     <SCRIPT LANGUAGE=javascript>
	 <!--	
		opener.focus();
		opener.window.location.reload(true);
		self.close();
	 //-->
	</SCRIPT>      
  <%  end if
  dateStart = Day(date()) & "/" & Month(date()) & "/" & Year(date())
  %>
<body>


<table border="0" cellpadding="0" cellspacing="0" width="100%" ID="Table1">
<tr>
   <td align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0" ID="Table8">
	  <tr><td class="page_title">הוספת&nbsp;שיחה&nbsp;</td></tr>		   
	  		       	
   </table></td></tr>  

<tr><td width=100% bgcolor="#E6E6E6"> 
<table align=center border="0" cellpadding="3" cellspacing="1" width="500" ID="Table2">    

<FORM name="frmMain" ACTION="AddMessage.asp" METHOD="post" onSubmit="return CheckFields()" ID="Form1">
<tr>
	<td align=right><%=vfix(dateStart)%></td>
	<td align="right" nowrap width="100">תאריך שיחה&nbsp;</td>
</tr>
<tr>
	<td align=right>	
			<select name="selSeries"  dir="<%=dir_obj_var%>" class="sel" style="width: 250px" ID="selSeries"  onchange="selectToursDate(this)" >
			<option value="0">בחר סידרה</option>
			
			<%set SeriesList=con.GetRecordSet("select  Series_Id,Series_Name from Series ORDER BY Series_Name")
			do while not SeriesList.EOF
			selSeriesID=SeriesList(0)
			selSeriesName=SeriesList(1)%>
			
			<option value="<%=selSeriesID%>"> <%=selSeriesName%></option>
	<% SeriesList.MoveNext
			Loop
			Set SeriesList=Nothing%>
			</select>	</td>
		
	<td align="right" nowrap width="100">סידרה&nbsp;*</td>
</tr>
<tr>
	<td align=right>	<select name="ToursDate_Id" id="ToursDate_Id" class="sel" dir=rtl  style="width:380px;font-size:10pt;display:block;">
	<option value="0">בחר טיול</option>
			
	</select>
	</td>
	<td align="right" nowrap width="100" valign=top>טיול&nbsp;*</td>
<tr>
	<td align=right><textarea dir=rtl name="Messages_Content" style="width:300px;height:200px;font-family:arial" ID="Messages_Content"><%=vfix(Messages_Content)%></textarea></td>
	<td align="right" nowrap width="100" valign=top>פרוט שיחה&nbsp;*</td>
</tr>




<tr><td colspan="2" height=10></td></tr>

<tr>
<td colspan=2>
<table cellpadding=0 cellspacing=0 align=center ID="Table3">
<tr>
<td width=50% align=right><input type=button class="but_menu" style="width:90px" onclick="return exit_();" value="ביטול" ID="Button1" NAME="Button1"></td>
<td width=50 nowrap></td>
<td width=50% align=left><input type=button class="but_menu" style="width:90px" onclick="javascript:CheckFields(); return false;" value="אישור" ID="Button2" NAME="Button2"></td></tr>
</table></td></tr>
</form>
<tr><td colspan="2" height=10></td></tr>
</table>
</td></tr></table>
</body>
<%
set con = nothing
%>
</html>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="calendarAct.aspx.vb" Inherits="bizpower_pegasus.calendarAct" %>
<!DOCTYPE HTML>
<html>
  <head>
    <title>gradesByDepId</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
<script>
function EditStatus(n)
{

var  formData = "year="+<%=currYear%>+"&selV="+document.getElementById("SelStatus").value; 
//alert (formData);
 
$.ajax({
    url : "UpdateYearStatus.aspx",
    type: "POST",
    async: "false",
    data : formData,
      success: function(data)
    {
    var t=data;
    //alert(t);
    document.getElementById("statusName").innerHTML=t;
   
 // alert("data")
  
         //data - response from server
    },
    error: function (jqXHR, textStatus, errorThrown)
    {
 //alert("textStatus")
    }
});

}
function edit_row(value)
{
document.getElementById("dvT_"+value).style.display="none";
document.getElementById("dv_"+value).style.display="block";

//alert(value)
}


	function ChangeSelect(value,IsHoliday,WeekDayNumber)
		{
		//alert(document.getElementById("sel_"+ value).value)
	var  formData = "name="+value+"&selV="+document.getElementById("sel_"+ value).value+"&isH="+IsHoliday+"&wday="+WeekDayNumber; 
//alert (formData);
 
$.ajax({
    url : "UpdateDataDay.aspx",
    type: "POST",
    data : formData,
    success: function(data, textStatus, jqXHR)
    {
  //alert(data)
   // var ResData new Array();
    var ResData=data.split("_");
    
 //alert(ResData[0]);
 document.getElementById("typeName_"+value).innerHTML=ResData[0];
 document.getElementById("Td_"+value).style="height:50px;color:"+ResData[2]+";background-color:"+ResData[1]
 document.getElementById("dv_"+value).style.display="none";
		document.getElementById("dvT_"+value).style.display="block";
        //data - response from server
    },
    error: function (jqXHR, textStatus, errorThrown)
    {
 
    }
});
		
		
		}
function edit_row(value)
{
document.getElementById("dvT_"+value).style.display="none";
document.getElementById("dv_"+value).style.display="block";

//alert(value)
}

</script>
  </head>
  <body style="margin:0px">

    <form id="Form1" method="post" runat="server">
    <table border=0 cellpadding=0 cellspacing=0 width=100% align=left>
    <tr><td width=100% align=center><table border=0 cellpadding=0 cellspacing=0>
   <tr><td width=200></td><td><a   class="but_menu" style="width:90;cursor:pointer" id="ButStatus" name="ButStatus" onclick="EditStatus(this);return false">שמור סטטוס</a></td><td width=20></td><td>
   <select id=SelStatus name=SelStatus><option value="0" <%if YearStatus<>"1" then%>selected<%end if%> style="background-color:#FFCE42">מלאתי את העמוד חלקית</option>
   <option value="1" <%if YearStatus="1" then%>selected<%end if%> style="background-color:#7CFC00">סיימתי בקפדנות</option></select>
   </select></td><td width=200></td>
   <td nowrap style="color:#6F6DA6;font-weight:600;font-size:18px;vertical-align:middle;padding-right:10px" align=center>  סטטוס שנה <%=currYear%> : <div id="statusName" style="display: inline-block;display: inline;"><%=YearStatusName%></div></td>
   </tr></table></td></tr>
   <tr>
<td width=100% style="color:#000000;background-color:#e1e1e1e1;"> 
<table border=0 cellpadding=0 cellpadding=0 width=100% style="color:#000000;background-color:#E6E6E6;">
<tr><td align=center colspan=13><span style="font-size:14pt"><select id="y" name="y" onchange="this.form.submit()">
<%for yy=2014 to 2040%>
<option value="<%=yy%>" <%if currYear=yy then%>selected<%end if%> ><%=yy%></option>
<%next%>
</select> </span></td>   </tr>
<tr>
<asp:Repeater ID=rptMonth Runat=server>
<ItemTemplate>
<td width=8% valign=top><table border=0 cellspacing=1 cellpadding=0 width=100%>
<tr>
<th align=center><%#MonthName(Container.DataItem) %></th></tr>
<asp:Repeater ID=rptDays Runat=server>
<ItemTemplate>

<tr><td id="Td_<%#Container.DataItem("DateKey")%>" onclick="edit_row('<%#Container.DataItem("DateKey")%>')" style="height:50px;color:<%#Container.DataItem("DayFontColor")%>;background-color:<%#Container.DataItem("DayTypeColor")%>" align=center><%#Container.DataItem("WeekdayNameHeb")%><br><%#Container.DataItem("HolidayName")%><br>
<div id="dvT_<%#Container.DataItem("DateKey")%>" >
<div id="typeName_<%#Container.DataItem("DateKey")%>"><%#Container.DataItem("DayTypeName")%></div>
</div>
<div id="dv_<%#Container.DataItem("DateKey")%>" style="display:none">
<select id="sel_<%#Container.DataItem("DateKey")%>" onchange="javascript:ChangeSelect('<%#Container.DataItem("DateKey")%>','<%#Container.DataItem("IsHoliday")%>','<%#Container.DataItem("WeekdayNumber")%>');">
<option value="1" <%#IIF(Container.DataItem("DayTypeName")="יום עבודה מלא","selected","")%>  style="background-color:#7CFC00">יום עבודה מלא</option>
<option value="2" <%#IIF(Container.DataItem("DayTypeName")="חצי יום עבודה","selected","")%> style="background-color:#FFC0CB">חצי יום עבודה</option>
<option value="3" <%#IIF(Container.DataItem("DayTypeName")="אין עבודה","selected","")%> style="background-color:#FF0000">אין עבודה</option>
</select></div>
</td>
</tr>
<tr><td colspan=13 height=3></td></tr>
</ItemTemplate></asp:Repeater>

</table></td>

</ItemTemplate>
</asp:Repeater>

<td width=4% valign=top bgcolor=#ffffff><table border=0 cellspacing=1 cellpadding=0 width=100%>
<tr><th >יום</th>
</tr>
<%For i=1 to 31 %>
<tr><td align=center bgcolor=#BBBAD6 style="height:50px"><%=i%></td>
</tr>
<tr><td colspan=13 height=3></td></tr>

<%next%>

</table></td>

</tr>
<tr><td colspan=13 height=5>&nbsp;</td></tr>

<tr><td colspan=13>

</td></tr>

</table>

    </form>

  </body>
</html>

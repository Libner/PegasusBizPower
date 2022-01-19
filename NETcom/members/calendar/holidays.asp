<!--#include file="../../connect.asp"-->
<%Response.Buffer = True%>
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<script>
	function checkDelete()
	{	 
		return window.confirm("?האם ברצונך לאפס את הנתונים");		
	}
</script>
<%

If Request("y")="" then
currYear=year(Now())
else
currYear=Request("y")
end if

%>


<html>
	<head>
		<!-- #include file="../../title_meta_inc.asp" -->
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<script src="../../include/CalendarPopup.js" charset="windows-1255"></script>
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
		<script>
function ChangeYear()
{
document.form1.submit()
}
function checkForm()
{
var flag = 1;
$(".Myinput").each(function(i){
    if ($(this).val() == "")
        flag++;
});

if (flag == 1)
{
  var input = document.createElement('input');
    input.type = 'hidden';
    input.name = 'add'; // 'the key/name of the attribute/field that is sent to the server
    input.value = 1;
    document.form1.appendChild(input);

   document.form1.submit()
}
else
{ alert('נא למלא כל התאריכים'); return false;}
}

		</script>
	</head>
	<body>
		<FORM action="holidays.asp" method="post" id="form1" name="form1">
			<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table1">
				<tr>
					<td width="100%" align="<%=align_var%>">
						<!--#include file="../../logo_top.asp"-->
					</td>
				</tr>
				<tr>
					<td width="100%" align="<%=align_var%>">
						<%numOftab = 4%>
						<%numOfLink = 15%>
<%topLevel2 = 104 'current bar ID in top submenu - added 03/10/2019%>
						<!--#include file="../../top_in.asp"-->
					</td>
				</tr>
				<tr>
					<td class="page_title">&nbsp;
						<%'=arrTitles(25)%>
					</td>
				</tr>
				<tr>
					<td width="100%" valign="top" align="center">
						<table dir="<%=dir_var%>" border="0" cellpadding="0" cellspacing="1" width="100%" ID="Table2">
							<tr>
								<td width="100%" style="color:#000000;background-color:#e1e1e1e1;">
									<table border="0" cellpadding="1" cellpadding="1" width="100%" style="color:#000000;background-color:#E6E6E6;"
										ID="Table3">
										<tr>
											<td align="left"></td>
											<td align="center" width="100%"><span style="font-size:14pt">
													<select id="y" name="y" onchange="javascript:ChangeYear()">
														<%for yy=2014 to 2040%>
														<option value="<%=yy%>" <%if Cint(currYear)=yy then%>selected<%end if%> ><%=yy%></option>
														<%next%>
													</select>
												</span>
											</td>
										</tr>
										<tr>
									</table>
								</td>
							</tr>
							<tr>
								<td align="center">
									<table border="1" cellpadding="3" cellspacing="3">
										<tr>
											
											<td align="center">תאריך קלנדרי לפי שנת
												<%=currYear%>
											</td>
											<td align="center">הגדרת העבודה ביום זה</td>
											<td align="center">שם מועד/ חג
											</td>
										</tr>
										<%'sqlstr = "Select * from DaysInYear Where Year = " & currYear 	
										sqlstr = "Select * from Holidays Where Year = " & currYear  &" order by DayOrder"
	set rsdays = con.getRecordSet(sqlstr)		
	If not rsdays.eof Then
		do while not rsdays.eof%>
										<tr>
											<td align="center"><table border="0" cellpadding="0" cellspacing="0" ID="Table4">
													<tr>
														<td align="right" dir="rtl"><%if rsdays("DayAction")="1" then%><a href="" onclick="cal1xx.select(document.getElementById('Date_<%=rsdays("Id")%>'),'AsDate_<%=rsdays("Id")%>','dd/MM/yyyy'); return false;" id="A1">
																<img src="../../images/calendar.gif" border="0" align="center"></a> <input type="text" name="Date_<%=rsdays("Id")%>" id="Text1" size="10" value="<%=rsdays("DayDate")%>" readonly dir="rtl" class=Myinput>
															<%else%>
															<%=Day(rsdays("Date"))%>/<%=Month(rsdays("Date"))%>/<%=year(rsdays("Date"))%>
															<%end if%>
														</td>
													</tr>
												</table>
											</td>
											<td align=right style="background-color:<%=rsdays("DayTypeColor")%>"><%=rsdays("DayTypeName")%></td>
											<td align="right" nowrap style="background-color:#e1e1e1" dir="rtl"><%=rsdays("DayName")%></td>
										</tr>
										<%		rsdays.Movenext
		loop
	
	else
	sqlstr = "Select * from DaysTemplate order by DayOrder"
	set rsdaysT = con.getRecordSet(sqlstr)	
	do while not rsdaysT.eof%>
										<tr>
											<td align="center" style="color:#001BC0"><%if IsNumeric(rsdaysT("DayCalculate")) then%>+<%=rsdaysT("DayCalculate")%><%end if%></td>
											<td><table border="0" cellpadding="0" cellspacing="0">
													<tr>
														<td align="right" dir="rtl"><%if rsdaysT("DayAction")="1" then%><a href="" onclick="cal1xx.select(document.getElementById('Date_<%=rsdaysT("Id")%>'),'AsDate_<%=rsdaysT("Id")%>','dd/MM/yyyy'); return false;" id="AsDate_<%=rsdaysT("Id")%>">
																<img src="../../images/calendar.gif" border="0" align="center"></a> <input type="text" name="Date_<%=rsdaysT("Id")%>" id="Date_<%=rsdaysT("Id")%>" size="10" value="" readonly dir="rtl" class=Myinput>
															<%end if%>
														</td>
													</tr>
												</table>
											</td>
											<td align=right style="background-color:<%=rsdaysT("DayTypeColor")%>"><%=rsdaysT("DayTypeName")%></td>
											<td align="right" style="background-color:#e1e1e1" dir="rtl"><%=rsdaysT("DayName")%></td>
										</tr>
										<%	rsdaysT.Movenext
		loop
		set rsdaysT = Nothing
	End If
	set rsdays = Nothing%>
									</table>
								</td>
							</tr>
						</table>
		</FORM>
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
	</body>
</html>
<%set con=Nothing%>

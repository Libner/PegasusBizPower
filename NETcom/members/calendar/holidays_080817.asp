<%Response.Buffer = True%>
<!--#include file="../../connect.asp"-->
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
If Request.QueryString("deleteId") <> nil Then
		delId = trim(Request.QueryString("deleteId"))
		If Len(delId) > 0 Then		   
			con.executeQuery("Delete From DaysInYear Where year = " & delId)
		End If
		Response.Redirect "holidays.asp?y="& delId
		
	End If
response.Write "aaa="& request.Form ("add")
%>
<%if request.Form ("add")="1" then
'For Each Item In Request.Form
'    fieldName = Item
'    fieldValue = Request.Form(Item)
'    Response.Write( fieldName &" = "& fieldValue)       
'  Next 
    sqlstr = "Select * from DaysInYear Where Year = " & currYear 				
	set rsdays = con.getRecordSet(sqlstr)		
	If not rsdays.eof Then
	sqlDelete ="Delete from DaysInYear Where Year = " & currYear 	
	con.executeQuery(sqlDelete)
	end if
	response.Write ("ddrdd")
	for i=1 to 45 
	select case i
	case "1"
	'sqlIns="SET DATEFORMAT DMY;	Insert into DaysInYear (Year,DayDate,DayType) values ('"& currYear &"','"& Request.Form("Date_1") &"','1')"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 1,'"& currYear &"','"& Request.Form("Date_1") &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=1" 
	con.executeQuery (sqlIns)
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 2,'"& currYear &"','"& DateAdd("d",1,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=2" 
	con.executeQuery (sqlIns)
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 3,'"& currYear &"','"& DateAdd("d",2,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=3" 
	con.executeQuery (sqlIns)

	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 4,'"& currYear &"','"& DateAdd("d",9,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=4" 
	con.executeQuery (sqlIns)

	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 5,'"& currYear &"','"& DateAdd("d",10,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=5" 
	con.executeQuery (sqlIns)
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 6,'"& currYear &"','"& DateAdd("d",14,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=6" 
	con.executeQuery (sqlIns)

   sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 7,'"& currYear &"','"& DateAdd("d",15,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=7" 
	con.executeQuery (sqlIns)
   sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 8,'"& currYear &"','"& DateAdd("d",16,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=8"
 	con.executeQuery (sqlIns)

      sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 9,'"& currYear &"','"& DateAdd("d",17,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=9" 
	con.executeQuery (sqlIns)
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 10,'"& currYear &"','"& DateAdd("d",18,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=10" 
	con.executeQuery (sqlIns)
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 11,'"& currYear &"','"& DateAdd("d",19,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=11" 
	con.executeQuery (sqlIns)
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 12,'"& currYear &"','"& DateAdd("d",20,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=12" 
	con.executeQuery (sqlIns)
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 13,'"& currYear &"','"& DateAdd("d",21,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=13" 
	con.executeQuery (sqlIns)
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 14,'"& currYear &"','"& DateAdd("d",22,Request.Form("Date_1"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=14" 
	con.executeQuery (sqlIns)
case "15"   
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 15,'"& currYear &"','"& Request.Form("Date_15")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=15" 
	con.executeQuery (sqlIns)

    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 16,'"& currYear &"','"& DateAdd("d",1,Request.Form("Date_15"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=16" 
	con.executeQuery (sqlIns)
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 17,'"& currYear &"','"& DateAdd("d",2,Request.Form("Date_15"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=17" 
	con.executeQuery (sqlIns)
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 18,'"& currYear &"','"& DateAdd("d",3,Request.Form("Date_15"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=18" 
	con.executeQuery (sqlIns)
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 19,'"& currYear &"','"& DateAdd("d",4,Request.Form("Date_15"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=19" 
	con.executeQuery (sqlIns)
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 20,'"& currYear &"','"& DateAdd("d",5,Request.Form("Date_15"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=20" 
	con.executeQuery (sqlIns)
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 21,'"& currYear &"','"& DateAdd("d",6,Request.Form("Date_15"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=21" 
	con.executeQuery (sqlIns)
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 22,'"& currYear &"','"& DateAdd("d",7,Request.Form("Date_15"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=22" 
	con.executeQuery (sqlIns)
case "23"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 23,'"& currYear &"','"& Request.Form("Date_23")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=23" 
	con.executeQuery (sqlIns)

case "24"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 24,'"& currYear &"','"& Request.Form("Date_24")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=24" 
	con.executeQuery (sqlIns)
case "25"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 25,'"& currYear &"','"& Request.Form("Date_25")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=25" 
	con.executeQuery (sqlIns)  
case "26"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 26,'"& currYear &"','"& Request.Form("Date_26")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=26" 
	con.executeQuery (sqlIns)  
 case "27"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 27,'"& currYear &"','"& Request.Form("Date_27")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=27" 
	con.executeQuery (sqlIns)  
j=1
for k=28 to 34

dayN=DateAdd("d", j  ,Request.Form("Date_27")) 
  sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select "& k &",'"& currYear &"','"& dayN & "',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id="& k 
'response.Write sqlIns
con.executeQuery (sqlIns)
j=j+1
next
	'	response.Write sqlIns
	case "35"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 35,'"& currYear &"','"& Request.Form("Date_35")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=35" 
	con.executeQuery (sqlIns)  
case "36"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 36,'"& currYear &"','"& Request.Form("Date_36")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=36" 
	con.executeQuery (sqlIns)  
case "37"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 37,'"& currYear &"','"& Request.Form("Date_37")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=37" 
	con.executeQuery (sqlIns)  
	case "38"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 38,'"& currYear &"','"& Request.Form("Date_38")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=38" 
	con.executeQuery (sqlIns)  
	case "39"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 39,'"& currYear &"','"& Request.Form("Date_39")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=39" 
	con.executeQuery (sqlIns)  
	case "40"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 40,'"& currYear &"','"& Request.Form("Date_40")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=40" 
	con.executeQuery (sqlIns)  
case "41"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 41,'"& currYear &"','"& Request.Form("Date_41")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=41" 
	con.executeQuery (sqlIns)  
    sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 42,'"& currYear &"','"& DateAdd("d",1,Request.Form("Date_41"))  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=42" 
	con.executeQuery (sqlIns)
	case "43"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 43,'"& currYear &"','"& Request.Form("Date_43")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=43" 
	con.executeQuery (sqlIns)  
case "44"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 44,'"& currYear &"','"& Request.Form("Date_44")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=44" 
	con.executeQuery (sqlIns)  
case "45"
	sqlIns="SET DATEFORMAT DMY;Insert into DaysInYear (DayId,Year,DayDate,DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate)  Select 45,'"& currYear &"','"& Request.Form("Date_45")  &"',DayName,DayType,DayTypeName,DayTypeColor,DayOrder,DayCalculate from DaysTemplate where Id=45" 
	con.executeQuery (sqlIns)  

	end select
	next
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
											<td align="left"><%if request.Form ("add")="1" then%><a href="holidays.asp?deleteId=<%=currYear%>" ONCLICK="return checkDelete()"><input type="button" value="איפוס" class="but_menu" onclick="javascript:DeleteData(<%=currentYear%>)" style="width:90"  id="Button2" name="Button2"></a><%else%><input type="submit" value="שמור נתונים" class="but_menu" style="width:90" onclick="return checkForm()"
													id="Button1" name="Button1"><%end if%></td>
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
											<td align="center">אופן חישוב</td>
											<td align="center">תאריך קלנדרי לפי שנת
												<%=currYear%>
											</td>
											<td align="center">הגדרת העבודה ביום זה</td>
											<td align="center">שם מועד/ חג
											</td>
										</tr>
										<%sqlstr = "Select * from DaysInYear Where Year = " & currYear 				
	set rsdays = con.getRecordSet(sqlstr)		
	If not rsdays.eof Then
		do while not rsdays.eof%>
										<tr>
											<td align="center" style="color:#001BC0"><%if IsNumeric(rsdays("DayCalculate")) then%>+<%=rsdays("DayCalculate")%><%end if%></td>
											<td align="center"><table border="0" cellpadding="0" cellspacing="0" ID="Table4">
													<tr>
														<td align="right" dir="rtl"><%if rsdays("DayAction")="1" then%><a href="" onclick="cal1xx.select(document.getElementById('Date_<%=rsdays("Id")%>'),'AsDate_<%=rsdays("Id")%>','dd/MM/yyyy'); return false;" id="A1">
																<img src="../../images/calendar.gif" border="0" align="center"></a> <input type="text" name="Date_<%=rsdays("Id")%>" id="Text1" size="10" value="<%=rsdays("DayDate")%>" readonly dir="rtl" class=Myinput>
															<%else%>
															<%=rsdays("DayDate")%>
															<%end if%>
														</td>
													</tr>
												</table>
											</td>
											<td align=right style="background-color:<%=rsdays("DayTypeColor")%>"><%=rsdays("DayTypeName")%></td>
											<td align="right" style="background-color:#e1e1e1" dir="rtl"><%=rsdays("DayName")%></td>
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

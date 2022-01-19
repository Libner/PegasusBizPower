<%@ Page Language="vb" AutoEventWireup="false" Codebehind="calendar.aspx.vb" Inherits="bizpower_pegasus.calendar"%>
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
  </head>
  <body style="margin:0px">

    <form id="Form1" method="post" runat="server">
    <table border=0 cellpadding=0 cellspacing=0 width=100% align=left>
  
   <tr>
<td width=100% style="color:#000000;background-color:#e1e1e1e1;"> 
<table border=0 cellpadding=1 cellpadding=1 width=100% style="color:#000000;background-color:#E6E6E6;">
<tr><td align=center colspan=13><span style="font-size:14pt"><select id="y" name="y" onchange="this.form.submit()">
<%for yy=2014 to 2040%>
<option value="<%=yy%>" <%if currYear=yy then%>selected<%end if%> ><%=yy%></option>
<%next%>
</select> </span></td>   </tr>
<tr>
<%For i=12 to 1 step-1%>
<th><%=MonthName(i)%></th>
<%next%>
<th>יום</th>
</tr>
<tr><td colspan=13 height=5>&nbsp;</td></tr>
<%for j=1  to 31%>
<tr>
<%For i=12 to 1 step-1
dayT=CStr(j) & "/" & CStr(i) & "/" & CStr(currYear)
%>
<%if IsDate(dayT) then%>
<td align=center <%if Weekday(dayT)=6 then%> style="height:100%;background-color:#ff0000;color:#ffffff"<%end if%>
<%if Weekday(dayT)=5 then%> style="height:100%;background-color:#FFC0CB;color:#000000"<%end if%>
<%if Weekday(dayT)<5 or  Weekday(dayT)=7 then%>style="height:100%;background-color:#7CFC00;color:#000000"<%end if%>
>
<table border=0 cellpadding=0 cellspacing=0 width=100% height=100%>
<tr><td align=center>
<%if Weekday(dayT)=6 then%> אי עבודה<%'=WeekDayName(Weekday(dayT))%><%end if%>
<%if Weekday(dayT)=5 then%> חצי יום עבודה<%'=WeekDayName(Weekday(dayT))%><%end if%>
<%if Weekday(dayT)<5 or  Weekday(dayT)=7 then%>יום עבודה מלא
<%end if%>
<%'=WeekDayName(Weekday(dayT))%>
</td></tr>
</table>
</td>


<%else%>

<td align=center style=background-color:#ffffff></td>

<%end if%>
<%next%>
<th align=center style=background-color:#e1e1e1><%=j%></th>
</tr>
<%next%>
</table></td>
</tr>
<tr><td colspan=13>

</td></tr>

</table>

    </form>

  </body>
</html>

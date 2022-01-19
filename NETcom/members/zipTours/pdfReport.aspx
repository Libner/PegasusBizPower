<%@ Page Language="vb" AutoEventWireup="false" Codebehind="pdfReport.aspx.vb" Inherits="bizpower_pegasus.pdfReport"%>

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
  </head>
  <body>

    <form id="Form1" method="post" runat="server">
<table border=0 cellpadding=0 cellspacing=0 width=100%>
<tr ><td>
<table border=0 cellpadding=0 cellspacing=0 width=100%>
<tr >
<td height=30></td></tr></table>
</td></tr>
<tr><td>
<table border=0 cellpadding=0 cellspacing=0 width=100%>

<tr><td colspan=17 style="width:100%;height:2px;background-color:#C9C9C9"></td></tr>
<tr >
<td  class="title_sort">&nbsp;</td>

<asp:Repeater ID=rptTitle runat=server>
<ItemTemplate>
<td class="title_sort" align=center><%#Container.DataItem("Column_Name")%></td>
</ItemTemplate>
<FooterTemplate><td class="title_sort" align=center>מופיע במסך צפיה</td></FooterTemplate>
</asp:Repeater>

<asp:Repeater ID=rptCustomers Runat=server>
<ItemTemplate>
<tr bgcolor="#C9C9C9" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';" style="background-color: rgb(201, 201, 201);">
<td>&nbsp;</td>
<td align=center class=price><%#Container.Dataitem("Price")%></td>
<td align=center class="FlyingCompany"><%#Container.DataItem("Flying_Company")%></td>
<td align=center class="UserName"><%#Container.DataItem("User_Name")%></td>
<td align=center><%#func.dbNullFix(Container.DataItem("CountlastMonth"))%></td>
<td align=center><%#func.dbNullFix(Container.DataItem("Countlast2Week"))%></td>
<td align=center><%#func.dbNullFix(Container.DataItem("CountlastWeek"))%></td>
<td align=center nowrap><%#func.dbNullFix(Container.DataItem("LastDateSale"))%></td>
<td align=center><%#func.dbNullFix(Container.DataItem("CountMembersBitulim"))%></td>
<td align=center><%#Container.DataItem("Dep_Name")%></td>
<td align=center><%#func.dbNullFix(Container.DataItem("CountMembers"))%></td>
<td align=center><%#Container.Dataitem("CountryName")%></td>
<td align=center><%#Container.Dataitem("Date_Begin")%></td>
<td align=center><%#Container.DataItem("Series_Name")%></td>
<%'IF ViewColumn(1)=true then%>
<td align=right style="padding-top:1px;padding-bottom:1px">
<table border=0 style="background-color:<%#Container.DataItem("Status_Color")%>;width:100%" cellpadding=1 cellspacing=1 align=right>
<tr><td align=right><span style="color:#000000">&nbsp;<%#Container.DataItem("Status_Name")%>&nbsp;</span></td>
<td align=right>
</td>
</tr></table></td><%'end if%>

<td align=center><input type="checkbox" ID="chkTour" Runat="server" NAME="chkTour"></td>

<%if false then%>
	<td class="td_admin_4" align="center"><a class="table_fa_icon" href="" onclick="return moveTo('<%#Container.DataItem("Departure_Id")%>')"><img src="../../images/arrow_top_bot.gif" align="top" border=0 alt="העבר למקום אחר"></a></td>
    <td class="td_admin_4" align="center"><input type="text" value="" onKeyPress="return getNumbers(this)" class=Form style="width:30" ID="inpOrder<%#Container.DataItem("Departure_Id")%>"></td>		    			
	<td class="td_admin_4" align="center"><font color="#060165"><B>&nbsp;<%#Container.ItemIndex+1%>&nbsp;</b></font></td>
<%end if%>
</tr>
</ItemTemplate>
<AlternatingItemTemplate>
<tr bgcolor="#E6E6E6" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#E6E6E6';" style="background-color: rgb(230, 230, 230);">
<td>&nbsp;</td>
<td align=center><%#Container.Dataitem("Price")%></td>
<td align=center><%#Container.DataItem("Flying_Company")%></td>
<td align=center><%#Container.DataItem("User_Name")%></td>
<td align=center><%#func.dbNullFix(Container.DataItem("CountlastMonth"))%></td>
<td align=center><%#func.dbNullFix(Container.DataItem("Countlast2Week"))%></td>
<td align=center><%#func.dbNullFix(Container.DataItem("CountlastWeek"))%></td>
<td align=center><%#func.dbNullFix(Container.DataItem("LastDateSale"))%></td>
<td align=center><%#func.dbNullFix(Container.DataItem("CountMembersBitulim"))%></td>
<td align=center><%#Container.DataItem("Dep_Name")%></td>
<td align=center><%#func.dbNullFix(Container.DataItem("CountMembers"))%></td>
<td align=center><%#Container.Dataitem("CountryName")%></td>
<td align=center><%#Container.Dataitem("Date_Begin")%></td>
<td align=center><%#Container.DataItem("Series_Name")%></td>
<%'IF ViewColumn(1)=true then%>
<td align=right style="padding-top:1px;padding-bottom:1px">
<table border=0 style="background-color:<%#Container.DataItem("Status_Color")%>;width:100%" cellpadding=1 cellspacing=1 align=right>
<tr><td align=right><span style="color:#000000">&nbsp;<%#Container.DataItem("Status_Name")%>&nbsp;</span></td>
<td align=right>
</td>
</tr></table></td><%'end if%>
<td align=center><input type="checkbox" ID="chkTour" Runat="server" NAME="chkTour"></td>


</tr>
</AlternatingItemTemplate>
</asp:Repeater>

 <!--  <div id="dvCustomers">-->
 	<!--paging-->
			
</table>

</td>

	</tr></table>
</form>

  </body>
</html>

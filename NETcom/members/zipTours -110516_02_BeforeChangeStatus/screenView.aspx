<%@ Page Language="vb" AutoEventWireup="false" Codebehind="screenView.aspx.vb" Inherits="bizpower_pegasus.screenView" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>screenView</title>  
       <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>

		<script src="../../include/CalendarPopup.js" charset="windows-1255"></script>

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
<tr bgcolor=#d8d8d8 style="height:25px" >

	<td class="td_admin_5" align="center" width="50" ><asp:LinkButton Runat="server" ID="btnSearchAll" Width="100%" CssClass="button_small1">הצג הכל</asp:LinkButton></td>
									<td class="td_admin_5" align="center"><asp:LinkButton Runat="server" ID="btnSearch" CssClass="button_small1">חפש</asp:LinkButton></td>
<%IF ViewColumn(13)=true then%><td></td><%end if%>
<%IF ViewColumn(12)=true then%><td></td><%end if%>
<%IF ViewColumn(11)=true then%><td align=center><select runat="server" id="sUsers" class="searchList" dir="rtl" name="sUsers"></select></td><%end if%>
<%IF ViewColumn(6)=true then%><td></td><%end if%>
<%IF ViewColumn(7)=true then%><td></td><%end if%>
<%IF ViewColumn(8)=true then%><td></td><%end if%>
<%IF ViewColumn(9)=true then%><td></td><%end if%>
<%IF ViewColumn(10)=true then%><td></td><%end if%>

<%IF ViewColumn(5)=true then%><td align=center><select runat="server" id="sDepartments" class="searchList" dir="rtl" name="sDepartments"></select></td><%end if%>
<%IF ViewColumn(4)=true then%><td align=center></td><%end if%>

<%IF ViewColumn(3)=true then%>
	<td class="td_admin_5" align="center" nowrap>
			<a href="" onclick="cal1xx.select(document.getElementById('sPayFromDate'),'AsPayFromDate','dd/MM/yyyy'); return false;"
											id="AsPayFromDate"><img src="../../images/calendar.gif" border="0" align="center"></a>
										<input runat="server" type="text" id="sPayFromDate" class="searchList" style="WIDTH:70px"
											NAME="sPayFromDate"> מי  <br>
	<a href="" onclick="cal1xx.select(document.getElementById('sPayToDate'),'AsPayToDate','dd/MM/yyyy'); return false;"
											id="AsPayToDate"><img src="../../images/calendar.gif" border="0" align="center"></a>
										<input runat="server" type="text" id="sPayToDate" class="searchList" style="WIDTH:70px"
											NAME="sPayToDate"> עד 
							</td><%end if%>

<%IF ViewColumn(2)=true then%><td class="td_admin_5"><select runat="server" id="sSeries" class="searchList" dir="rtl" name="sSeries" ></select></td><%end if%>

<%IF ViewColumn(1)=true then%><td align=center class="td_admin_5"><select runat="server" id="sStatus" style ="font-size:12px;width:100px" dir="rtl" name="sStatus"></select></td><%end if%>
<td class="td_admin_5"></td>
</tr>
<tr><td colspan=14 style="width:100%;height:2px;background-color:#C9C9C9"></td></tr>
<tr >
<td  class="title_sort">&nbsp;</td><td  class="title_sort">&nbsp;</td>

<asp:Repeater ID=rptTitle runat=server>
<ItemTemplate>
<td class="title_sort" align=center><%#Container.DataItem("Column_Name")%></td>

</ItemTemplate>

</asp:Repeater>

<asp:Repeater ID=rptCustomers Runat=server>
<ItemTemplate>
<tr bgcolor="#C9C9C9" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';" style="background-color: rgb(201, 201, 201);height:30px;">
<td>&nbsp;</td><td>&nbsp;</td>
<%IF ViewColumn(13)=true then%><td align=center class=price><%#Container.Dataitem("Price")%></td><%end if%>
<%IF ViewColumn(12)=true then%><td align=center class="FlyingCompany"><%#Container.DataItem("Flying_Company")%></td><%end if%>
<%IF ViewColumn(11)=true then%><td align=center class="UserName"><%#Container.DataItem("User_Name")%></td><%end if%>
<%IF ViewColumn(6)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountlastMonth"))%></td><%end if%>
<%IF ViewColumn(7)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("Countlast2Week"))%></td><%end if%>
<%IF ViewColumn(8)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountlastWeek"))%><%'#func.GetCountLastDateSale(Container.DataItem("Departure_Code"))%></td><%end if%>
<%IF ViewColumn(9)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("LastDateSale"))%><%'#func.GetLastDateSale(Container.DataItem("Departure_Code"))%></td><%end if%>
<%IF ViewColumn(10)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountMembersBitulim"))%><%'#func.GetCountMembersBitulim(Container.DataItem("Departure_Code"))%></td><%end if%>

<%IF ViewColumn(5)=true then%><td align=center><%#Container.DataItem("Dep_Name")%></td><%end if%>
<%IF ViewColumn(4)=true then%><td align=center><%#Container.DataItem("CountMembers")%><%'#func.GetCountMembers(Container.DataItem("Departure_Code"))%></td><%end if%>
<%IF ViewColumn(3)=true then%><td align=center><%#Container.Dataitem("Date_Begin")%></td><%end if%>
<%#IIF(ViewColumn(2),"<td align=center>" & Container.DataItem("Series_Name") &"</td>","")%>
<%IF ViewColumn(1)=true then%>
<td align=right align=center style="padding-top:1px;padding-bottom:1px">
<table border=0 style="background-color:<%#Container.DataItem("Status_Color")%>;width:100%;height:30px;" cellpadding=1 cellspacing=1 align=right>
<tr><td align=right><span style="color:#000000">&nbsp;<%#Container.DataItem("Status_Name")%>&nbsp;</span></td>

</tr></table></td><%end if%>

</tr>
</ItemTemplate>
<AlternatingItemTemplate>
<tr bgcolor="#E6E6E6" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#E6E6E6';" style="background-color: rgb(230, 230, 230);height:30px;">
<td>&nbsp;</td><td>&nbsp;</td>

<%IF ViewColumn(13)=true then%><td align=center><%#Container.Dataitem("Price")%></td><%end if%>
<%IF ViewColumn(12)=true then%><td align=center><%#Container.DataItem("Flying_Company")%></td><%end if%>
<%IF ViewColumn(11)=true then%><td align=center><%#Container.DataItem("User_Name")%></td><%end if%>
<%IF ViewColumn(6)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountlastMonth"))%></td><%end if%>
<%IF ViewColumn(7)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("Countlast2Week"))%></td><%end if%>
<%IF ViewColumn(8)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountlastWeek"))%><%'#func.GetCountLastDateSale(Container.DataItem("Departure_Code"))%></td><%end if%>
<%IF ViewColumn(9)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("LastDateSale"))%></td><%end if%>
<%IF ViewColumn(10)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountMembersBitulim"))%><%'#func.GetCountMembersBitulim(Container.DataItem("Departure_Code"))%></td><%end if%>


<%IF ViewColumn(5)=true then%><td align=center><%#Container.DataItem("Dep_Name")%></td><%end if%>
<%IF ViewColumn(4)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountMembers"))%><%'#func.GetCountMembers(Container.DataItem("Departure_Code"))%></td><%end if%>
<%IF ViewColumn(3)=true then%><td align=center><%#Container.Dataitem("Date_Begin")%></td><%end if%>
<%#IIF(ViewColumn(2),"<td align=center>" & Container.DataItem("Series_Name") &"</td>","")%>
<%IF ViewColumn(1)=true then%>
<td align=right style="padding-top:1px;padding-bottom:1px;">
<table border=0 style="background-color:<%#Container.DataItem("Status_Color")%>;width:100%;height:30px;" cellpadding=1 cellspacing=1 align=right>
<tr><td align=right><span style="color:#000000">&nbsp;<%#Container.DataItem("Status_Name")%>&nbsp;</span></td>
</tr></table></td><%end if%>

</tr>
</AlternatingItemTemplate>
</asp:Repeater>

 <!--  <div id="dvCustomers">-->
 	<!--paging-->
						<asp:PlaceHolder id="pnlPages" Runat="server">
						<tr><td height="2" colspan="13"></td></tr>
						<tr>
							<td class="plata_paging" vAlign="top" align="center" colspan="18" bgcolor=#D8D8D8>
								<table dir="ltr" cellSpacing="0" cellPadding="2" width="100%" border="0" >
									<tr>
										<td class="plata_paging" vAlign="baseline" noWrap align="left" width="150">&nbsp;הצג
										<asp:DropDownList id="PageSize" CssClass="PageSelect" Runat="server" AutoPostBack="true">
										<asp:ListItem Value="10">10</asp:ListItem>
										<asp:ListItem Value="20" >20</asp:ListItem>
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

</td>
 </td></tr>
	
	
	
		
	<tr><td height=10 nowrap></td></tr>
	</table>
	</td>
	</tr></table>
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

<%@ Page Language="vb" AutoEventWireup="false" Codebehind="screenSetting.aspx.vb" Inherits="bizpower_pegasus.screenSetting" %>

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

<script language=javascript>
function checkMoveTours()
{

alert("111")
}
function checkDisable(depId)
{
//alert(depId)
if (depId != '' ) {
$.ajax({
  type: "POST",
  url:"DepartureChk.aspx",
  data: depId,
  data: {depId: depId},
success: function(msg) {
//alert(msg);


	

	},
	async:   false
	
});
}
//alert("tftrr")

}
function openStatus(departureId)
{
	newWin=window.open("editStatus.aspx?DepartureId=" + escape(departureId), "newWin", "toolbar=0,menubar=0,width=300,height=100,top=100,left=5,scrollbars=auto");
	newWin.opener=window
}

</script>


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
			<a href="" onclick="cal1xx.select(document.Form1.sPayFromDate,'AsPayFromDate','dd/MM/yyyy'); return false;"
											id="AsPayFromDate"><img src="../../images/calendar.gif" border="0" align="center"></a>
										<input runat="server" type="text" id="sPayFromDate" class="searchList" style="WIDTH:70px"
											NAME="sPayFromDate"> מי  <br>
	<a href="" onclick="cal1xx.select(document.Form1.sPayToDate,'AsPayToDate','dd/MM/yyyy'); return false;"
											id="AsPayToDate"><img src="../../images/calendar.gif" border="0" align="center"></a>
										<input runat="server" type="text" id="sPayToDate" class="searchList" style="WIDTH:70px"
											NAME="sPayToDate"> עד 
							</td><%end if%>

<%IF ViewColumn(2)=true then%><td class="td_admin_5"><select runat="server" id="sSeries" class="searchList" dir="rtl" name="sSeries" ></select></td><%end if%>

<%IF ViewColumn(1)=true then%><td align=center class="td_admin_5"><select runat="server" id="sStatus" class="searchList" dir="rtl" name="sStatus"></select></td><%end if%>
<td class="td_admin_5"></td>
</tr>
<tr><td colspan=16 style="width:100%;height:2px;background-color:#C9C9C9"></td></tr>
<tr >
<td  class="title_sort">&nbsp;</td><td  class="title_sort">&nbsp;</td>

<asp:Repeater ID=rptTitle runat=server>
<ItemTemplate>
<td class="title_sort" align=center><%#Container.DataItem("Column_Name")%></td>

</ItemTemplate>
<FooterTemplate><td class="title_sort" align=center>מופיע <BR>במסך צפיה</td>
	<%if false then%><td align="center"  class="title_sort" colspan="3" nowrap>סדר</td>><%end if%>
</FooterTemplate>
</asp:Repeater>

<asp:Repeater ID=rptCustomers Runat=server>
<ItemTemplate>
<tr bgcolor="#C9C9C9" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';" style="background-color: rgb(201, 201, 201);">
<td>&nbsp;</td><td>&nbsp;</td>
<%IF ViewColumn(13)=true then%><td align=center class=price><%#Container.Dataitem("Price")%></td><%end if%>
<%IF ViewColumn(12)=true then%><td align=center class="FlyingCompany"><%#Container.DataItem("Flying_Company")%></td><%end if%>
<%IF ViewColumn(11)=true then%><td align=center class="UserName"><%#Container.DataItem("User_Name")%></td><%end if%>
<%IF ViewColumn(6)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("lastMonth"))%></td><%end if%>
<%IF ViewColumn(7)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("last2Week"))%></td><%end if%>
<%IF ViewColumn(8)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("lastWeek"))%><%'#func.GetCountLastDateSale(Container.DataItem("Departure_Code"))%></td><%end if%>
<%IF ViewColumn(9)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("LastDateSale"))%><%'#func.GetLastDateSale(Container.DataItem("Departure_Code"))%></td><%end if%>
<%IF ViewColumn(10)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountMembersBitulim"))%><%'#func.GetCountMembersBitulim(Container.DataItem("Departure_Code"))%></td><%end if%>

<%IF ViewColumn(5)=true then%><td align=center><%#Container.DataItem("Dep_Name")%></td><%end if%>
<%IF ViewColumn(4)=true then%><td align=center><%#func.GetCountMembers(Container.DataItem("Departure_Code"))%></td><%end if%>
<%IF ViewColumn(3)=true then%><td align=center><%#Container.Dataitem("Date_Begin")%></td><%end if%>
<%#IIF(ViewColumn(2),"<td align=center>" & Container.DataItem("Series_Name") &"</td>","")%>
<%IF ViewColumn(1)=true then%>
<td align=right align=center style="padding-top:1px;padding-bottom:1px">
<table border=0 style="background-color:<%#Container.DataItem("Status_Color")%>;width:100%" cellpadding=1 cellspacing=1 align=right>
<tr><td align=right><span style="color:#000000">&nbsp;<%#Container.DataItem("Status_Name")%>&nbsp;</span></td>
<td align=right><a href="#" class="link_categ" onclick="openStatus('<%#Container.DataItem("Departure_Id")%>');return false" title="עדכון סטטוס" style="cursor:hand"><img src="../../images/edit_icon.gif" border="0" style="margin-left:1px;"></a></td>
</tr></table></td><%end if%>
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
<td>&nbsp;</td><td>&nbsp;</td>

<%IF ViewColumn(13)=true then%><td align=center><%#Container.Dataitem("Price")%></td><%end if%>
<%IF ViewColumn(12)=true then%><td align=center><%#Container.DataItem("Flying_Company")%></td><%end if%>
<%IF ViewColumn(11)=true then%><td align=center><%#Container.DataItem("User_Name")%></td><%end if%>
<%IF ViewColumn(6)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("lastMonth"))%></td><%end if%>
<%IF ViewColumn(7)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("last2Week"))%></td><%end if%>
<%IF ViewColumn(8)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("lastWeek"))%><%'#func.GetCountLastDateSale(Container.DataItem("Departure_Code"))%></td><%end if%>
<%IF ViewColumn(9)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("LastDateSale"))%></td><%end if%>
<%IF ViewColumn(10)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountMembersBitulim"))%><%'#func.GetCountMembersBitulim(Container.DataItem("Departure_Code"))%></td><%end if%>


<%IF ViewColumn(5)=true then%><td align=center><%#Container.DataItem("Dep_Name")%></td><%end if%>
<%IF ViewColumn(4)=true then%><td align=center><%#func.GetCountMembers(Container.DataItem("Departure_Code"))%></td><%end if%>
<%IF ViewColumn(3)=true then%><td align=center><%#Container.Dataitem("Date_Begin")%></td><%end if%>
<%#IIF(ViewColumn(2),"<td align=center>" & Container.DataItem("Series_Name") &"</td>","")%>
<%IF ViewColumn(1)=true then%>
<td align=right style="padding-top:1px;padding-bottom:1px">
<table border=0 style="background-color:<%#Container.DataItem("Status_Color")%>;width:100%" cellpadding=1 cellspacing=1 align=right>
<tr><td align=right><span style="color:#000000">&nbsp;<%#Container.DataItem("Status_Name")%>&nbsp;</span></td>
<td align=right><a href="#" class="link_categ" onclick="openStatus('<%#Container.DataItem("Departure_Id")%>');return false" title="עדכון סטטוס" style="cursor:hand"><img src="../../images/edit_icon.gif" border="0" style="margin-left:1px;"></a></td>
</tr></table></td><%end if%>
<td align=center> <input type="checkbox" ID="chkTour" Runat="server"  NAME="chkTour" ></td>
<%if false then%>
	<td class="td_admin_4" align="center"><a class="table_fa_icon" href="" onclick="return moveTo('<%#Container.DataItem("Departure_Id")%>')"><img src="../../images/arrow_top_bot.gif" align="top" border=0 alt="העבר למקום אחר"></a></td>
    <td class="td_admin_4" align="center"><input type="text" value="" onKeyPress="return getNumbers(this)" class=Form style="width:30" ID="inpOrder<%#Container.DataItem("Departure_Id")%>"></td>		    			
	<td class="td_admin_4" align="center"><font color="#060165"><B>&nbsp;<%#Container.ItemIndex+1%>&nbsp;</b></font></td>
<%end if%>
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
												<asp:ListItem Value="20" Selected="True">20</asp:ListItem>
								
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
											<asp:Label CssClass="plata_paging" id="lblCount" Runat="server"></asp:Label>&nbsp;לקוחות
										</td>
										</tr>				
									</table>
								</td>
							</tr>
						</asp:PlaceHolder>
</table>

</td>
  <td width=100 nowrap valign=top class="td_menu" style="border: 1px solid #808080" dir="rtl">
    <table cellpadding=1 cellspacing=0 width="100%" border="0" dir="rtl" ID="Table3">
    <tr><td align="right" colspan="2" class="title_search">&nbsp;</td></tr>
	<tr><td colspan="2" height=10 nowrap>
	<asp:button  ID=btnSaveChecked EnableViewState="false" runat="server" Cssclass="button_edit_1" Text="עדכן מסך צפייה"></asp:Button> </td></tr>
	
	<%if false then%><a class="button_edit_1" style="width:94;" href='javascript:void(0)' onclick="return checkMoveTours();">עדכן מסך צפייה<%end if%>
	
		
	<tr><td colspan="2" height=10 nowrap></td></tr>
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

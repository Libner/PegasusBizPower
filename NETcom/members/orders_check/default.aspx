<%@ Page Language="vb" AutoEventWireup="false" Codebehind="default.aspx.vb" validateRequest="false" Inherits="bizpower_pegasus2018.orders_check" %>
<%@ Register TagPrefix="DS" TagName="TopIn" Src="../../top_in.ascx" %>
<%@ Register TagPrefix="DS" TagName="LogoTop" Src="../../logo_top.ascx" %>
<%@ Register TagPrefix="DS" TagName="metaInc" Src="../../title_meta_inc.ascx" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<head>
		<DS:metaInc id="metaInc" runat="server"></DS:metaInc>
	
		<meta charset="windows-1255">
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script src="../../include/CalendarPopup.js" charset="windows-1255"></script>
		<script type="text/javascript" language="javascript">
		<!-- 
		function openExcel()
		{
			window.open("upload.aspx?table=gilboa","ExcelReport","scrollbars=yes,menubar=yes,width=420,height=200,left=100,top=100");
		}	
		
		function openReport()
		{
			var dateStart = document.getElementById("start_date").value;
			var dateEnd = document.getElementById("end_date").value;
			var UserId = document.getElementById("cmbUser").value;
			window.open("report_orders.aspx?dateStart=" + dateStart + "&dateEnd=" + dateEnd + "&UserId=" + UserId,"ReportExcel","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");
		}
		
		function openGReport()
		{
			var dateStart = document.getElementById("start_date").value;
			var dateEnd = document.getElementById("end_date").value;
			var UserId = document.getElementById("cmbUser").value;
			window.open("report_graph.aspx?dateStart=" + dateStart + "&dateEnd=" + dateEnd + "&UserId=" + UserId,"ReportExcel","scrollbars=yes,menubar=yes, toolbar=yes, width=980,height=600,left=10,top=10");
		}
		
		function DoCal(elTarget)
		{
		if (!window.showModalDialog) {
		//alert("ffff")

		}
			if (showModalDialog){
				var sRtn;
				sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
				if (sRtn!="")
				elTarget.value = sRtn;
			}else
			alert("Internet Explorer 4.0 or later is required.");
			return false;
			window.document.focus;   
		}	
		
		function doSearch()
		{
			var srchPFile = document.getElementById("srchPFile").value;
			var srchService = document.getElementById("srchService").value;
			document.location.href = "default.aspx?srchPFile=" + escape(srchPFile) + "&srchService=" + escape(srchService);
			return false;
		}		
		
		function entsubsearch(e) { 
			if( typeof( e ) == "undefined" && typeof( window.event ) != "undefined" )
				e = window.event;
	
			if (e && e.keyCode == 13)
				return doSearch();		
			else
				return true; }			
        // -->
		</script>			
</head>
<body>
<form id="Form1"  method="post" runat="server" autocomplete=off>	
	<DS:LOGOTOP id="logotop"  runat="server"></DS:LOGOTOP>
	<DS:TOPIN id="topin" numOfLink="6"  numOfTab="4" toplevel2=37 runat="server"></DS:TOPIN>
	<table cellSpacing="0" cellPadding="0" width="100%">
								
    <tr><td width="100%">
	<table cellpadding="1" cellspacing="1" width="100%" border="0" dir="<%=dir_var%>">
	<tr>
		<td width="100%" bgcolor="#E6E6E6" valign="top">
	   <table cellspacing="1" cellpadding="2" border="0" style="background-color:White;border-width:0px;width:100%;">
		<tr align="Right">
			<td class="title_sort"><input type="button" id="btnShowAll" value="הצג הכל" onclick="document.location.href='default.aspx';" 
			class="but_map">&nbsp;&nbsp;&nbsp;<input type="button" id="btnSearch" value="חפש" onclick="return doSearch();" class="but_site"></td>
			<td class="title_sort" dir="rtl">מספר דוקט&nbsp;<input dir="ltr"  title="מספר דוקט" type="text" id="srchPFile" value="<%=func.vFix(srchPFile)%>" onkeypress="return entsubsearch(event)" /></td>
			<td class="title_sort" dir="rtl">קוד טיול&nbsp;<input dir="ltr" title="קוד טיול" type="text" id="srchService" value="<%=func.vFix(srchService)%>" onkeypress="return entsubsearch(event)" /></td>
			<td class="title_sort" width="80" nowrap><b>חפש לפי</b></td>
		</tr></table>
		<asp:DataGrid ID="dtgGilboa" Runat="server" AutoGenerateColumns=False DataKeyField="Service Name"
		CellPadding=2 CellSpacing=1 BorderWidth=0 ShowFooter=False ShowHeader="True" AllowPaging="True" 
		PageSize="50" PagerStyle-Mode="NumericPages" OnPageIndexChanged="dtgGilboa_PageChanger"
		Width="100%" BackColor="#FFFFFF" AllowSorting="True" >
		<HeaderStyle CssClass="title_sort" HorizontalAlign="Right" ></HeaderStyle>
		<ItemStyle CssClass="card" HorizontalAlign="Right" ></ItemStyle>		
		<Columns>
		<asp:BoundColumn DataField="Pax Num" HeaderText="מספר המשויכים לדוקט" HeaderStyle-CssClass="title_sort" ></asp:BoundColumn>
		<asp:BoundColumn SortExpression="PFile" DataField="PFile" HeaderText="מספר דוקט" HeaderStyle-CssClass="title_sort"></asp:BoundColumn>
		<asp:BoundColumn SortExpression="Service Name" DataField="Service Name" HeaderText="קוד טיול" HeaderStyle-CssClass="title_sort"></asp:BoundColumn>
		</Columns>
		<PagerStyle HorizontalAlign="Center" BackColor="#E6E6E6" CssClass="forms" PageButtonCount=10></PagerStyle> 
		</asp:DataGrid>
		</td>
		<td width="110" nowrap align="<%=align_var%>" valign="top" style="border:1px solid #808080;" class="td_menu">
		<table cellpadding="1" cellspacing="1" width="100%" border="0">	
		<tr><td colspan="2" height="10"></td></tr>
		<tr> 
		<td align="right" width="120" class="td_title_2" nowrap><input type="button" onclick="javascript:openExcel()" 
		class="button_edit_1" value="xls ייבוא מקובץ"></td>
		</tr>				
		<tr><td colspan="2" height="10"></td></tr>
		<tr><td align="<%=align_var%>" colspan="2"><b>מתאריך</b>&nbsp;</td></tr>
		<tr>
			<td align="center" colspan="2" nowrap>
			<a href="" onclick="cal1xx.select(document.getElementById('start_date'),'Asstart_date','dd/MM/yyyy'); return false;"
											id="Asstart_date"><img src="../../images/calendar.gif" border="0" align="center"></a>
			<input dir="ltr" class="texts" type="text" id="start_date" name="start_date" value="" 
			style="width: 70px"		 readonly></td>					
		</tr>
		<tr><td width=100% align="<%=align_var%>" colspan=2>&nbsp;<b>עד תאריך</b>&nbsp;</td></tr>
		<tr>
			<td align="center" colspan="2" nowrap>
				<a href="" onclick="cal1xx.select(document.getElementById('end_date'),'Asend_date','dd/MM/yyyy'); return false;"
											id="Asend_date"><img src="../../images/calendar.gif" border="0" align="center"></a>
			<input dir="ltr" class="texts" type="text" id="end_date" name="end_date" value="" 			style="width: 70px" readonly></td>		
		</tr>
		<tr>
		<td align="<%=align_var%>" colspan="2"><select name="cmbUser" dir="rtl" class="norm" runat="server"
		style="width:100%" ID="cmbUser"></select></td>
		</tr>		
		<tr> 
		<td align="right" width="120" class="td_title_2" nowrap><a href="javascript:openReport()" class="button_edit_1">דוח מנהל</a></td>
		</tr>	
		<tr> 
		<td align="right" width="120" class="td_title_2" nowrap><a href="javascript:openGReport()" class="button_edit_1">דוח גרפי</a></td>
		</tr>			
		</table></td></tr>
		</table></td></tr>			
	</table>
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
</HTML>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="default.aspx.vb" validateRequest="false" Inherits="bizpower_pegasus2018.logs_report" %>
<%@ Register TagPrefix="DS" TagName="TopIn" Src="../../top_in.ascx" %>
<%@ Register TagPrefix="DS" TagName="LogoTop" Src="../../logo_top.ascx" %>
<%@ Register TagPrefix="DS" TagName="metaInc" Src="../../title_meta_inc.ascx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<head>
	<DS:metaInc id="metaInc" runat="server"></DS:metaInc>
	
	
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<script language="javascript" type="text/javascript" src="../CalendarPopup.js"></script>

		<script type="text/javascript" language="javascript">
		<!-- 
		function callCalendar(pf,pid)
	{
		cal1xx.select(pf,pid,'dd/MM/yyyy')
	}

		function DoCal(elTarget)
		{
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
        // -->
		</script>			
</head>
<body>
<form id="Form1" name="Form1" method="post" runat="server" autocomplete=off>	
	<DS:LOGOTOP id="logotop"  runat="server"></DS:LOGOTOP>
	<DS:TOPIN id="topin"  numOfLink="0"  numOfTab="4" toplevel2="38" runat="server"></DS:TOPIN>
	<table cellSpacing="0" cellPadding="0" width="100%">
					
    <tr><td width="100%" align="center">
	<table cellpadding="1" cellspacing="1" width="400" border="0" dir="<%=dir_var%>" align="center" >
	<tr><td height="20" nowrap></td></tr>
	<tr><td align=center nowrap><span style="font-size:20px;">דוח לוגים על ביצוע פעולות כלליות במערכת</span></td></tr>
    <tr><td><table bgcolor="#CCCCCC" cellpadding="2" cellspacing="1" width="100%" border="0">
    <tr bgcolor="#F3F3F3">
    <td width="295" nowrap align="right" style="padding-right: 10px;">
    						<a href='' onclick='callCalendar(document.getElementById("Form1").date_start,"AsdateStart");return false;' id='AsdateStart'><image src='../../images/calend.gif' border=0  valign=bottom></a>&nbsp;<input dir='ltr' type='text' class="form_text" size=10 id="date_start" name='date_start'  value='<%=date_start%>' maxlength=10 readonly>
    </td>
    <td width="100" nowrap align="right" class="forms" bgcolor="#E3ECF2" style="padding-right: 10px;">מתאריך</td>
    </tr>
    <tr bgcolor="#F3F3F3">
    <td width="295" nowrap align="right" style="padding-right: 10px;">
        						<a href='' onclick='callCalendar(document.getElementById("Form1").date_end,"Asdate_end");return false;' id='Asdate_end'><image src='../../images/calend.gif' border=0  valign=bottom></a>&nbsp;<input dir='ltr' type='text' class="form_text" size=10 id="date_end" name='date_end'   value='<%=date_end%>' maxlength=10 readonly>
    </td>
    <td width="100" nowrap align="right" class="forms" bgcolor="#E3ECF2" style="padding-right: 10px;">עד תאריך</td>
    </tr>    
    <tr bgcolor="#F3F3F3">
    <td width="295" nowrap align="right" style="padding-right: 10px;" dir="rtl"><select id="cmbChangeType" 
    name="cmbChangeType" class="form_text" runat="server" style="width: 95px;" ></select></td>
    <td width="100" nowrap align="right" class="forms" bgcolor="#E3ECF2" style="padding-right: 10px;">סוג פעילות</td>
    </tr>    
    <tr bgcolor="#F3F3F3">
    <td width="295" nowrap align="right" style="padding-right: 10px;" dir="rtl"><select id="cmbChangeTable" 
    name="cmbChangeTable" class="form_text" runat="server" style="width: 95px;" ></select></td>
    <td width="100" nowrap align="right" class="forms" bgcolor="#E3ECF2" style="padding-right: 10px;">סוג אובייקט</td>
    </tr>
    <tr bgcolor="#F3F3F3"><td align="center" colspan="2"><input type="button" value="דוח לוגים" class="but_menu" 
    style="width: 120px" onclick="window.open('changes.aspx?date_start='+document.getElementById('date_start').value+
    '&date_end='+document.getElementById('date_end').value +
    '&cmbChangeType='+document.getElementById('cmbChangeType').value +
    '&cmbChangeTable='+document.getElementById('cmbChangeTable').value,'ReportChanges','')"></td></tr>
    
   
	</table></td></tr>	
 	</table></td></tr>		
   <tr><td height="40" nowrap ></td></tr>
	<tr><td align=center nowrap ><span style="font-size:20px;">דוח לוגים על ביצוע פעולות בדש בורד מכירות</span></td></tr>
  <tr><td align=center><table bgcolor="#CCCCCC" cellpadding="2" cellspacing="1" width="400" border="0">
  <tr bgcolor="#F3F3F3">
    <td width="295" nowrap align="right" style="padding-right: 10px;">
    						<a href='' onclick='callCalendar(document.getElementById("Form1").date_startSales,"AsdateStartSales");return false;' id='AsdateStartSales'><image src='../../images/calend.gif' border=0  valign=bottom></a>&nbsp;<input dir='ltr' type='text' class="form_text" size=10 id="date_startSales" name='date_startSales'  value='<%=date_startSales%>' maxlength=10 readonly>
    </td>
    <td width="100" nowrap align="right" class="forms" bgcolor="#E3ECF2" style="padding-right: 10px;">מתאריך</td>
    </tr>
    <tr bgcolor="#F3F3F3">
    <td width="295" nowrap align="right" style="padding-right: 10px;">
        						<a href='' onclick='callCalendar(document.getElementById("Form1").date_endSales,"Asdate_endSales");return false;' id='Asdate_endSales'><image src='../../images/calend.gif' border=0  valign=bottom></a>&nbsp;<input dir='ltr' type='text' class="form_text" size=10 id="date_endSales" name='date_endSales'   value='<%=date_endSales%>' maxlength=10 readonly>
    </td>
    <td width="100" nowrap align="right" class="forms" bgcolor="#E3ECF2" style="padding-right: 10px;">עד תאריך</td>
    </tr>   
     <tr bgcolor="#F3F3F3">
    <td width="295" nowrap align="right" style="padding-right: 10px;" dir="rtl"><select id="cmbChangeTableSales" 
    name="cmbChangeTableSales" class="form_text" runat="server" style="width: 125px;" >
    <option value=""></option>
    <option value="111">שליטה חודשי / שנתי</option>
     <option value="112">מסך הגדרות יעדים</option>
    </select></td>
    <td width="100" nowrap align="right" class="forms" bgcolor="#E3ECF2" style="padding-right: 10px;">סוג אובייקט</td>
    </tr> 
    <tr bgcolor="#F3F3F3"><td align="center" colspan="2"><input type="button" value="דוח לוגים" class="but_menu" 
    style="width: 120px" onclick="window.open('changesSales.aspx?date_startSales='+document.getElementById('date_startSales').value+
    '&date_endSales='+document.getElementById('date_endSales').value +
    '&cmbChangeTypeSales='+document.getElementById('cmbChangeTableSales').value +
    '&cmbChangeTableSales='+document.getElementById('cmbChangeTableSales').value,'ReportChangesSales','')"></td></tr>


    </table></td></tr>
 <tr><td height="40" nowrap ></td></tr>

    	<tr><td align=center nowrap ><span style="font-size:20px;">דוח לוגים על שינוי שדה מעוניין בקבלת דיוור</span></td></tr>
  <tr><td align=center><table bgcolor="#CCCCCC" cellpadding="2" cellspacing="1" width="400" border="0">
  <tr bgcolor="#F3F3F3">
    <td width="295" nowrap align="right" style="padding-right: 10px;">
    						<a href='' onclick='callCalendar(document.getElementById("Form1").date_startNewsL,"AsdatestartNewsL");return false;' id='AsdatestartNewsL'><image src='../../images/calend.gif' border=0  valign=bottom></a>&nbsp;<input dir='ltr' type='text' class="form_text" size=10 id="date_startNewsL" name='date_startNewsL'  value='<%=date_startNewsL%>' maxlength=10 readonly>
    </td>
    <td width="100" nowrap align="right" class="forms" bgcolor="#E3ECF2" style="padding-right: 10px;">מתאריך</td>
    </tr>
    <tr bgcolor="#F3F3F3">
    <td width="295" nowrap align="right" style="padding-right: 10px;">
        						<a href='' onclick='callCalendar(document.getElementById("Form1").date_endNewsL,"Asdate_endNewsL");return false;' id='Asdate_endNewsL'><image src='../../images/calend.gif' border=0  valign=bottom></a>&nbsp;<input dir='ltr' type='text' class="form_text" size=10 id="date_endNewsL" name='date_endNewsL'   value='<%=date_endNewsL%>' maxlength=10 readonly>
    </td>
    <td width="100" nowrap align="right" class="forms" bgcolor="#E3ECF2" style="padding-right: 10px;">עד תאריך</td>
    </tr>   
  
    <tr bgcolor="#F3F3F3"><td align="center" colspan="2"><input type="button" value="דוח לוגים" class="but_menu" 
    style="width: 120px" onclick="window.open('ChangesNewsLetters.aspx?date_startNewsL='+document.getElementById('date_startNewsL').value+
    '&date_endNewsL='+document.getElementById('date_endNewsL').value,'ReportChangesNewsLetters','')"></td></tr>
    


	</table>
				 	<SCRIPT LANGUAGE="JavaScript">document.write(getCalendarStyles());</SCRIPT>
			<SCRIPT LANGUAGE="JavaScript" type="text/javascript">
            <!--
            var cal1xx = new CalendarPopup('CalendarDiv');
                cal1xx.showNavigationDropdowns();
                cal1xx.yearSelectStart
                cal1xx.offsetX = -50;
                cal1xx.offsetY = 0;
            //-->
			</SCRIPT>
					<DIV ID='CalendarDiv' name='CalendarDiv' STYLE='VISIBILITY:hidden;POSITION:absolute;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>


	</form>
  </body> 
</HTML>
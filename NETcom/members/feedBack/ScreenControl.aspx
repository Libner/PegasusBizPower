<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ScreenControl.aspx.vb" Inherits="bizpower_pegasus2018.ScreenControl"%>
<%@ Register TagPrefix="DS" TagName="TopIn" Src="../../top_in.ascx" %>
<%@ Register TagPrefix="DS" TagName="LogoTop" Src="../../logo_top.ascx" %>
<%@ Register TagPrefix="DS" TagName="metaInc" Src="../../title_meta_inc.ascx" %>


<html>
	<head>
		<title>screenSetting</title>
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<DS:metaInc id="metaInc" runat="server"></DS:metaInc>
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
		<script src="../../include/CalendarPopup.js" charset="windows-1255"></script>
		<script language="javascript">
function getNumbers(){
	var ch=event.keyCode;
	event.returnValue =((ch >= 48 && ch <= 57) || ch ==46);
}


          function isDate(valDate) {
       var objDate,  // date object initialized from the valDate string 
         mSeconds, // valDate in milliseconds 
         day,      // day 
         month,    // month 
         year;     // year 
       // date length should be 10 characters (no more no less) 
       if (valDate.length !== 10) {
           return false;
       }
       // third and sixth character should be '/' 
       if (valDate.substring(2, 3) !== '/' || valDate.substring(5, 6) !== '/') {
           return false;
       }
       // extract month, day and year from the valDate (expected format is mm/dd/yyyy) 
       // subtraction will cast variables to integer implicitly (needed 
       // for !== comparing) 

       day = valDate.substring(0, 2) - 0; // because months in JS start from 0 
       month = valDate.substring(3, 5) - 1;
       year = valDate.substring(6, 10) - 0;
       // test year range 
       if (year < 1000 || year > 3000) {
           return false;
       }
       // convert valDate to milliseconds 
       mSeconds = (new Date(year, month, day)).getTime();
       // initialize Date() object from calculated milliseconds 
       objDate = new Date();
       objDate.setTime(mSeconds);
       // compare input date and parts from Date() object 
       // if difference exists then date isn't valid 
       if (objDate.getFullYear() !== year ||
           objDate.getMonth() !== month ||
           objDate.getDate() !== day) {
           return false;
       }
       // otherwise return true 
       return true;
   }
 function checkDate(valDate) {
                      // define date string to test 
                      var vDate = document.getElementById(valDate).value;
                      //  alert(vDate)
                      // check date and print message 
                      // alert(isDate(vDate))
                      if (isDate(vDate)) {
                          //      alert('OK');
                      }
                      else {
                          alert('Invalid date format!\nCorrect format dd/MM/YYYY');
               document.getElementById(valDate).value=''               document.getElementById(valDate).focus()        }
                  }
function ClearAll()
{

window.location ="screenControl.aspx"

}
function FuncSort(query,par,srt)
{

var FromGrade,ToGrade,sGuides,sPayFromDate,sPayToDate,sSeries,sTours
FromGrade=$("#FromGrade").val()
ToGrade=$("#ToGrade").val()
sGuides=$("#sGuides").val()
sFromDate=$("#sPayFromDate").val()
sToDate=$("#sPayToDate").val()
sSeries=$("#sSeries").val()
sTours=$("#sTours").val()
sPage=$("#pageList").val()

if (query!="")
{
//alert(query)
var r  ;
r="&"+par +"=ASC"
query = query.replace(r, "");
r="&"+par +"=DESC"
query = query.replace(r, "");
//alert(query)

window.location ="screenControl.aspx?" +query +"&"+par +"="+ srt
//document.forms.Form1.action="screenControl.aspx?" +query +"&"+par +"="+ srt
//document.forms.Form1.submit()

}
else
{ 
window.location ="screenControl.aspx?"+"FromGrade="+ escape(FromGrade) +"&ToGrade="+ escape(ToGrade) +"&sGuides=" + escape(sGuides) +"&sPayFromDate="+ escape(sPayFromDate) +"&sPayToDate="+ escape(sPayToDate) +"&sPage=" + escape(sPage)+ "&sSeries=" + escape(sSeries) + "&sTours="+ escape(sTours) +"&"+par +"="+ srt
//document.forms.Form1.action="screenControl.aspx?sFromDate="+ escape(sFromDate) +"&sToDate="+ escape(sToDate) +"&sPage=" + escape(sPage)+"&"+par +"="+ srt
//document.forms.Form1.submit()

}
}
		</script>
		<script>
		function openDetails(depId)
		{
		//	window.open("gradesByDepId.aspx?depId="+depId, "" ,"scrollbars=1,toolbar=0,top=50,left=10,width=800,height=480,align=center,resizable=0");
	window.open("default_gradesByDepId.asp?depId="+depId,"_parent")
		}
		
		</script>
		<script>
			function SelectSeria()
	{
	
	//alert(document.getElementById("UsersSelect").value)

	var val =document.getElementById("SeriesSelect").value
	var xx=event.clientX -150+ "px";
	newWin=window.open("SelectSeria.aspx?v="+val, "SS", "toolbar=0,menubar=0,width=300,height=320,top=" + event.clientY +",left="+ xx +",scrollbars=auto");
	newWin.opener=window

	}
	function SelectTours()
		{
var val =document.getElementById("ToursSelect").value
	newWin=window.open("SelectTours.aspx?v="+val, "SS", "toolbar=0,menubar=0,width=800,height=600,top=" + event.clientY +",left="+event.clientX +",scrollbars=auto");
	newWin.opener=window
	
	}
	function SelectGuides()
	{
var val =document.getElementById("GuidesSelect").value
	newWin=window.open("SelectGuides.aspx?v="+val, "SS", "toolbar=0,menubar=0,width=500,height=570,top=" + event.clientY +",left="+event.clientX +",scrollbars=auto");
	newWin.opener=window
	
	}
		</script>
	
	</head>
	<body style="margin-right:10px;">
		<form id="Form1" method="post" runat="server" name="Form1">
			<DS:LOGOTOP id="logotop" runat="server"></DS:LOGOTOP>
			<DS:TOPIN id="topin" numOfLink="0" numOfTab="77" toplevel2="79" runat="server"></DS:TOPIN>
			<input type="hidden" id="SeriesSelect" name="SeriesSelect" value="" runat="server">
			<input type="hidden" id="GuidesSelect" name="GuidesSelect" value="" runat="server">
			<input type="hidden" id="ToursSelect" name="ToursSelect" value="" runat="server">
			<table cellpadding="0" cellspacing="0" width="100%" align="left">
			
				<tr>
					<td valign="top" align="left">
						<table border="0" cellpadding="1" cellspacing="1" width="100%" align="left">
							<tr bgcolor="#d8d8d8" style="height:25px">
								<td colspan="3" align="center" class="title_sort">
									<table border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td class="td_admin_5" align="center" width="120"><a href="#" OnClick="javascript:ClearAll();" Class="button_small1">הצג 
													הכל</a>
											</td>
											<td width="50"></td>
											<td class="td_admin_5" align="center" width="120"><asp:LinkButton Runat="server" ID="btnSearch" CssClass="button_small1">חפש</asp:LinkButton></td>
										</tr>
									</table>
								</td>
								<td>&nbsp;</td>
								<td class="title_sort" align="center">
								<table cellspacing="0" cellpadding="0">
								<tr>
								<td class="td_admin_5" align="left">
								<input  type="text" id="FromGrade" class="searchList" style="WIDTH:40px" NAME="FromGrade"  <%if pFromGrade>0 then%> value="<%=pFromGrade%>" <%end if%>  onKeyPress="return getNumbers(this)" maxlength=3>
									-מ
									</td>
								</tr>
								<tr>
								<td class="td_admin_5" align="left">
									<input  type="text" id="ToGrade" class="searchList" style="WIDTH:40px" NAME="ToGrade" onKeyPress="return getNumbers(this)" maxlength=3 <%if pToGrade>0 then%> value="<%=pToGrade%>" <%end if%>>
									עד</td>
								</tr>
                            </table>
									</td>
								<td class="title_sort" align="center" nowrap>
									<a href="#" onclick="SelectGuides('');return false" style="color:#000000;text-decoration:none;vertical-align:middle">
										<img src=../../images/<%if GuidesSelect.Value<>"" and GuidesSelect.Value<>"0" then%>select_act.png<%else%>select.png<%end if%> width=20 border=0 id="SelectGuidesAlt" title="<%=func.GetSelectGuidesName(GuidesSelect.Value)%>"></a>
									<%if false then%>
									<select runat="server" id="sGuides" class="searchList" dir="rtl" name="sGuides">
									</select><%end if%></td>
								<td class="title_sort" align="center" nowrap>
									<table cellspacing="0" cellpadding="0">
								<tr>
								<td class="td_admin_5" align="left">
								<a href="" onclick="cal1xx.select(document.getElementById('sPayFromDate'),'AsPayFromDate','dd/MM/yyyy'); return false;"
										id="AsPayFromDate"><img src="../../images/calendar.gif" border="0" align="center"></a>
									<input runat="server" type="text" id="sPayFromDate" onChange="javascript:checkDate('sPayFromDate')"
										class="searchList" style="WIDTH:70px;" NAME="sPayFromDate"> -מ
									
                            </td>
								</tr>
								<tr>
								<td class="td_admin_5" align="left">
									<a href="" onclick="cal1xx.select(document.getElementById('sPayToDate'),'AsPayToDate','dd/MM/yyyy'); return false;"
										id="AsPayToDate"><img src="../../images/calendar.gif" border="0" align="center"></a>
									<input runat="server" type="text" id="sPayToDate" onChange="javascript:checkDate('sPayToDate')"
										class="searchList" style="WIDTH:70px;" NAME="sPayToDate"> עד
										</td>
								</tr>
                            </table>
								</td>
								<td class="title_sort" align="center">
									<a href="#" onclick="SelectSeria('');return false" class="link1"><img src=../../images/<%if SeriesSelect.Value<>"" and SeriesSelect.Value<>"0" then%>select_act.png<%else%>select.png<%end if%> width=20 border=0 id="SelectSeriasAlt" title="<%'=func.GetSelectUserName(UsersSelect.Value)%>"></a>
									<%if false then%>
									<select runat="server" id="sSeries" class="searchList" dir="rtl" name="sSeries">
									</select><%end if%></td>
								<td class="title_sort" align="center">
									<a href="#" onclick="SelectTours('');return false" class="link1"><img src=../../images/<%if ToursSelect.Value<>"" and ToursSelect.Value<>"0" then%>select_act.png<%else%>select.png<%end if%> width=20 border=0 id="SelectToursAlt" title="<%'=func.GetSelectUserName(UsersSelect.Value)%>"></a>
									<%if false then%>
									<select runat="server" id="sTours" class="searchList" dir="rtl" name="sTours">
									</select><%end if%>&nbsp;</td>
							</tr>
							<tr bgcolor="#d8d8d8" style="height:25px">
								<td class="title_sort" align="center">כמות הטיולים מהם מחושב הציון</td>
								<td class="title_sort" align="center" nowrap>ציון ממוצע של המדריך<Br>
									ב4 שנים</td>
								<td class="title_sort" align="center" nowrap>הלקוחות %<br>
									שמילאו משוב</td>
								<td class="title_sort" align="center" nowrap>/מס' משובים<br>
									/מס' לקוחות שמילאו<br>
									כמות נרשמים<BR>
								</td>
								<td class="title_sort" align="center"><a href="javascript:FuncSort('<%=qrystring%>','sort_5','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_5")="ASC" then %>_act<%end if%>.gif"  title="למיין לפי ציון טיול" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_5','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_5")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי ציון טיול" border=0></a>&nbsp;ציון 
									טיול&nbsp;</td>
								<td class="title_sort" align="center"><a href="javascript:FuncSort('<%=qrystring%>','sort_4','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_4")="ASC" then %>_act<%end if%>.gif"  title="למיין לפי שם מדריך" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_4','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_4")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי שם מדריך" border=0></a>&nbsp;שם 
									מדריך&nbsp;</td>
								<td class="title_sort" align="center"><a href="javascript:FuncSort('<%=qrystring%>','sort_3','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_3")="ASC" then %>_act<%end if%>.gif"  title="למיין לפי תאריך יציאה" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_3','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_3")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי תאריך יציאה" border=0></a>&nbsp;תאריך 
									יציאה&nbsp;</td>
								<td class="title_sort" align="center"><a href="javascript:FuncSort('<%=qrystring%>','sort_2','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_2")="ASC" then %>_act<%end if%>.gif"  title="למיין לפי קוד סידרה" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_2','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_2")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי קוד סידרה" border=0></a>&nbsp;קוד 
									סידרה</td>
								<td class="title_sort" align="right"><a href="javascript:FuncSort('<%=qrystring%>','sort_1','ASC')"><img src="../../../images/arrow_top<%If Request.Querystring("sort_1")="ASC" then %>_act<%end if%>.gif"  title="למיין לפי טול" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_1','DESC')"><img src="../../../images/arrow_bot<%If Request.Querystring("sort_1")="DESC" then %>_act<%end if%>.gif"  title="למיין לפי טיול" border=0></a>&nbsp;טיול&nbsp;</td>
							</tr>
							<tr>
								<td colspan="9" style="width:100%;height:2px;background-color:#C9C9C9"></td>
							</tr>
							<asp:Repeater ID="rptCustomers" Runat="server">
								<ItemTemplate>
									<tr style="<%#func.TourGradeCss(Container.dataItem("Departure_Id"),"background-color: rgb(201, 201, 201)")%>;height:25px;cursor: pointer;"   onClick="javascript:openDetails('<%#Container.Dataitem("Departure_Id")%>')">
										<td align="center">&nbsp;<%#Container.dataItem("Guide_Grade4_Count")%></td>
										<td align="center" nowrap>&nbsp;<%#IIF(func.dbNullFix(Container.dataItem("Guide_Grade4"))="","",Container.dataItem("Guide_Grade4") &"%")%></td>
										<td align="center" id="tdProc" runat="server">&nbsp;</td>
										<td align="center" id="tdRes" runat="server"><%#Container.dataItem("CountMembers")%></td>
										<td align="center" nowrap>&nbsp;<%#IIF(func.TourGrade(Container.dataItem("Departure_Id"))="0","",func.TourGrade(Container.dataItem("Departure_Id")))%></td>
										<td align="center" nowrap><%#Container.Dataitem("Guide_Name")%></td>
										<td align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date", "{0:dd/MM/yyyy}") %></td>
										<td align="center"><%#Container.DataItem("Series_Name")%></td>
										<td align="right"><%#Container.DataItem("Tour_Name")%>&nbsp;</td>
									</tr>
								</ItemTemplate>
								<AlternatingItemTemplate>
									<!--onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#E6E6E6';" -->
									<tr  style="<%#func.TourGradeCss(Container.dataItem("Departure_Id"),"background-color: rgb(230, 230, 230)")%>;height:25px;cursor: pointer;"  onClick="javascript:openDetails('<%#Container.Dataitem("Departure_Id")%>')">
										<td align="center">&nbsp;<%#Container.dataItem("Guide_Grade4_Count")%></td>
										<td align="center" nowrap>&nbsp;<%#IIF(func.dbNullFix(Container.dataItem("Guide_Grade4"))="","",Container.dataItem("Guide_Grade4") &"%")%></td>
										<td align="center" id="tdProc" runat="server">&nbsp;</td>
										<td align="center" nowrap id="tdRes" runat="server"><%#Container.dataItem("CountMembers")%></td>
										<td align="center" nowrap>&nbsp;<%#IIF(func.TourGrade(Container.dataItem("Departure_Id"))="0","",func.TourGrade(Container.dataItem("Departure_Id")))%></td>
										<td align="center" nowrap><%#Container.Dataitem("Guide_Name")%></td>
										<td align="center" nowrap><%#DataBinder.Eval(Container.DataItem, "Departure_Date", "{0:dd/MM/yyyy}")%></td>
										<td align="center"><%#Container.DataItem("Series_Name")%></td>
										<td align="right"><%#Container.DataItem("Tour_Name")%>&nbsp;</td>
									</tr>
								</AlternatingItemTemplate>
							</asp:Repeater>
							<!--  <div id="dvCustomers">-->
							<!--paging-->
							<asp:PlaceHolder id="pnlPages" Runat="server">
								<tr>
									<td height="2" colspan="9"></td>
								</tr>
								<tr>
									<td class="plata_paging" vAlign="top" align="center" colspan="9" bgcolor="#D8D8D8">
										<table dir="ltr" cellSpacing="0" cellPadding="2" width="100%" border="0">
											<tr>
												<td class="plata_paging" vAlign="baseline" noWrap align="left" width="150">&nbsp;הצג
													<asp:DropDownList id="PageSize" CssClass="PageSelect" Runat="server" AutoPostBack="true">
														<asp:ListItem Value="10">10</asp:ListItem>
														<asp:ListItem Value="20">20</asp:ListItem>
														<asp:ListItem Value="50" Selected="True">50</asp:ListItem>
													</asp:DropDownList>&nbsp;בדף&nbsp;
												</td>
												<td vAlign="baseline" noWrap align="right">
													<asp:LinkButton id="cmdNext" runat="server" CssClass="page_link" text="«עמוד הבא"></asp:LinkButton></td>
												<td class="plata_paging" vAlign="baseline" noWrap align="center" width="160">
													<asp:Label id="lblTotalPages" Runat="server"></asp:Label>&nbsp;דף&nbsp;
													<asp:DropDownList id="pageList" CssClass="PageSelect" Runat="server" AutoPostBack="true"></asp:DropDownList>&nbsp;מתוך&nbsp;
												</td>
												<td vAlign="baseline" noWrap align="left">
													<asp:LinkButton id="cmdPrev" runat="server" CssClass="page_link" text="עמוד קודם»"></asp:LinkButton></td>
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
					<td width="100" nowrap valign="top" class="td_menu" style="border: 1px solid #808080;height:400px"
						dir="rtl">
						<table cellpadding="1" cellspacing="0" width="100%" border="0" dir="rtl" ID="Table3">
							<tr>
								<td colspan="2" height="10" nowrap></td>
							</tr>
							<%if false then%>
							<tr>
								<td colspan="2" height="10" nowrap><a class="button_edit_1" style="width:94;" href='javascript:void(0)' onclick="return openReport();">הצג 
										דוח</a></td>
							</tr>
							<%end if%>
							<tr>
								<td colspan="2" height="10" nowrap></td>
							</tr>
						</table>
					</td>
				</tr>
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
</html>

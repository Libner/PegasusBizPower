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
    <link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
 
  <style>
  .ui-widget {
    font-family: Arial,Helvetica,sans-serif;
    font-size: 11px;
}

.ui-menu .ui-menu-item-wrapper {
    position: relative;
    font-size:11px;
    padding: 2px 1em 2px .1em;
}
  .custom-sSeries {
    position: relative;
    display: inline-block;
    
  }
.ui-widget.ui-widget-content {
    border: 1px solid #c5c5c5;
     max-height: 250px;
     overflow-y: scroll;
     overflow-x:hidden;
   
}

  .custom-sSeries-toggle {
    position: absolute;
    top: 0;
    bottom: 0;
    margin-left: -1px;
    padding: 0;
     
  }
  .custom-sSeries-input {
    margin: 0;
    padding: 2px 3px;
   font-size:12px;
   width:30px;
   height:15px;
  }
  </style>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
  <script>
  $( function() {
    $.widget( "custom.sSeries", {
      _create: function() {
        this.wrapper = $( "<span>" )
          .addClass( "custom-sSeries" )
          .insertAfter( this.element );
 
        this.element.hide();
        this._createAutocomplete();
        this._createShowAllButton();
      },
 
      _createAutocomplete: function() {
        var selected = this.element.children( ":selected" ),
          value = selected.val() ? selected.text() : "";
 
        this.input = $( "<input>" )
          .appendTo( this.wrapper )
          .val( value )
          .attr( "title", "" )
          .addClass( "custom-sSeries-input ui-widget ui-widget-content ui-state-default ui-corner-left" )
          .autocomplete({
            delay: 0,
            minLength: 0,
            source: $.proxy( this, "_source" )
          })
          .tooltip({
            classes: {
              "ui-tooltip": "ui-state-highlight"
            }
          });
 
        this._on( this.input, {
          autocompleteselect: function( event, ui ) {
            ui.item.option.selected = true;
            this._trigger( "select", event, {
              item: ui.item.option
            });
          },
 
          autocompletechange: "_removeIfInvalid"
        });
      },
 
      _createShowAllButton: function() {
        var input = this.input,
          wasOpen = false;
 
        $( "<a>" )
          .attr( "tabIndex", -1 )
          .attr( "title", "הכל" )
          .tooltip()
          .appendTo( this.wrapper )
          .button({
            icons: {
              primary: "ui-icon-triangle-1-s"
            },
            text: false
          })
          .removeClass( "ui-corner-all" )
          .addClass( "custom-sSeries-toggle ui-corner-right" )
          .on( "mousedown", function() {
            wasOpen = input.autocomplete( "widget" ).is( ":visible" );
          })
          .on( "click", function() {
            input.trigger( "focus" );
 
            // Close if already visible
            if ( wasOpen ) {
              return;
            }
 
            // Pass empty string as value to search for, displaying all results
            input.autocomplete( "search", "" );
          });
      },
 
      _source: function( request, response ) {
        var matcher = new RegExp( $.ui.autocomplete.escapeRegex(request.term), "i" );
        response( this.element.children( "option" ).map(function() {
          var text = $( this ).text();
          if ( this.value && ( !request.term || matcher.test(text) ) )
            return {
              label: text,
              value: text,
              option: this
            };
        }) );
      },
 
      _removeIfInvalid: function( event, ui ) {
 
        // Selected an item, nothing to do
        if ( ui.item ) {
          return;
        }
 
        // Search for a match (case-insensitive)
        var value = this.input.val(),
          valueLowerCase = value.toLowerCase(),
          valid = false;
        this.element.children( "option" ).each(function() {
          if ( $( this ).text().toLowerCase() === valueLowerCase ) {
            this.selected = valid = true;
            return false;
          }
        });
 
        // Found a match, nothing to do
        if ( valid ) {
          return;
        }
 
        // Remove invalid value
        this.input
          .val( "" )
          .attr( "title", value + " didn't match any item" )
          .tooltip( "open" );
        this.element.val( "" );
        this._delay(function() {
          this.input.tooltip( "close" ).attr( "title", "" );
        }, 2500 );
        this.input.autocomplete( "instance" ).term = "";
      },
 
      _destroy: function() {
        this.wrapper.remove();
        this.element.show();
      }
    });
 
    $( "#sSeries" ).sSeries();
    $( "#toggle" ).on( "click", function() {
      $( "#sSeries" ).toggle();
    });
  } );
  </script>
<script>
function scrollWin() {
    window.scrollTo(0, 0);
}
</script>
<script language=javascript>
function ClearAll()
{
/*document.getElementById("sUsers").value=0;
document.getElementById("Departments").value=0;
document.getElementById("Series").value=0;
sFromD=$("#sPayFromDate").val()
sToD=$("#sPayToDate").val()
sPage=$("#pageList").val()
sStatus=$("#sStatus").val()*/
window.location ="screenSetting.aspx"

}
function openSendReport()
		{
		var query 
		query="<%=qrystring%>"
if (query!="")
{
			window.open("https://www.pegasusisrael.co.il/biz_form/SendreportScreen.aspx?"+query,"ReportExcel","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

		}
if (query=="")
{
var sUser,dep,sSer,sFromD,sToD,sPage,sStatus,sCountry
sUser=$("#sUsers").val()
sdep=$("#sDepartments").val()
sSer =$("#sSeries").val()
sFromD=$("#sPayFromDate").val()
sToD=$("#sPayToDate").val()
sPage=$("#pageList").val()
sStatus=$("#sStatus").val()
sCountry=$("#sCountries").val()
window.open ("https://www.pegasusisrael.co.il/biz_form/SendreportScreen.aspx?"+"sdep="+ escape(sdep) +"&sUser="+ escape(sUser) +"&sSer=" + escape(sSer) +"&sFromD="+ escape(sFromD) +"&sToD="+ escape(sToD) +"&sPage=" + escape(sPage)+ "&sStatus=" + escape(sStatus) + "&sCountry="+ escape(sCountry),"ReportExcel","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

}
		

		}
		function openPdfReport()
		{
			var query 
		query="<%=qrystring%>"
if (query!="")
{
			window.open("https://www.pegasusisrael.co.il/biz_form/reportScreenPdf.aspx?"+query,"ReportPdf","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

		}
if (query=="")
{
var sUser,dep,sSer,sFromD,sToD,sPage,sStatus,sCountry
sUser=$("#sUsers").val()
sdep=$("#sDepartments").val()
sSer =$("#sSeries").val()
sFromD=$("#sPayFromDate").val()
sToD=$("#sPayToDate").val()
sPage=$("#pageList").val()
sStatus=$("#sStatus").val()
sCountry=$("#sCountries").val()
window.open ("https://www.pegasusisrael.co.il/biz_form/reportScreenPdf.aspx?"+"sdep="+ escape(sdep) +"&sUser="+ escape(sUser) +"&sSer=" + escape(sSer) +"&sFromD="+ escape(sFromD) +"&sToD="+ escape(sToD) +"&sPage=" + escape(sPage)+ "&sStatus=" + escape(sStatus) + "&sCountry="+ escape(sCountry),"ReportPdf","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

}
		}
		function openReport()
		{
		var query 
		query="<%=qrystring%>"
if (query!="")
{
			window.open("reportScreen.aspx?"+query,"ReportExcel","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

		}
if (query=="")
{
var sUser,dep,sSer,sFromD,sToD,sPage,sStatus,sCountry
sUser=$("#sUsers").val()
sdep=$("#sDepartments").val()
sSer =$("#sSeries").val()
sFromD=$("#sPayFromDate").val()
sToD=$("#sPayToDate").val()
sPage=$("#pageList").val()
sStatus=$("#sStatus").val()
sCountry=$("#sCountries").val()
window.open ("reportScreen.aspx?"+"sdep="+ escape(sdep) +"&sUser="+ escape(sUser) +"&sSer=" + escape(sSer) +"&sFromD="+ escape(sFromD) +"&sToD="+ escape(sToD) +"&sPage=" + escape(sPage)+ "&sStatus=" + escape(sStatus) + "&sCountry="+ escape(sCountry),"ReportExcel","scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

}
		

		}
function FuncSort(query,par,srt)
{

var sUser,dep,sSer,sFromD,sToD,sPage,sStatus,sCountry
sUser=$("#sUsers").val()
sdep=$("#sDepartments").val()
sSer =$("#sSeries").val()
sFromD=$("#sPayFromDate").val()
sToD=$("#sPayToDate").val()
sPage=$("#pageList").val()
sStatus=$("#sStatus").val()
sCountry=$("#sCountries").val()
if (query!="")
{
//alert(query)
var r  ;
r="&"+par +"=ASC"
query = query.replace(r, "");
r="&"+par +"=DESC"
query = query.replace(r, "");
//alert(query)

window.location ="screenSetting.aspx?" +query +"&"+par +"="+ srt
}
else
{ 
window.location ="screenSetting.aspx?"+"sdep="+ escape(sdep) +"&sUser="+ escape(sUser) +"&sSer=" + escape(sSer) +"&sFromD="+ escape(sFromD) +"&sToD="+ escape(sToD) +"&sPage=" + escape(sPage)+ "&sStatus=" + escape(sStatus) + "&sCountry="+ escape(sCountry) +"&"+par +"="+ srt
}
}
function checkMoveTours()
{

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
var sUser,dep,sSer,sFromD,sToD,sPage,sStatus,sCountry
sUser=$("#sUsers").val()
sdep=$("#sDepartments").val()
sSer =$("#sSeries").val()
sFromD=$("#sPayFromDate").val()
sToD=$("#sPayToDate").val()
sPage=$("#pageList").val()
sStatus=$("#sStatus").val()
sCountry=$("#sCountries").val()


	newWin=window.open("editStatus.aspx?DepartureId=" + escape(departureId) +"&sdep="+ escape(sdep) +"&sUser="+ escape(sUser) +"&sSer=" + escape(sSer) +"&sFromD="+ escape(sFromD) +"&sToD="+ escape(sToD) +"&sPage=" + escape(sPage)+ "&sStatus=" + escape(sStatus) + "&sCountry="+ escape(sCountry), "newWin", "toolbar=0,menubar=0,width=300,height=100,top=100,left=5,scrollbars=auto");
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
<table border=0 cellpadding=1 cellspacing=1 width=100%>

<tr bgcolor=#d8d8d8 style="height:25px" >

	<td class="td_admin_5" align="center" width="50" ><a href="#" OnClick ="javascript:ClearAll();" Class="button_small1">הצג הכל</a>
	<%if false then%><asp:LinkButton Runat="server" ID="btnSearchAll" Width="100%" CssClass="button_small1">הצג הכל</asp:LinkButton><%end if%>
	</td>
									<td class="td_admin_5" align="center"><asp:LinkButton Runat="server" ID="btnSearch" CssClass="button_small1">חפש</asp:LinkButton></td>
<%'IF ViewColumn(13)=true then%><td></td><%'end if%>
<%'IF ViewColumn(11)=true then%><td align=center><select runat="server" id="sUsers" class="searchList" dir="rtl" name="sUsers"></select></td><%'end if%>
<%'IF ViewColumn(12)=true then%><td></td><%'end if%>

<%'IF ViewColumn(6)=true then%><td></td><%'end if%>
<%'IF ViewColumn(7)=true then%><td></td><%'end if%>
<%'IF ViewColumn(8)=true then%><td></td><%'end if%>
<%'IF ViewColumn(9)=true then%><td></td><%'end if%>
<%'IF ViewColumn(5)=true then%><td align=center><select runat="server" id="sDepartments" class="searchList" dir="rtl" name="sDepartments"></select></td><%'end if%>
<%'IF ViewColumn(10)=true then%><td></td><%'end if%>

<%'IF ViewColumn(4)=true then%><td align=center><select runat="server" id="sCountries" class="searchList" dir="rtl" name="sCountries"></select>
</td><%'end if%>

<%'IF ViewColumn(3)=true then%>
	<td class="td_admin_5" align="center" nowrap>
			<a href="" onclick="cal1xx.select(document.getElementById('sPayFromDate'),'AsPayFromDate','dd/MM/yyyy'); return false;"
											id="AsPayFromDate"><img src="../../images/calendar.gif" border="0" align="center"></a>
										<input runat="server" type="text" id="sPayFromDate" class="searchList" style="WIDTH:70px"
											NAME="sPayFromDate"> מי  <br>
	<a href="" onclick="cal1xx.select(document.getElementById('sPayToDate'),'AsPayToDate','dd/MM/yyyy'); return false;"
											id="AsPayToDate"><img src="../../images/calendar.gif" border="0" align="center"></a>
										<input runat="server" type="text" id="sPayToDate" class="searchList" style="WIDTH:70px"
											NAME="sPayToDate"> עד 
							</td><%'end if%>

<%'IF ViewColumn(2)=true then%><td class="td_admin_5">
<div class="ui-widget">
<select runat="server" id="sSeries" dir="rtl" name="sSeries"  ></select>
</div>
</td><%'end if%>

<%'IF ViewColumn(1)=true then%><td align=center class="td_admin_5"><select runat="server" id="sStatus" style ="font-size:12px;width:100px"  dir="rtl" name="sStatus"></select></td>
<%'end if%>
	<%if Request.Cookies("bizpegasus")("ScreenTourVisible")="1" then%>
<td align=center class="td_admin_5"></td>
<%end if%>
</tr>
<tr><td colspan=17 style="width:100%;height:2px;background-color:#C9C9C9"></td></tr>
<tr >
<td  class="title_sort">&nbsp;</td>

<asp:Repeater ID=rptTitle runat=server>
<ItemTemplate>
<td class="title_sort" align=center><a href="javascript:FuncSort('<%=qrystring%>','sort_<%#Container.DataItem("Column_Id")%>','ASC')"><img src="../../../images/arrow_top<%#IIf (Request.Querystring("sort_"& Container.DataItem("Column_Id"))="ASC","_act","")%>.gif"  title="למיין לפי <%#Container.DataItem("Column_Name")%>" border=0></a><a href="javascript:FuncSort('<%=qrystring%>','sort_<%#Container.DataItem("Column_Id")%>','DESC')"><img src="../../../images/arrow_bot<%#IIf (Request.Querystring("sort_"& Container.DataItem("Column_Id"))="DESC","_act","")%>.gif"  title="למיין לפי <%#Container.DataItem("Column_Name")%>" border=0></a>&nbsp;<%#Container.DataItem("Column_Name")%></td>

</ItemTemplate>
<FooterTemplate><%if Request.Cookies("bizpegasus")("ScreenTourVisible")="1" then%><td class="title_sort" align=center>מופיע <BR>במסך צפיה</td><%end if%>
	<%if false then%><td align="center"  class="title_sort" colspan="3" nowrap>סדר</td>><%end if%>
</FooterTemplate>
</asp:Repeater>

<asp:Repeater ID=rptCustomers Runat=server>
<ItemTemplate>
<tr bgcolor="#C9C9C9" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';" style="background-color: rgb(201, 201, 201);">
<td>&nbsp;</td>
<%'IF ViewColumn(13)=true then%><td align=center class=price><%#Container.Dataitem("Price")%></td><%'end if%>
<%'IF ViewColumn(12)=true then%><td align=center class="FlyingCompany"><%#Container.DataItem("Flying_Company")%></td><%'end if%>
<%'IF ViewColumn(11)=true then%><td align=center class="UserName"><%#Container.DataItem("User_Name")%></td><%'end if%>
<%'IF ViewColumn(6)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountlastMonth"))%><%'#func.GetCountLastMonth(Container.DataItem("Departure_Code"))%></td><%'end if%>
<%'IF ViewColumn(7)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("Countlast2Week"))%><%'#func.GetCountLast2Week(Container.DataItem("Departure_Code"))%></td><%'end if%>
<%'IF ViewColumn(8)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountlastWeek"))%><%'#func.GetCountLastDateSale(Container.DataItem("Departure_Code"))%></td><%'end if%>
<%'IF ViewColumn(9)=true then%><td align=center nowrap><%#func.dbNullFix(Container.DataItem("LastDateSale"))%><%'#func.GetLastDateSale(Container.DataItem("Departure_Code"))%></td><%'end if%>
<%'IF ViewColumn(10)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountMembersBitulim"))%><%'#func.GetCountMembersBitulim(Container.DataItem("Departure_Code"))%></td><%'end if%>

<%'IF ViewColumn(5)=true then%><td align=center><%#Container.DataItem("Dep_Name")%></td><%'end if%>
<%'IF ViewColumn(4)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountMembers"))%><%'#func.GetCountMembers(Container.DataItem("Departure_Code"))%></td><%'end if%>
<td align=center><%#Container.Dataitem("CountryName")%></td>
<%'IF ViewColumn(3)=true then%><td align=center><%#Container.Dataitem("Date_Begin")%></td><%'end if%>
<td align=center><%#Container.DataItem("Series_Name")%></td><%'#IIF(ViewColumn(2),"<td align=center>" & Container.DataItem("Series_Name") &"</td>","")%>
<%'IF ViewColumn(1)=true then%>
<td align=right style="padding-top:1px;padding-bottom:1px">
<table border=0 style="background-color:<%#Container.DataItem("Status_Color")%>;width:100%" cellpadding=1 cellspacing=1 align=right>
<tr><td align=right><span style="color:#000000">&nbsp;<%#Container.DataItem("Status_Name")%>&nbsp;</span></td>
<td align=right>
<%if Request.Cookies("bizpegasus")("ScreenTourStatus")="1" then%><a href="#" class="link_categ" onclick="openStatus('<%#Container.DataItem("Departure_Id")%>');return false" title="עדכון סטטוס" style="cursor:hand"><img src="../../images/edit_icon.gif" border="0" style="margin-left:1px;"></a><%end if%></td>
</tr></table></td><%'end if%>
	<%if Request.Cookies("bizpegasus")("ScreenTourVisible")="1" then%>
<td align=center><input type="checkbox" ID="chkTour" Runat="server" NAME="chkTour"></td>
<%end if%>
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

<%'IF ViewColumn(13)=true then%><td align=center><%#Container.Dataitem("Price")%></td><%'end if%>
<%'IF ViewColumn(12)=true then%><td align=center><%#Container.DataItem("Flying_Company")%></td><%'end if%>
<%'IF ViewColumn(11)=true then%><td align=center><%#Container.DataItem("User_Name")%></td><%'end if%>
<%'IF ViewColumn(6)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountlastMonth"))%><%'#func.GetCountLastMonth(Container.DataItem("Departure_Code"))%></td><%'end if%>
<%'IF ViewColumn(7)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("Countlast2Week"))%><%'#func.GetCountLast2Week(Container.DataItem("Departure_Code"))%></td><%'end if%>
<%'IF ViewColumn(8)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountlastWeek"))%><%'#func.GetCountLastDateSale(Container.DataItem("Departure_Code"))%></td><%'end if%>
<%'IF ViewColumn(9)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("LastDateSale"))%><%'#func.GetLastDateSale(Container.DataItem("Departure_Code"))%></td><%'end if%>
<%'IF ViewColumn(10)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountMembersBitulim"))%><%'#func.GetCountMembersBitulim(Container.DataItem("Departure_Code"))%></td><%'end if%>


<%'IF ViewColumn(5)=true then%><td align=center><%#Container.DataItem("Dep_Name")%></td><%'end if%>
<%'IF ViewColumn(4)=true then%><td align=center><%#func.dbNullFix(Container.DataItem("CountMembers"))%><%'#func.GetCountMembers(Container.DataItem("Departure_Code"))%></td><%'end if%>
<td align=center><%#Container.Dataitem("CountryName")%></td>

<%'IF ViewColumn(3)=true then%><td align=center><%#Container.Dataitem("Date_Begin")%></td><%'end if%>
<td align=center><%#Container.DataItem("Series_Name")%></td><%'#IIF(ViewColumn(2),"<td align=center>" & Container.DataItem("Series_Name") &"</td>","")%>
<%'IF ViewColumn(1)=true then%>
<td align=right style="padding-top:1px;padding-bottom:1px">
<table border=0 style="background-color:<%#Container.DataItem("Status_Color")%>;width:100%" cellpadding=1 cellspacing=1 align=right>
<tr><td align=right><span style="color:#000000">&nbsp;<%#Container.DataItem("Status_Name")%>&nbsp;</span></td>
<td align=right>
<%if Request.Cookies("bizpegasus")("ScreenTourStatus")="1" then%>
<a href="#" class="link_categ" onclick="openStatus('<%#Container.DataItem("Departure_Id")%>');return false" title="עדכון סטטוס" style="cursor:hand"><img src="../../images/edit_icon.gif" border="0" style="margin-left:1px;"></a><%end if%></td>
</tr></table></td><%'end if%>
	<%if Request.Cookies("bizpegasus")("ScreenTourVisible")="1" then%>
<td align=center> <input type="checkbox" ID="chkTour" Runat="server"  NAME="chkTour" ></td>
<%end if%>

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
 	<tr>
<td class="title_sort">&nbsp;</td>
<td class="title_sort" align="center"></td>
<td class="title_sort" align="center"></td>
<td class="title_sort" align="center">4</td>
<td class="title_sort" align="center">5</td>
<td class="title_sort" align="center">6--</td>
<td class="title_sort" align="center">7</td>
<td class="title_sort" align="center">8</td>
<td class="title_sort" align="center">9</td>
<td class="title_sort" align="center">10</td>
<td class="title_sort" align="center"><%=allCountMembers%></td>
<td class="title_sort" align="center">12</td>
<td class="title_sort" align="center">13</td>
<td class="title_sort" align="center">14</td>
<td class="title_sort" align="center">15</td>
<td class="title_sort" align="center">16</td>
</tr>
						<asp:PlaceHolder id="pnlPages" Runat="server">
						<tr><td height="2" colspan="13"></td></tr>
						<tr>
							<td class="plata_paging" vAlign="top" align="center" colspan="18" bgcolor=#D8D8D8>
								<table dir="ltr" cellSpacing="0" cellPadding="2" width="100%" border="0" >
									<tr>
										<td class="plata_paging" vAlign="baseline" noWrap align="left" width="150">&nbsp;הצג
											<asp:DropDownList id="PageSize" CssClass="PageSelect" Runat="server" AutoPostBack="true">
													<asp:ListItem Value="10">10</asp:ListItem>
												<asp:ListItem Value="20">20</asp:ListItem>
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
  <td width=100 nowrap valign=top class="td_menu" style="border: 1px solid #808080" dir="rtl">
    <table cellpadding=1 cellspacing=0 width="100%" border="0" dir="rtl" ID="Table3">
    <tr><td align="right" colspan="2" class="title_search">&nbsp;</td></tr>
		<%if Request.Cookies("bizpegasus")("ScreenTourVisible")="1" then%>
	<tr><td colspan="2" height=10 nowrap>
	<asp:button  ID=btnSaveChecked EnableViewState="false" runat="server" Cssclass="button_edit_1" Text="עדכן מסך צפייה"></asp:Button> </td></tr>
	<%end if%>

	<%if false then%><a class="button_edit_1" style="width:94;" href='javascript:void(0)' onclick="return checkMoveTours();">עדכן מסך צפייה<%end if%>
			<tr><td colspan="2" height=10 nowrap></td></tr>
			<tr><td colspan="2" height=10 nowrap>
<a class="button_edit_1" style="width:94;" href='javascript:void(0)' onclick="return openReport();">הצג דוח Excel</a></td></tr>
	<tr><td colspan="2" height=10 nowrap></td></tr>
    <tr><td colspan="2" height=10 nowrap>
<a class="button_edit_1" style="width:94;" href='javascript:void(0)' onclick="return openPdfReport();">הצג דוח PDF</a></td></tr>
	<tr><td colspan="2" height=10 nowrap></td></tr>
	<%if Request.Cookies("bizpegasus")("ScreenSendMail")="1" then%>
			<tr><td colspan="2" height=10 nowrap>
<a class="button_edit_1" style="width:94;" href='javascript:void(0)' onclick="return openSendReport();">שליחת מייל</a></td></tr>
<%end if%>
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

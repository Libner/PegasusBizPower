<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="logs.aspx.vb" Inherits="bizpower_pegasus2018.logsRegistrationSendings"%>
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
		<link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
		<script src="//code.jquery.com/ui/1.11.4/jquery-ui.min.js"></script>
		<script src="../../include/CalendarPopup.js" charset="windows-1255"></script>
		<style>
		/*.search {
			font-size:12px
		}*/
        .custom-sSeries
        {
            position: relative;
            display: inline-block;
        }
        .custom-sSeries-toggle
        {
            position: absolute;
            top: 0;
            bottom: 0;
            margin-left: -1px;
            padding: 0;
        }
        .custom-sSeries-input
        {
            margin: 0;
            /*width: 58px;*/
            direction: rtl;
            padding: 2px 5px;
            height: 21px;
        }
        li
        {
			font-weight:normal
        }
		</style>
		<style>
        .ui-widget
        {
            font-family: Arial;
            font-size: 11px;
            direction: rtl;
        }
        
        .ui-menu .ui-menu-item-wrapper
        {
            position: relative;
            font-size: 12px;
            padding: 2px 1em 2px .1em;
        }
        
        .ui-widget.ui-widget-content
        {
            border: 1px solid #c5c5c5;
            background: #ffffff;
            max-height: 250px;
           max-width: 200px;
            /*width: 58px;*/
            overflow-y: scroll;
            overflow-x: hidden;
        }
        .ui-button, .ui-button:link, .ui-button:visited, .ui-button:hover, .ui-button:active
        {
            text-decoration: none;
            width: 16px;
        }
         .ui-button
        {
            margin-right: -1.1em;
        }
        .ui-menu .ui-menu-item {
    padding: 3px 3px 1px 3px;
    line-height: normal;
    }
		</style>
		<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
		<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
		<script>
		window.name="mainwindow";

        $(function () {
            $.widget("custom.sSeries", {
                _create: function () {
                    this.wrapper = $("<span>")
          .addClass("custom-sSeries")
          .insertAfter(this.element);

                    this.element.hide();
                    this._createAutocomplete();
                    this._createShowAllButton();
                },

                _createAutocomplete: function () {
                    var selected = this.element.children(":selected"),
          value = selected.val() ? selected.text() : "";

                    this.input = $("<input>")
          .appendTo(this.wrapper)
          .val(value)
          .attr("title", "")
          .css('width', $(this.element).width()-5)
          .addClass("custom-sSeries-input ui-widget ui-widget-content ui-state-default ui-corner-left")
           .autocomplete({
              delay: 0,
              minLength: 0,
              source: $.proxy(this, "_source")
          })
          .tooltip({
              classes: {
                  "ui-tooltip": "ui-state-highlight"
              }
          });

                    this._on(this.input, {
                        autocompleteselect: function (event, ui) {
                            ui.item.option.selected = true;
                            this._trigger("select", event, {
                                item: ui.item.option
                            });
                        },

                        autocompletechange: "_removeIfInvalid"
                    });
                },

                _createShowAllButton: function () {
                    var input = this.input,
          wasOpen = false;

                    $("<a>")
          .attr("tabIndex", -1)
          .attr("title", "הכל")
          .tooltip()
          .appendTo(this.wrapper)
          .button({
              icons: {
                  primary: "ui-icon-triangle-1-s"
              },
              text: false
          })
          .removeClass("ui-corner-all")
          .addClass("custom-sSeries-toggle ui-corner-right")
          .on("mousedown", function () {
              wasOpen = input.autocomplete("widget").is(":visible");
          })
          .on("click", function () {
              input.trigger("focus");

              // Close if already visible
              if (wasOpen) {
                  return;
              }

              // Pass empty string as value to search for, displaying all results
              input.autocomplete("search", "");
          });
                },

                _source: function (request, response) {
                    var matcher = new RegExp($.ui.autocomplete.escapeRegex(request.term), "i");
                    response(this.element.children("option").map(function () {
                        var text = $(this).text();
                        if (this.value && (!request.term || matcher.test(text)))
                            return {
                                label: text,
                                value: text,
                                option: this
                            };
                    }));
                },

                _removeIfInvalid: function (event, ui) {

                    // Selected an item, nothing to do
                    if (ui.item) {
                        return;
                    }

                    // Search for a match (case-insensitive)
                    var value = this.input.val(),
          valueLowerCase = value.toLowerCase(),
          valid = false;
                    this.element.children("option").each(function () {
                        if ($(this).text().toLowerCase() === valueLowerCase) {
                            this.selected = valid = true;
                            return false;
                        }
                    });

                    // Found a match, nothing to do
                    if (valid) {
                        return;
                    }

                    // Remove invalid value
                    this.input
          .val("")
          .attr("title", value + " didn't match any item")
          .tooltip("open");
                    this.element.val("");
                    this._delay(function () {
                        this.input.tooltip("close").attr("title", "");
                    }, 2500);
                    this.input.autocomplete("instance").term = "";
                },

                _destroy: function () {
                    this.wrapper.remove();
                    this.element.show();
                }
            });


			/*$("#sContactName").sSeries();
            $("#toggle").on("click", function () {
                $("#sContactName").toggle();
            });*/
            
           /* $("#sCountries").sSeries();
            $("#toggle").on("click", function () {
                $("#sCountries").toggle();
            });*/
        });
		</script>
		<script>
        function scrollWin() {
            window.scrollTo(0, 0);
        }
		</script>
		<script language="javascript">
        function ClearAll() {
            window.location = "logScreen.aspx"

        }
       
        function openReport() {
            var query
            query = "<%=qrystring%>"
            if (query != "") {
                window.open("reportScreen.aspx?" + query, "ReportExcel", "scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

            }
            if (query == "") {
                var querySearch
				querySearch=""
				$("#blockSearchTerms").find("input").each(function () {
                if($(this).val()!="" && $(this).attr("name") != '' && $(this).attr("name") != 'undefined')
					{
					if(querySearch!="") {querySearch=querySearch + "&"}			
					querySearch=querySearch + $(this).attr("name") + "=" + $(this).val()
					}
				});
				$("#blockSearchTerms").find("select").each(function () {
                if($(this).val()!="")
					{
					if(querySearch!="") {querySearch=querySearch + "&"}			
					querySearch=querySearch + $(this).attr("name") + "=" +$(this).val()
					}
				});
				
				
                window.open("reportScreen.aspx?" + querySearch  , "ReportExcel", "scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

            }


        }
        function FuncSort(query, par, srt) {

           query=""
            if (query != "") {
                //alert(query)
                var r;
                r = "&" + par + "=ASC"
                query = query.replace(r, "");
                r = "&" + par + "=DESC"
                query = query.replace(r, "");
                //alert(query)

//alert("logScreen.aspx?" + query + "&" + par + "=" + srt)
                window.location = "logs.aspx?" + query + "&" + par + "=" + srt
            }
            else {
				var querySearch
				querySearch=""
				
				$("#blockSearchTerms").find("input").each(function () {
                if($(this).val()!="" && $(this).attr("name"))
					{
					if(querySearch!="") {querySearch=querySearch + "&"}			
					querySearch=querySearch + $(this).attr("name") + "=" + encodeURIComponent($(this).val())
					}
				});
				$("#blockSearchTerms").find("select").each(function () {
                if($(this).val()!="" && $(this).attr("name"))
					{
					if(querySearch!="") {querySearch=querySearch + "&"}			
					querySearch=querySearch + $(this).attr("name") + "=" + encodeURIComponent($(this).val())
					}
				});
//alert("logScreen.aspx?" + querySearch + "&" + par + "=" + srt)
                window.location = "logs.aspx?" + querySearch + "&" + par + "=" + srt
            }
        }
        function checkMoveTours() {

        }
        
        



		</script>
	</head>
	<body>
		<form id="Form1" method="post" runat="server" autocomplete="off">
			<DS:LOGOTOP id="logotop" runat="server"></DS:LOGOTOP>
			<DS:TOPIN id="topin" numOfLink="5" numOfTab="108" topLevel2="155" runat="server"></DS:TOPIN>
			<table border="0" cellpadding="0" cellspacing="0" >
							<tr>
											
								<td height="30" align="left" style="WIDTH: 232px;">
												<a class="but_menu" style="WIDTH: 220px;padding:6px;BACKGROUND-COLOR: #98fc00;color:#000000"
												 href="default.aspx">חזרה למסך שליחת טפסי רישום</a>
											</td>
							</tr>
							<tr><td height="10" colspan="2"></td></tr>
						</table>
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td valign="top">
						<table border="0" cellpadding="0" cellspacing="0" width="100%">
							
						
							<tr>
								<td align="center">
									<table border="0" cellpadding="1" cellspacing="1" dir="rtl">
										<tr bgcolor="#d8d8d8" id="blockSearchTerms">
											<td></td>
											
											<td class="td_admin_5" align="center" nowrap dir="ltr">
												<table cellspacing="0" cellpadding="0">
													<tr>
														<td class="td_admin_4" align="left" nowrap>
															<a href="" onclick="cal1xx.select(document.getElementById('sSendingFromDate'),'AsSendingFromDate','dd/MM/yyyy'); return false;"
																id="AsSendingFromDate" title="תאריך שליחה"><img src="../../images/calendar.gif" border="0" align="center"></a>
															<input runat="server" type="text" id="sSendingFromDate" class="search" style="width: 70px"
																name="sSendingFromDate" title="תאריך שליחה"> מי
														</td>
													</tr>
													<tr>
														<td class="td_admin_4" align="left" nowrap>
															<a href="" onclick="cal1xx.select(document.getElementById('sSendingToDate'),'AsSendingToDate','dd/MM/yyyy'); return false;"
																id="AsSendingToDate" title="תאריך שליחה"><img src="../../images/calendar.gif" border="0" align="center"></a>
															<input runat="server" type="text" id="sSendingToDate" class="search" style="width: 70px"
																name="sSendingToDate" title="תאריך שליחה"> עד
														</td>
													</tr>
												</table>
											</td>
												<td class="td_admin_5" nowrap width="100">
													<select runat="server" id="sCountries" dir="rtl" name="sCountries">
												</select>
											</td>
											<td class="td_admin_5" nowrap width="100">
												<input runat="server" id="sCountSendings" dir="rtl" name="sCountSendings" class="search" style="width:50px">
											</td>
											<td class="td_admin_5" align="center" width="50" nowrap>
												<asp:LinkButton runat="server" ID="btnSearch" CssClass="button_small1" BackColor="#ff9900">חפש</asp:LinkButton>
											</td>
											<td class="td_admin_5" align="center" width="50" nowrap>
												<a href="logs.aspx" class="button_small1">הצג הכל</a>
											</td>
										</tr>
														<tr>
											<td class="td_admin_5" align="center" nowrap>
												סדר
											</td>
											<asp:Repeater ID="rptTitle" runat="server">
												<ItemTemplate>
													<td class="title_sort" align="center">
														<asp:PlaceHolder ID="phSortArrows" Runat="server" Visible="true">
                        <a href="javascript:FuncSort('<%=qrystring%>','sort_<%#Container.DataItem("Column_Id")%>','ASC')"><img src="../../../images/arrow_top<%#IIf (Request.Querystring("sort_"& Container.DataItem("Column_Id"))="ASC","_act","")%>.gif"
                                title="למיין לפי <%#Container.DataItem("Column_Title")%>" border="0"></a><a href="javascript:FuncSort('<%=qrystring%>','sort_<%#Container.DataItem("Column_Id")%>','DESC')"><img
                                    src="../../../images/arrow_bot<%#IIf (Request.Querystring("sort_"& Container.DataItem("Column_Id"))="DESC","_act","")%>.gif"
                                    title="למיין לפי <%#Container.DataItem("Column_Title")%>" border="0"></a>&nbsp;
                                    </asp:PlaceHolder>
														<%#Container.DataItem("Column_Title")%>
													</td>
												</ItemTemplate>
											</asp:Repeater>
													<td colspan="2" class="td_admin_5" ></td>
										</tr>
										
										<asp:Repeater ID="rptLogs" runat="server">
											<ItemTemplate>
												<tr bgcolor="#C9C9C9" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';"
													style="background-color: rgb(201, 201, 201);">
													<td class="td_admin_4" align="center">
														<font color="#060165"><b>&nbsp;<%#Container.ItemIndex+1%>&nbsp;</b></font>
													</td>
													<td align="center" class="td_admin_4">
													<%#func.dbNullDate(Container.DataItem("Job_Running_Date"),"dd/MM/yyyy HH:mm")%>
													</td>
													<td align="center" class="td_admin_4">
														<%#func.dbNullFix(Container.DataItem("Countries_List"))%>
													</td>
													<td align="right" class="td_admin_4">
														<%#func.fixnumeric(Container.DataItem("SentMail_Count"))%>
													</td>
													<td colspan="2"></td>
												</tr>
											</ItemTemplate>
											<AlternatingItemTemplate>
												<tr bgcolor="#E6E6E6" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#E6E6E6';"
													style="background-color: rgb(230, 230, 230);">
													<td class="td_admin_4" align="center">
														<font color="#060165"><b>&nbsp;<%#Container.ItemIndex+1%>&nbsp;</b></font>
													</td>
													<td align="center" class="td_admin_4">
													<%#func.dbNullDate(Container.DataItem("Job_Running_Date"),"dd/MM/yyyy HH:mm")%>
													</td>
													<td align="center" class="td_admin_4">
														<%#func.dbNullFix(Container.DataItem("Countries_List"))%>
													</td>
													<td align="right" class="td_admin_4">
														<%#func.fixnumeric(Container.DataItem("SentMail_Count"))%>
													</td>
													<td colspan="2"></td>
												</tr>
											</AlternatingItemTemplate>
										</asp:Repeater>
										<asp:PlaceHolder ID="pnlPages" runat="server">
											<tr>
												<td height="2" colspan="14">
												</td>
											</tr>
											<tr>
												<td class="plata_paging" valign="top" align="center" colspan="18" bgcolor="#D8D8D8">
													<table dir="ltr" cellspacing="0" cellpadding="2" width="100%" border="0">
														<tr>
															<td class="plata_paging" valign="baseline" nowrap align="left" width="150">
																&nbsp;הצג
																<asp:DropDownList ID="PageSize" CssClass="PageSelect" runat="server" AutoPostBack="true">
																	<asp:ListItem Value="10">10</asp:ListItem>
																	<asp:ListItem Value="30">30</asp:ListItem>
																	<asp:ListItem Value="50">50</asp:ListItem>
																	<asp:ListItem Value="100" Selected="True">100</asp:ListItem>
																</asp:DropDownList>
																&nbsp;בדף&nbsp;
															</td>
															<td valign="baseline" nowrap align="right">
																<asp:LinkButton ID="cmdNext" runat="server" CssClass="page_link" Text="«עמוד הבא"></asp:LinkButton>
															</td>
															<td class="plata_paging" valign="baseline" nowrap align="center" width="160">
																<asp:Label ID="lblTotalPages" runat="server"></asp:Label>&nbsp;דף&nbsp;
																<asp:DropDownList ID="pageList" CssClass="PageSelect" runat="server" AutoPostBack="true"></asp:DropDownList>
																&nbsp;מתוך&nbsp;
															</td>
															<td valign="baseline" nowrap align="left">
																<asp:LinkButton ID="cmdPrev" runat="server" CssClass="page_link" Text="עמוד קודם»"></asp:LinkButton>
															</td>
															<td class="plata_paging" valign="baseline" align="right">
																&nbsp;נמצאו&nbsp;&nbsp;&nbsp;
																<asp:Label CssClass="plata_paging" ID="lblCount" runat="server"></asp:Label>&nbsp;יציאות
															</td>
														</tr>
													</table>
												</td>
											</tr>
										</asp:PlaceHolder>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<asp:DataGrid ID="dtTest" Runat="server" EnableViewState="false"></asp:DataGrid>
			<script LANGUAGE="JavaScript">
    document.write(getCalendarStyles());
			</script>
			<script type="text/javascript">
            <!--
        var cal1xx = new CalendarPopup('CalendarDiv');
        cal1xx.setYearSelectStartOffset(120);
        cal1xx.setYearSelectEndOffset(2);
        //cal1xx.setYearSelectStart(1910); //Added by Mila
        cal1xx.showNavigationDropdowns();
        cal1xx.offsetX = -50;
        cal1xx.offsetY = -40;
          
                //-->
			</script>
			<div id='CalendarDiv' style='position: absolute; visibility: hidden; background-color: white;
        layer-background-color: white'>
			</div>
		</form>
		<!-- <form id="frmSendform" name="frmSendform" method="post" target="ifrPost" action="">
		<input type="hidden" id="itemId" name="itemId" value="">
		<input type="hidden" id="usid" name="usid"  value="<%=userId%>">
		<input type="hidden" id="squery" name="squery" value="<%=queryParamSearch%>">
    <form>
    <iframe id="ifrPost" height="100" width="100" ></iframe>-->
		<!--   <a href="" onclick="window.open('<%=ConfigurationSettings.AppSettings.Item("PegasusUrl")%>/biz_form/email_ReservationFormReturnSending.aspx?usid=<%=userId%>&Forms=343475&squery=<%=queryParamSearch%>','','top=20,left=50,width=600,height=550,scrollbars=1,resizable=1,menubar=1');return false">hhhhhhhhhhhhhh</a>-->
	</body>
</html>

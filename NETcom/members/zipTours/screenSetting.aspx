<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="screenSetting.aspx.vb"
    Inherits="bizpower_pegasus2018.screenSetting" %>
<%@ Register TagPrefix="DS" TagName="TopIn" Src="../../top_in.ascx" %>
<%@ Register TagPrefix="DS" TagName="LogoTop" Src="../../logo_top.ascx" %>
<%@ Register TagPrefix="DS" TagName="metaInc" Src="../../title_meta_inc.ascx" %>


<html>
<head>
    <title>screenSetting</title>
    	<DS:metaInc id="metaInc" runat="server"></DS:metaInc>
	
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
 
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
    <script src="../../include/CalendarPopup.js" charset="windows-1255"></script>
    <link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
    <link rel="stylesheet" href="/resources/demos/style.css">
    <script src="https://ajax.googleapis.com/ajax/libs/webfont/1.4.7/webfont.js"></script>
    <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.6.4/jquery.min.js"></script>
    <style>
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
            width: 75px;
            direction: rtl;
            padding: 2px 5px;
            height: 23px;
            font-weight:normal;
        }
        .ui-button, .ui-button:link, .ui-button:visited, .ui-button:hover, .ui-button:active
        {
            text-decoration: none;
            width: 16px;
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
            font-weight:normal;
            padding: 2px 1em 2px .1em;
        }
        
        .ui-widget.ui-widget-content
        {
            border: 1px solid #c5c5c5;
            background: #ffffff;
            max-height: 250px;
            width: 68px;
            overflow-y: scroll;
            overflow-x: hidden;
        }
    </style>
    <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
    <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
    <script>
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

            $("#sSeries").sSeries();
            $("#toggle").on("click", function () {
                $("#sSeries").toggle();
            });
        });
    </script>
    <script>
        function scrollWin() {
            window.scrollTo(0, 0);
        }
    </script>
    <script language="javascript">
        function ClearAll() {
            /*document.getElementById("sUsers").value=0;
            document.getElementById("Departments").value=0;
            document.getElementById("Series").value=0;
            sFromD=$("#sPayFromDate").val()
            sToD=$("#sPayToDate").val()
            sPage=$("#pageList").val()
            sStatus=$("#sStatus").val()*/
            window.location = "screenSetting.aspx"

        }
        function openFile() {
            window.open("PriceEdit.aspx", "PriceEdit", "scrollbars=yes,menubar=yes, toolbar=yes, width=1000,height=600,left=10,top=10");

        }
        function openSendReport() {
            var query
            query = "<%=qrystring%>"
            if (query != "") {
                window.open("https://www.pegasusisrael.co.il/biz_form/SendreportScreen.aspx?" + query, "ReportExcel", "scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

            }
            if (query == "") {
                var sUser, dep, sSer, sFromD, sToD, sPage, sStatus, sCountry
                sUser = $("#sUsers").val()
                sdep = $("#sDepartments").val()
                sSer = $("#sSeries").val()
                sFromD = $("#sPayFromDate").val()
                sToD = $("#sPayToDate").val()
                sPage = $("#pageList").val()
                sStatus = $("#sStatus").val()
                sCountry = $("#sCountries").val()
                window.open("https://www.pegasusisrael.co.il/biz_form/SendreportScreen.aspx?" + "sdep=" + escape(sdep) + "&sUser=" + escape(sUser) + "&sSer=" + escape(sSer) + "&sFromD=" + escape(sFromD) + "&sToD=" + escape(sToD) + "&sPage=" + escape(sPage) + "&sStatus=" + escape(sStatus) + "&sCountry=" + escape(sCountry), "ReportExcel", "scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

            }


        }
        function openPdfReport() {
            var query
            query = "<%=qrystring%>"
            if (query != "") {
                window.open("https://www.pegasusisrael.co.il/biz_form/reportScreenPdf.aspx?" + query, "ReportPdf", "scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

            }
            if (query == "") {
                var sUser, dep, sSer, sFromD, sToD, sPage, sStatus, sCountry
                sUser = $("#sUsers").val()
                sdep = $("#sDepartments").val()
                sSer = $("#sSeries").val()
                sFromD = $("#sPayFromDate").val()
                sToD = $("#sPayToDate").val()
                sPage = $("#pageList").val()
                sStatus = $("#sStatus").val()
                sCountry = $("#sCountries").val()
                window.open("https://www.pegasusisrael.co.il/biz_form/reportScreenPdf.aspx?" + "sdep=" + escape(sdep) + "&sUser=" + escape(sUser) + "&sSer=" + escape(sSer) + "&sFromD=" + escape(sFromD) + "&sToD=" + escape(sToD) + "&sPage=" + escape(sPage) + "&sStatus=" + escape(sStatus) + "&sCountry=" + escape(sCountry), "ReportPdf", "scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

            }
        }
        function openReport() {
            var query
            query = "<%=qrystring%>"
            if (query != "") {
                window.open("reportScreen.aspx?" + query, "ReportExcel", "scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

            }
            if (query == "") {
                var sUser, dep, sSer, sFromD, sToD, sPage, sStatus, sCountry
                sUser = $("#sUsers").val()
                sdep = $("#sDepartments").val()
                sSer = $("#sSeries").val()
                sFromD = $("#sPayFromDate").val()
                sToD = $("#sPayToDate").val()
                sPage = $("#pageList").val()
                sStatus = $("#sStatus").val()
                sCountry = $("#sCountries").val()
                window.open("reportScreen.aspx?" + "sdep=" + escape(sdep) + "&sUser=" + escape(sUser) + "&sSer=" + escape(sSer) + "&sFromD=" + escape(sFromD) + "&sToD=" + escape(sToD) + "&sPage=" + escape(sPage) + "&sStatus=" + escape(sStatus) + "&sCountry=" + escape(sCountry), "ReportExcel", "scrollbars=yes,menubar=yes, toolbar=yes, width=710,height=600,left=10,top=10");

            }


        }
        function FuncSort(query, par, srt) {

            var sUser, dep, sSer, sFromD, sToD, sPage, sStatus, sCountry
            sUser = $("#sUsers").val()
            sdep = $("#sDepartments").val()
            sSer = $("#sSeries").val()
            sFromD = $("#sPayFromDate").val()
            sToD = $("#sPayToDate").val()
            sPage = $("#pageList").val()
            sStatus = $("#sStatus").val()
            sCountry = $("#sCountries").val()
            if (query != "") {
                //alert(query)
                var r;
                r = "&" + par + "=ASC"
                query = query.replace(r, "");
                r = "&" + par + "=DESC"
                query = query.replace(r, "");
                //alert(query)

                window.location = "screenSetting.aspx?" + query + "&" + par + "=" + srt
            }
            else {
                window.location = "screenSetting.aspx?" + "sdep=" + escape(sdep) + "&sUser=" + escape(sUser) + "&sSer=" + escape(sSer) + "&sFromD=" + escape(sFromD) + "&sToD=" + escape(sToD) + "&sPage=" + escape(sPage) + "&sStatus=" + escape(sStatus) + "&sCountry=" + escape(sCountry) + "&" + par + "=" + srt
            }
        }
        function checkMoveTours() {

        }
        function checkDisable(depId) {
            //alert(depId)
            if (depId != '') {
                $.ajax({
                    type: "POST",
                    url: "DepartureChk.aspx",
                    data: depId,
                    data: { depId: depId },
                    success: function (msg) {
                        //alert(msg);




                    },
                    async: false

                });
            }
            //alert("tftrr")

        }
        function getX(obj) {
            return (obj.offsetParent == null ? obj.offsetLeft : obj.offsetLeft + getX(obj.offsetParent));
        }

        function getY(obj) {
            return (obj.offsetParent == null ? obj.offsetTop : obj.offsetTop + getY(obj.offsetParent));
        }
        function openStatus(departureId) {
            var sUser, dep, sSer, sFromD, sToD, sPage, sStatus, sCountry
            sUser = $("#sUsers").val()
            sdep = $("#sDepartments").val()
            sSer = $("#sSeries").val()
            sFromD = $("#sPayFromDate").val()
            sToD = $("#sPayToDate").val()
            sPage = $("#pageList").val()
            sStatus = $("#sStatus").val()
            sCountry = $("#sCountries").val()
            //alert(e.offsetY)
            //alert(e.clientX)
            obj = document.getElementById('sStatus')
            var leftC, topC
            leftC = getX(obj)
            topC = getY(obj) + 200


            newWin = window.open("editStatus.aspx?DepartureId=" + escape(departureId) + "&sdep=" + escape(sdep) + "&sUser=" + escape(sUser) + "&sSer=" + escape(sSer) + "&sFromD=" + escape(sFromD) + "&sToD=" + escape(sToD) + "&sPage=" + escape(sPage) + "&sStatus=" + escape(sStatus) + "&sCountry=" + escape(sCountry), "newWin", "toolbar=0,menubar=0,width=300,height=100,top=" + topC + ",left=" + leftC + ",scrollbars=auto");
            newWin.opener = window
        }

    </script>
</head>
<body>
    <form id="Form1" method="post" runat="server">
    <DS:LOGOTOP id="logotop"  runat="server"></DS:LOGOTOP>
	<DS:TOPIN id="topin" numOfLink="0"  numOfTab="71" toplevel2="73" runat="server"></DS:TOPIN>


    <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr>
            <td>
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td height="30">
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td valign="top" dir="rtl" width="100%">
                <table cellpadding="2" cellspacing="0" border="0" dir="rtl" id="Table3" align="left"
                    width="100%" bgcolor="#d8d8d8">
                    <tr bgcolor="#d8d8d8">
                        <td align="right" width="100%" bgcolor="#d8d8d8">
                            &nbsp;
                        </td>
                        <%If Request.Cookies("bizpegasus")("ScreenTourVisible") = "1" Then%>
                        <td align="left" nowrap>
                            <%End If%>
                            <td align="left" nowrap>
                                <a class="button_edit_1" style="width: 94;" href='javascript:void(0)' onclick="return openFile();">
                                    טיולים אשר יש לבחון העלאת מחיר</a>
                            </td>
                            <td align="left" nowrap>
                                <a class="button_edit_1" style="width: 94;" href='javascript:void(0)' onclick="return openReport();">
                                    הצג דוח Excel</a>
                            </td>
                            <td align="left" nowrap>
                                <a class="button_edit_1" style="width: 124;" href='javascript:void(0)' onclick="return openPdfReport();">
                                    הצג דוח PDF</a>
                            </td>
                            <%If Request.Cookies("bizpegasus")("ScreenSendMail") = "1" Then%>
                            <td align="left" nowrap>
                                <a class="button_edit_1" style="width: 94;" href='javascript:void(0)' onclick="return openSendReport();">
                                    שליחת מייל</a>
                            </td>
                            <%End If%>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table border="0" cellpadding="1" cellspacing="1" width="100%">
                    <tr bgcolor="#d8d8d8" style="height: 25px">
                        <td class="td_admin_5" align="center" width="50" nowrap>
                            <a href="#" onclick="javascript:ClearAll();" class="button_small1">הצג הכל</a>
                            <%If False Then%><asp:LinkButton runat="server" ID="btnSearchAll" Width="100%" CssClass="button_small1">הצג הכל</asp:LinkButton><%End If%>
                        </td>
                        <td class="td_admin_5" align="center">
                            <asp:LinkButton runat="server" ID="btnSearch" CssClass="button_small1">חפש</asp:LinkButton>
                        </td>
                        <%'IF ViewColumn(13)=true then%><td>
                        </td>
                        <%'end if%>
                        <%'IF ViewColumn(11)=true then%><td align="center">
                            <select runat="server" id="sUsers" class="searchList" dir="rtl" name="sUsers">
                            </select>
                        </td>
                        <%'end if%>
                        <%'IF ViewColumn(12)=true then%><td>
                        </td>
                        <%'end if%>
                        <%'IF ViewColumn(6)=true then%><td>
                        </td>
                        <%'end if%>
                        <%'IF ViewColumn(7)=true then%><td>
                        </td>
                        <%'end if%>
                        <%'IF ViewColumn(8)=true then%><td>
                        </td>
                        <%'end if%>
                        <%'IF ViewColumn(9)=true then%><td>
                        </td>
                        <%'end if%>
                        <%'IF ViewColumn(5)=true then%><td align="center">
                            <select runat="server" id="sDepartments" class="searchList" dir="rtl" name="sDepartments">
                            </select>
                        </td>
                        <%'end if%>
                        <%'IF ViewColumn(10)=true then%><td>
                        </td>
                        <%'end if%>
                        <%'IF ViewColumn(4)=true then%><td align="center">
                            <select runat="server" id="sCountries" class="searchList" dir="rtl" name="sCountries">
                            </select>
                        </td>
                        <%'end if%>
                        <%'IF ViewColumn(3)=true then%>
                        <td class="td_admin_5" align="center" nowrap>
                        <table cellspacing="0" cellpadding="0">
								<tr>
								<td class="td_admin_5" align="left" nowrap>
                            <a href="" onclick="cal1xx.select(document.getElementById('sPayFromDate'),'AsPayFromDate','dd/MM/yyyy'); return false;"
                                id="AsPayFromDate">
                                <img src="../../images/calendar.gif" border="0" align="center" class="calendar_icon"></a>
                            <input runat="server" type="text" id="sPayFromDate" class="searchList" style="width: 70px;font-size:9pt;font-weight:normal"
                                name="sPayFromDate">
                                                       </td>
								</tr>
								<tr>
								<td class="td_admin_5" align="left" nowrap>
                            <a href="" onclick="cal1xx.select(document.getElementById('sPayToDate'),'AsPayToDate','dd/MM/yyyy'); return false;"
                                id="AsPayToDate"><img src="../../images/calendar.gif" border="0" align="center"  class="calendar_icon"></a>
                            <input runat="server" type="text" id="sPayToDate" class="searchList" style="width: 70px;font-size:9pt;font-weight:normal"
                                name="sPayToDate">
                                                       </td>
								</tr>
                            </table>
                        </td>
                        <%'end if%>
                        <%'IF ViewColumn(2)=true then%><td class="td_admin_5" nowrap width="100">
                            &nbsp;
                            <div class="ui-widget">
                                <select runat="server" id="sSeries" dir="rtl" name="sSeries">
                                </select>
                            </div>
                            &nbsp;
                        </td>
                        <%'end if%>
                        <%'IF ViewColumn(1)=true then%><td align="center" class="td_admin_5">
                            <select runat="server" id="sStatus" style="font-size: 12px; width: 100px" dir="rtl"
                                name="sStatus">
                            </select>
                        </td>
                        <%'end if%>
                        <%If Request.Cookies("bizpegasus")("ScreenTourVisible") = "1" Then%>
                        <td align="center">
                            <asp:Button ID="btnSaveChecked" EnableViewState="false" runat="server" CssClass="button_edit_1"
                                Text="עדכן מסך צפייה"></asp:Button>
                        </td>
            </td>
            <%End If%>
        </tr>
        <tr>
            <td colspan="17" style="width: 100%; height: 2px; background-color: #C9C9C9">
            </td>
        </tr>
        <tr>
            <td class="title_sort">
                &nbsp;
            </td>
            <asp:Repeater ID="rptTitle" runat="server">
                <HeaderTemplate>
                </HeaderTemplate>
                <ItemTemplate>
                    <td class="title_sort" align="center">
                        <a href="javascript:FuncSort('<%=qrystring%>','sort_<%#Container.DataItem("Column_Id")%>','ASC')">
                            <img src="../../../images/arrow_top<%#IIf (Request.Querystring("sort_"& Container.DataItem("Column_Id"))="ASC","_act","")%>.gif"
                                title="למיין לפי <%#Container.DataItem("Column_Name")%>" border="0"></a><a href="javascript:FuncSort('<%=qrystring%>','sort_<%#Container.DataItem("Column_Id")%>','DESC')"><img
                                    src="../../../images/arrow_bot<%#IIf (Request.Querystring("sort_"& Container.DataItem("Column_Id"))="DESC","_act","")%>.gif"
                                    title="למיין לפי <%#Container.DataItem("Column_Name")%>" border="0"></a>&nbsp;<%#Container.DataItem("Column_Name")%>
                    </td>
                </ItemTemplate>
                <FooterTemplate>
                    <%If Request.Cookies("bizpegasus")("ScreenTourVisible") = "1" Then%><td class="title_sort"
                        align="center">
                        מופיע
                        <br>
                        במסך צפיה
                    </td>
                    <%End If%>
                </FooterTemplate>
            </asp:Repeater>
        </tr>
        <asp:Repeater ID="rptCustomers" runat="server">
            <ItemTemplate>
                <tr bgcolor="#C9C9C9" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';"
                    style="background-color: rgb(201, 201, 201);">
                    <td>
                        &nbsp;
                    </td>
                    <%'IF ViewColumn(13)=true then%><td align="center" class="price" <%#IIF(Container.dataitem("Price_Edit")=True,"style=background-color:#ff0000;color:#ffffff","")%>>
                        <%#Container.Dataitem("Price")%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(12)=true then%><td align="center" class="FlyingCompany">
                        <%#Container.DataItem("Flying_Company")%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(11)=true then%><td align="center" class="UserName">
                        <%#Container.DataItem("User_Name")%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(6)=true then%><td align="center">
                        <%#func.dbNullFix(Container.DataItem("CountlastMonth"))%><%'#func.GetCountLastMonth(Container.DataItem("Departure_Code"))%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(7)=true then%><td align="center">
                        <%#func.dbNullFix(Container.DataItem("Countlast2Week"))%><%'#func.GetCountLast2Week(Container.DataItem("Departure_Code"))%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(8)=true then%><td align="center">
                        <%#func.dbNullFix(Container.DataItem("CountlastWeek"))%><%'#func.GetCountLastDateSale(Container.DataItem("Departure_Code"))%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(9)=true then%><td align="center" nowrap>
                        <%#func.dbNullFix(Container.DataItem("LastDateSale"))%><%'#func.GetLastDateSale(Container.DataItem("Departure_Code"))%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(10)=true then%><td align="center">
                        <%#func.dbNullFix(Container.DataItem("CountMembersBitulim"))%><%'#func.GetCountMembersBitulim(Container.DataItem("Departure_Code"))%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(5)=true then%><td align="center">
                        <%#Container.DataItem("Dep_Name")%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(4)=true then%><td align="center">
                        <%#func.dbNullFix(Container.DataItem("CountMembers"))%><%'#func.GetCountMembers(Container.DataItem("Departure_Code"))%>
                    </td>
                    <%'end if%>
                    <td align="center">
                        <%#Container.Dataitem("CountryName")%>
                    </td>
                    <%'IF ViewColumn(3)=true then%><td align="center">
                        <%#Container.Dataitem("Date_Begin")%>
                    </td>
                    <%'end if%>
                    <td align="center">
                        <%#Container.DataItem("Series_Name")%>
                    </td>
                    <%'#IIF(ViewColumn(2),"<td align=center>" & Container.DataItem("Series_Name") &"</td>","")%>
                    <%'IF ViewColumn(1)=true then%>
                    <td align="right" style="padding-top: 1px; padding-bottom: 1px">
                        <table border="0" style="background-color: <%#Container.DataItem("Status_Color")%>;
                            width: 100%" cellpadding="1" cellspacing="1" align="right">
                            <tr>
                                <td align="right">
                                    <span style="color: #000000">&nbsp;<%#Container.DataItem("Status_Name")%>&nbsp;</span>
                                </td>
                                <td align="right">
                                    <%If Request.Cookies("bizpegasus")("ScreenTourStatus") = "1" Then%><a href="#" class="link_categ"
                                        onclick="openStatus('<%#Container.DataItem("Departure_Id")%>');return false"
                                        title="עדכון סטטוס" style="cursor: hand"><%#IIF(Container.DataItem("Status_Id")=5,"","<img src=""../../images/edit_icon.gif"" border=""0"" style=""margin-left:1px;"">")%></a><%End If%>
                                </td>
                            </tr>
                        </table>
                    </td>
                    <%'end if%>
                    <%If Request.Cookies("bizpegasus")("ScreenTourVisible") = "1" Then%>
                    <td align="center">
                        <input type="checkbox" id="chkTour" runat="server" name="chkTour">
                    </td>
                    <%End If%>
                    <%If False Then%>
                    <td class="td_admin_4" align="center">
                        <a class="table_fa_icon" href="" onclick="return moveTo('<%#Container.DataItem("Departure_Id")%>')">
                            <img src="../../images/arrow_top_bot.gif" align="top" border="0" alt="העבר למקום אחר"></a>
                    </td>
                    <td class="td_admin_4" align="center">
                        <input type="text" value="" onkeypress="return getNumbers(this)" class="Form" style="width: 30"
                            id="inpOrder<%#Container.DataItem("Departure_Id")%>">
                    </td>
                    <td class="td_admin_4" align="center">
                        <font color="#060165"><b>&nbsp;<%#Container.ItemIndex+1%>&nbsp;</b></font>
                    </td>
                    <%End If%>
                </tr>
            </ItemTemplate>
            <AlternatingItemTemplate>
                <tr bgcolor="#E6E6E6" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#E6E6E6';"
                    style="background-color: rgb(230, 230, 230);">
                    <td>
                        &nbsp;
                    </td>
                    <%'IF ViewColumn(13)=true then%><td align="center" <%#IIF(Container.dataitem("Price_Edit")=True,"style=background-color:#ff0000;color:#ffffff","")%>>
                        <%#Container.Dataitem("Price")%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(12)=true then%><td align="center">
                        <%#Container.DataItem("Flying_Company")%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(11)=true then%><td align="center">
                        <%#Container.DataItem("User_Name")%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(6)=true then%><td align="center">
                        <%#func.dbNullFix(Container.DataItem("CountlastMonth"))%><%'#func.GetCountLastMonth(Container.DataItem("Departure_Code"))%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(7)=true then%><td align="center">
                        <%#func.dbNullFix(Container.DataItem("Countlast2Week"))%><%'#func.GetCountLast2Week(Container.DataItem("Departure_Code"))%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(8)=true then%><td align="center">
                        <%#func.dbNullFix(Container.DataItem("CountlastWeek"))%><%'#func.GetCountLastDateSale(Container.DataItem("Departure_Code"))%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(9)=true then%><td align="center">
                        <%#func.dbNullFix(Container.DataItem("LastDateSale"))%><%'#func.GetLastDateSale(Container.DataItem("Departure_Code"))%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(10)=true then%><td align="center">
                        <%#func.dbNullFix(Container.DataItem("CountMembersBitulim"))%><%'#func.GetCountMembersBitulim(Container.DataItem("Departure_Code"))%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(5)=true then%><td align="center">
                        <%#Container.DataItem("Dep_Name")%>
                    </td>
                    <%'end if%>
                    <%'IF ViewColumn(4)=true then%><td align="center">
                        <%#func.dbNullFix(Container.DataItem("CountMembers"))%><%'#func.GetCountMembers(Container.DataItem("Departure_Code"))%>
                    </td>
                    <%'end if%>
                    <td align="center">
                        <%#Container.Dataitem("CountryName")%>
                    </td>
                    <%'IF ViewColumn(3)=true then%><td align="center">
                        <%#Container.Dataitem("Date_Begin")%>
                    </td>
                    <%'end if%>
                    <td align="center">
                        <%#Container.DataItem("Series_Name")%>
                    </td>
                    <%'#IIF(ViewColumn(2),"<td align=center>" & Container.DataItem("Series_Name") &"</td>","")%>
                    <%'IF ViewColumn(1)=true then%>
                    <td align="right" style="padding-top: 1px; padding-bottom: 1px">
                        <table border="0" style="background-color: <%#Container.DataItem("Status_Color")%>;
                            width: 100%" cellpadding="1" cellspacing="1" align="right">
                            <tr>
                                <td align="right">
                                    <span style="color: #000000">&nbsp;<%#Container.DataItem("Status_Name")%>&nbsp;</span>
                                </td>
                                <td align="right">
                                    <%If Request.Cookies("bizpegasus")("ScreenTourStatus") = "1" Then%>
                                    <a href="#" class="link_categ" onclick="openStatus('<%#Container.DataItem("Departure_Id")%>');return false"
                                        title="עדכון סטטוס" style="cursor: hand">
                                        <%#IIF(Container.DataItem("Status_Id")=5,"","<img src=""../../images/edit_icon.gif"" border=""0"" style=""margin-left:1px;"">")%></a><%End If%>
                                </td>
                            </tr>
                        </table>
                    </td>
                    <%'end if%>
                    <%If Request.Cookies("bizpegasus")("ScreenTourVisible") = "1" Then%>
                    <td align="center">
                        <input type="checkbox" id="chkTour" runat="server" name="chkTour">
                    </td>
                    <%End If%>
                    <%If False Then%>
                    <td class="td_admin_4" align="center">
                        <a class="table_fa_icon" href="" onclick="return moveTo('<%#Container.DataItem("Departure_Id")%>')">
                            <img src="../../images/arrow_top_bot.gif" align="top" border="0" alt="העבר למקום אחר"></a>
                    </td>
                    <td class="td_admin_4" align="center">
                        <input type="text" value="" onkeypress="return getNumbers(this)" class="Form" style="width: 30"
                            id="inpOrder<%#Container.DataItem("Departure_Id")%>">
                    </td>
                    <td class="td_admin_4" align="center">
                        <font color="#060165"><b>&nbsp;<%#Container.ItemIndex+1%>&nbsp;</b></font>
                    </td>
                    <%End If%>
                </tr>
            </AlternatingItemTemplate>
        </asp:Repeater>
        <!--  <div id="dvCustomers">-->
        <!--paging-->
        <tr>
            <td colspan="15" style="height: 20px; bgcolor: #ffffff">
                &nbsp;
            </td>
            <%If Request.Cookies("bizpegasus")("ScreenTourVisible") = "1" Then%><td style="height: 20px;
                bgcolor: #ffffff">
                &nbsp;
            </td>
            <%End If%>
        </tr>
        <tr>
            <td class="title_sort" style="background-color: #E8A98A">
                &nbsp;
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;מכירות בחודש האחרון
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;מכירות בשבועיים האחרונים
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;מכירות בשבוע האחרון
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;מטיילים שביטלו רישומם
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;רשומים עכשווית
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;
            </td>
            <td class="title_sort" style="background-color: #E8A98A" align="center">
                &nbsp;
            </td>
            <%If Request.Cookies("bizpegasus")("ScreenTourVisible") = "1" Then%><td class="title_sort"
                style="background-color: #E8A98A" align="center">
                &nbsp;
            </td>
            <%End If%>
        </tr>
        <tr>
            <td class="title_sort">
                &nbsp;
            </td>
            <td class="title_sort" align="center">
                &nbsp;
            </td>
            <td class="title_sort" align="center">
                &nbsp;
            </td>
            <td class="title_sort" align="center">
                &nbsp;
            </td>
            <td class="title_sort" align="center">
                <%=AllCountlastMonth%>
            </td>
            <td class="title_sort" align="center">
                <%=AllCountlast2Week%>
            </td>
            <td class="title_sort" align="center">
                <%=AllCountlastWeek%>
            </td>
            <td class="title_sort" align="center">
                &nbsp;
            </td>
            <td class="title_sort" align="center">
                <%=AllCountMembersBitulim%>
            </td>
            <td class="title_sort" align="center">
                &nbsp;
            </td>
            <td class="title_sort" align="center">
                <%=AllCountMembers%>
            </td>
            <td class="title_sort" align="center">
                &nbsp;
            </td>
            <td class="title_sort" align="center">
                &nbsp;
            </td>
            <td class="title_sort" align="center">
                &nbsp;
            </td>
            <td class="title_sort" align="center">
                &nbsp;
            </td>
            <%If Request.Cookies("bizpegasus")("ScreenTourVisible") = "1" Then%><td class="title_sort"
                align="center">
                &nbsp;
            </td>
            <%End If%>
        </tr>
        <asp:PlaceHolder ID="pnlPages" runat="server">
            <tr>
                <td height="2" colspan="13">
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
                                    <asp:ListItem Value="20">20</asp:ListItem>
                                    <asp:ListItem Value="50" Selected="True">50</asp:ListItem>
                                </asp:DropDownList>
                                &nbsp;בדף&nbsp;
                            </td>
                            <td valign="baseline" nowrap align="right">
                                <asp:LinkButton ID="cmdNext" runat="server" CssClass="page_link" Text="«עמוד הבא"></asp:LinkButton>
                            </td>
                            <td class="plata_paging" valign="baseline" nowrap align="center" width="160">
                                <asp:Label ID="lblTotalPages" runat="server"></asp:Label>&nbsp;דף&nbsp;
                                <asp:DropDownList ID="pageList" CssClass="PageSelect" runat="server" AutoPostBack="true">
                                </asp:DropDownList>
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
    </td> </tr></table>
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
</body>
</html>

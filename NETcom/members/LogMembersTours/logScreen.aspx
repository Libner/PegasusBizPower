<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="logScreen.aspx.vb" Inherits="bizpower_pegasus2018.LogMembersToursScreen" %>
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
            padding: 2px 1em 2px .1em;
        }
        
        .ui-widget.ui-widget-content
        {
            border: 1px solid #c5c5c5;
            background: #ffffff;
            max-height: 250px;
            /*width: 58px;*/
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
          .css('width', $(this.element).width())
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
            $("#sCountries").sSeries();
            $("#toggle").on("click", function () {
                $("#sCountries").toggle();
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
                if($(this).val()!="")
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

           
            if (query != "") {
                //alert(query)
                var r;
                r = "&" + par + "=ASC"
                query = query.replace(r, "");
                r = "&" + par + "=DESC"
                query = query.replace(r, "");
                //alert(query)

//alert("logScreen.aspx?" + query + "&" + par + "=" + srt)
                window.location = "logScreen.aspx?" + query + "&" + par + "=" + srt
            }
            else {
				var querySearch
				querySearch=""
				
				$("#blockSearchTerms").find("input").each(function () {
                if($(this).val()!="")
					{
					if(querySearch!="") {querySearch=querySearch + "&"}			
					querySearch=querySearch + $(this).attr("name") + "=" + encodeURIComponent($(this).val())
					}
				});
				$("#blockSearchTerms").find("select").each(function () {
                if($(this).val()!="")
					{
					if(querySearch!="") {querySearch=querySearch + "&"}			
					querySearch=querySearch + $(this).attr("name") + "=" + encodeURIComponent($(this).val())
					}
				});
//alert("logScreen.aspx?" + querySearch + "&" + par + "=" + srt)
                window.location = "logScreen.aspx?" + querySearch + "&" + par + "=" + srt
            }
        }
        function checkMoveTours() {

        }
        function  SendAppleal(Log_Id) {
            //alert(depId)
            if (Log_Id != '') {
                $.ajax({
                    type: "POST",
                    url: "sendAppeal.aspx",
                    data: { logId: Log_Id},
                    success: function (msg) {
						var res=''
						if (msg != '' && !isNaN(msg) )
						{
							if(eval(msg)>0)
							{
								$("#blockExistsAppleal16504_" + Log_Id).css("background-color","#99ff33")
								$("#imgExistsAppleal16504_" + Log_Id).src="../../images/v.png"
								$("#chkSendAppleal_" + Log_Id).prop("checked",false)
								$("#chkSendAppleal_" + Log_Id).parent().css("background-color","")
								$("#chkSendAppleal_" + Log_Id).prop("checked",false)
							}
							else
							{
								alert('פעולה לא בוצע: ' + msg)
							}
						}
						else
						{
							alert('פעולה לא בוצע: ' + msg)
						}
                    },
                    async: false

                });
            }
        }
        
function SendApplealToAllMarked()
{

			var selected = new Array();
 
            $("[name='chkSendAppleal']:checked").each(function () {
                selected.push(this.value);
            });
 //alert(selected.length)
            if (selected.length > 0) {
				Logs=selected.join(",")
				
// alert(Logs)
                $.ajax({
                    type: "POST",
                    url: "sendAppeal.aspx",
                    data: { Logs: Logs},
                    success: function (msg) {
						if (msg != 'ERROR'  )
						{	
// alert(msg)
								var res=new Array
								res=msg.split(',')
								if (res.length>0)
								{
									for(i=0; i<res.length; i++)
									{
										Log_Id=res[i]
										$("#blockExistsAppleal16504_" + Log_Id).css("background-color","#99ff33")
											$("#imgExistsAppleal16504_" + Log_Id).attr("src","../../images/v.png")
										//$("#chkSendAppleal_" + Log_Id).prop("checked",false)
										//$("#chkSendAppleal_" + Log_Id).parent().css("background-color","")
										$("#chkSendAppleal_" + Log_Id).prop("disabled",true)
									}
								}
								else
								{
									alert('טפסי מתעניין קיימים לכל היעדים המסומנים')
								}
								$("#chkAllSendAppleal").prop("checked",false)
						}
						else
						{
							alert('פעולה לא בוצע: ' + msg)
						}
                    },
                    async: false

                });
                $("[name='chkSendAppleal']:checked").each(function () {
					$(this).prop("checked",false)
					$(this).parent().css("background-color","")
            });
            }
}
 
function chkToSendAppleal(Log_Id)
{
		if($("#chkSendAppleal_" + Log_Id).prop("checked"))
		{
			$("#chkSendAppleal_" + Log_Id).parent().css("background-color","#ffff99")
		}
		else
		{
			$("#chkSendAppleal_" + Log_Id).parent().css("background-color","")
		}
}      
function checkAllToSendAppleal()
{
	$("[name='chkSendAppleal']").each(function()
	{
		$(this).prop("checked",$("#chkAllSendAppleal").prop("checked"))
		if($(this).prop("checked"))
		{
			$(this).parent().css("background-color","#ffff99")
		}
		else
		{
			$(this).parent().css("background-color","")
		}
	})
	
}

function openConcactLog(CONTACT_Id)
{
            
             h = 700;
             w = 800;
             S_Wind = window.open("ConcactLog.aspx?contactId=" + CONTACT_Id, "S_Wind", "scrollbars=1,toolbar=0,top=50,left=50,width=" + w + ",height=" + h + ",align=center,resizable=1");
             S_Wind.focus();
             return false;
}


function openWinFilterByContact()
{
	var searchParams=new String
	$('#blockSearchTerms input, #blockSearchTerms select').each(
		function(){  
			var input = $(this);
			if (input.val() !='')
			{
			queryStrItem=input.attr('id')+'='+input.val()
			searchParams=(searchParams=='')? queryStrItem : searchParams+'&'+queryStrItem
			}
		}
	);

    searchParams=searchParams+'<%=sortQrystring%>'
    //console.log(searchParams)
             h = 700;
             w = 800;
             S_Wind = window.open("logScreenTotalByMembers.aspx?" + searchParams, "S_Wind", "scrollbars=1,toolbar=0,top=50,left=50,width=" + w + ",height=" + h + ",align=center,resizable=1");
             S_Wind.focus();
             return false;
}
    

function openWinLastEntersByContact()
{
	var searchParams=new String
	$('#blockSearchTerms input, #blockSearchTerms select').each(
		function(){  
			var input = $(this);
			if (input.val() !='')
			{
			queryStrItem=input.attr('id')+'='+input.val()
			searchParams=(searchParams=='')? queryStrItem : searchParams+'&'+queryStrItem
			}
		}
	);

    searchParams=searchParams+'<%=sortQrystring%>'
    //console.log(searchParams)
             h = $(window).height()-100;
             w = $(window).width()-100;
             S_Wind = window.open("logScreenMembersLastEnter.aspx?" + searchParams, "S_Wind", "scrollbars=1,toolbar=0,top=50,left=50,width=" + w + ",height=" + h + ",align=center,resizable=1");
             S_Wind.focus();
             return false;
}
    </script>
</head>
<body>
<form id="Form1"  method="post" runat="server" autocomplete=off>	
	<DS:LOGOTOP id="logotop"  runat="server"></DS:LOGOTOP>
	<DS:TOPIN id="topin" numOfLink="16"  numOfTab="108" topLevel2="131" runat="server"></DS:TOPIN>
	
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
            <td valign="top" dir="rtl" width="100%" >
                <table cellpadding="2" cellspacing="0" border="0" dir="rtl" id="Table3" align="left"
                    width="100%" bgcolor="#d8d8d8">
                    <tr bgcolor="#d8d8d8">
                      <td align="right" nowrap>
                              <a class="button_edit_1"  href='javascript:void(0)' onclick="return openWinLastEntersByContact();" style="WIDTH: 194px; COLOR: #000000; BACKGROUND-COLOR: #FF9900">
                                    תצוגה לאחר סינון המסך</a>
                            </td>
                    <%if 0 then%>
                      <td align="right" nowrap>
                              <a class="button_edit_1"  href='javascript:void(0)' onclick="return openWinFilterByContact();" style="WIDTH: 194px; COLOR: #000000; BACKGROUND-COLOR: #ccffff">
                                    סך כל כניסות לפי לקוח</a>
                            </td>
                        <%end if%>
                            
                         <td align="right" width="100%" bgcolor="#d8d8d8">
                            &nbsp;
                        </td>
                            <%if AddAppealPermitted  then%>
                             <td align="left" nowrap>
                              <a class="button_edit_1"  href='javascript:void(0)' onclick="return SendApplealToAllMarked();" style="width: 194;BACKGROUND-COLOR: #99cc00;color:#000000">
                                    ליצור טפסי מתעניין בטיול</a>
                            </td>
                            <%end if%>
                            <td align="left" nowrap>
                                <a class="button_edit_1" style="width: 94;" href='javascript:void(0)' onclick="return openReport();">
                                    הצג דוח Excel</a>
                            </td>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table border="0" cellpadding="1" cellspacing="1" width="100%" dir="rtl">
                    <tr bgcolor="#d8d8d8"  id="blockSearchTerms">
                    <td></td>
                    <td></td>
                    <td></td>
                        <td class="td_admin_5" nowrap>
                                <input runat="server" id="sContactName" dir="rtl" name="sContactName" class="search">
                        </td>
                        <td class="td_admin_5" nowrap>
                                <input runat="server" id="sPhone" dir="rtl" name="sPhone" class="search">
                        </td>
                        
                        <td class="td_admin_5" nowrap >
                                <input runat="server" id="sEmail" dir="rtl" name="sEmail" class="search">
                        </td>
                       
                        <td class="td_admin_5" >
                                  <select runat="server" id="sSubCategoriesTours" dir="rtl" name="sSubCategoriesTours" class="search" style="height:21px;width:200px">
                                </select>
                        </td>
                        
                        <td class="td_admin_5" nowrap  style="width: 250px">
                            &nbsp;
                            <div class="ui-widget" style="width: 100%">
                                <select runat="server" id="sCountries" dir="rtl" name="sCountries">
                                </select>
                            </div>
                            &nbsp;
                        </td>
                        <td class="td_admin_5" nowrap width="100">
                            &nbsp;
                            <div class="ui-widget">
                                <select runat="server" id="sSeries" dir="rtl" name="sSeries">
                                </select>
                            </div>
                            &nbsp;
                        </td>
                        
                        <td class="td_admin_5" align="center" nowrap dir="ltr">
							<table cellspacing="0" cellpadding="0">
								<tr>
								<td class="td_admin_5" align="left">
								<a href="" onclick="cal1xx.select(document.getElementById('sLastEnterFromDate'),'AsLastEnterFromDate','dd/MM/yyyy'); return false;"
                                id="AsLastEnterFromDate" title="תאריך אחרון בו נכנס הלקוח לעמוד הרלוונטי של הסדרה/יעד">
                                <img src="../../images/calendar.gif" border="0" align="center"></a>
                            <input runat="server" type="text" id="sLastEnterFromDate" class="search" style="width: 70px"  name="sLastEnterFromDate" title="תאריך אחרון בו נכנס הלקוח לעמוד הרלוונטי של הסדרה/יעד">
                            מי
								</td>
								</tr>
								<tr>
								<td class="td_admin_5" align="left">
								<a href="" onclick="cal1xx.select(document.getElementById('sLastEnterToDate'),'AsLastEnterToDate','dd/MM/yyyy'); return false;"
                                id="AsLastEnterToDate" title="תאריך אחרון בו נכנס הלקוח לעמוד הרלוונטי של הסדרה/יעד">
                                <img src="../../images/calendar.gif" border="0" align="center"></a>
                            <input runat="server" type="text" id="sLastEnterToDate" class="search" style="width: 70px" name="sLastEnterToDate" title="תאריך אחרון בו נכנס הלקוח לעמוד הרלוונטי של הסדרה/יעד">
                            עד
								</td>
								</tr>
                            </table>
                        </td>
                         <td class="td_admin_5" nowrap width="100">
                                <input runat="server" id="sCountEnters" dir="rtl" name="sCountEnters" class="search">
                        </td>
                         <td class="td_admin_5" nowrap width="100">
                                <select runat="server" id="sExistsAppleal16735" dir="rtl" name="sExistsAppleal16735" class="search">
                                   <option value="">הצג הכל</option>
                                 <option value="1">כן</option>
                                 <option value="0">לא</option>
                                </select>
                        </td>
                        <td class="td_admin_5" nowrap width="100">
                                <select runat="server" id="sExistsAppleal16504" dir="rtl" name="sExistsAppleal16504" class="search">
                                <option value="">הצג הכל</option>
                                 <option value="1">כן</option>
                                 <option value="0" >לא</option>
                                </select>
                        </td>
                            <td class="td_admin_5" align="center" width="50" nowrap>
                            <asp:LinkButton runat="server" ID="btnSearch" CssClass="button_small1"  BackColor=#ff9900>חפש</asp:LinkButton>
                        </td>
                      
                        <td class="td_admin_5" align="center" width="50" nowrap>
                            <a href="logScreen.aspx" class="button_small1">הצג הכל</a>
                         </td>
                      
                      </tr>
                           <tr>
                                  
                                    <td class="td_admin_5" align="center" nowrap>
                                        סדר
                                    </td>
                                    
                                     <td class="title_sort" align="center">
                                     </td>
                                     <td class="title_sort" align="center">
                                     </td>
            <asp:Repeater ID="rptTitle" runat="server">
            
                <ItemTemplate>
                    <td class="title_sort" align="center">
                    <asp:PlaceHolder ID="phSortArrows" Runat=server Visible=False>
                        <a href="javascript:FuncSort('<%=qrystring%>','sort_<%#Container.DataItem("Column_Id")%>','ASC')">
                            <img src="../../../images/arrow_top<%#IIf (Request.Querystring("sort_"& Container.DataItem("Column_Id"))="ASC","_act","")%>.gif"
                                title="למיין לפי <%#Container.DataItem("Column_Title")%>" border="0"></a><a href="javascript:FuncSort('<%=qrystring%>','sort_<%#Container.DataItem("Column_Id")%>','DESC')"><img
                                    src="../../../images/arrow_bot<%#IIf (Request.Querystring("sort_"& Container.DataItem("Column_Id"))="DESC","_act","")%>.gif"
                                    title="למיין לפי <%#Container.DataItem("Column_Title")%>" border="0"></a>&nbsp;
                                    </asp:PlaceHolder>
                                    <%#Container.DataItem("Column_Title")%>
                    </td>
                </ItemTemplate>
            </asp:Repeater>
               <td colspan="2" bgcolor=#ffffcc align="center">
              <%if AddAppealPermitted  then%>
                      <input type="checkbox" id="chkAllSendAppleal"  name="chkAllSendAppleal" value="1" onclick="checkAllToSendAppleal()" title="סמנו כל השורות לייצירת טופס מתעניין">&nbsp;<b>סמן הכל</b> 
<%end if%></td>
        </tr>
                    
 <asp:Repeater ID="rptLogs" runat="server">
            <ItemTemplate>
                <tr bgcolor="#C9C9C9" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#C9C9C9';"
                    style="background-color: rgb(201, 201, 201);">
                   
                    <td class="td_admin_4" align="center">
                        <font color="#060165"><b>&nbsp;<%#Container.ItemIndex+1%>&nbsp;</b></font>
                    </td>
                     <td  class="td_admin_4" align="center">
                      <a href="" onclick="openConcactLog('<%#func.dbNullFix(Container.DataItem("CONTACT_Id"))%>');return false"  title="היסטוריית כניסות של הלקוח"><img src="../../images/popup.gif" border="0" hspace="3" alt="לכרטיס הלקוח"></a>
                     </td>
                    <td  class="td_admin_4" align="center">
                      <a href="" onclick="window.open('../appeals/contact_appeals.asp?contactID=<%#func.dbNullFix(Container.DataItem("CONTACT_Id"))%>','winCA','top=20, left=10, width=950, height=500, scrollbars=1');return false" title="היסטוריית טפסים של הלקוח"><img src="../../images/forms_icon.gif" border="0" hspace="3" alt="לכרטיס הלקוח"></a>
                     </td>
                    <td align="right" class="td_admin_4">
                        <%#func.dbNullFix(Container.DataItem("CONTACT_NAME"))%>
                    </td>
                  
                    <td align="right" class="td_admin_4">
                        <%#func.dbNullFix(Container.DataItem("Phone"))%>
                    </td>
                    
                   <td align="right" class="td_admin_4">
                        <%#func.dbNullFix(Container.DataItem("Email"))%>
                    </td>
                    
                    <td align="right" class="td_admin_4">
                        <%#func.dbNullFix(Container.DataItem("SubCategory_Name"))%>
                    </td>
                    <td align="right" class="td_admin_4">
                        <%#func.dbNullFix(Container.DataItem("Country_Name"))%>
                    </td>
                    <td align="center" class="td_admin_4">
                        <%#func.dbNullFix(Container.DataItem("Series_Name"))%>
                    </td>
                    
                        
                     <td align="center" class="td_admin_4">
                        <%#func.dbNullDate(Container.DataItem("LastEnter_Date"),"dd/MM/yyyy HH:mm")%>
                    </td>
                    <td align="center" class="td_admin_4">
                        <%#func.fixnumeric(Container.DataItem("countEnters"))%>
                    </td>
                     <td align="center" class="td_admin_4">
						  <asp:PlaceHolder ID="phAppleal16735" Runat=server Visible=False>
                         <img src="../../images/<%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16735")),"v.png","notok_icon.gif")%>" border="0" height="20">
					</asp:PlaceHolder>
                    </td>
                   <td align="center" class="td_admin_4" id="blockExistsAppleal16504_<%#Container.DataItem("Log_Id")%>" <%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16504")),"style=background-color:#ccffcc","")%>>
                          <asp:PlaceHolder ID="phAppleal16504" Runat=server Visible=False>                        
                         <img id="imgExistsAppleal16504_<%#Container.DataItem("Log_Id")%>" src="../../images/<%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16504")),"v.png","notok_icon.gif")%>" border="0" height="20">
						 </asp:PlaceHolder>
                    </td>
                   <td align="center"  colspan="2" class="td_admin_4">
                   <asp:PlaceHolder ID="phChkSendAppleal" Runat=server Visible=False>
                         <input type="checkbox" id="chkSendAppleal_<%#Container.DataItem("Log_Id")%>"  name="chkSendAppleal" value="<%#Container.DataItem("Log_Id")%>" onclick="chkToSendAppleal('<%#Container.DataItem("Log_Id")%>')">
                </asp:PlaceHolder>
                  </td>
                </tr>
            </ItemTemplate>
            <AlternatingItemTemplate>
                <tr bgcolor="#E6E6E6" onmouseover="this.style.backgroundColor='#B1B0CF';" onmouseout="this.style.backgroundColor='#E6E6E6';"
                    style="background-color: rgb(230, 230, 230);">
                    
                  <td class="td_admin_4" align="center">
                        <font color="#060165"><b>&nbsp;<%#Container.ItemIndex+1%>&nbsp;</b></font>
                    </td>
         <td  class="td_admin_4" align="center">
                      <a href="" onclick="openConcactLog('<%#func.dbNullFix(Container.DataItem("CONTACT_Id"))%>');return false"  title="היסטוריית כניסות של הלקוח"><img src="../../images/popup.gif" border="0" hspace="3" alt="לכרטיס הלקוח"></a>
                     </td>
                    <td  class="td_admin_4" align="center">
                      <a href=""  onclick="window.open('../appeals/contact_appeals.asp?contactID=<%#func.dbNullFix(Container.DataItem("CONTACT_Id"))%>','winCA','top=20, left=10, width=950, height=500, scrollbars=1');return false" title="היסטוריית טפסים של הלקוח"><img src="../../images/forms_icon.gif" border="0" hspace="3" alt="לכרטיס הלקוח"></a>
                     </td>
                    
                    <td align="right" class="td_admin_4">
                        <%#func.dbNullFix(Container.DataItem("CONTACT_NAME"))%>
                    </td>
                    <td align="right" class="td_admin_4">
                        <%#func.dbNullFix(Container.DataItem("Phone"))%>
                    </td>
                    
                   <td align="right" class="td_admin_4">
                        <%#func.dbNullFix(Container.DataItem("Email"))%>
                    </td>
                    
                    <td align="right" class="td_admin_4">
                        <%#func.dbNullFix(Container.DataItem("SubCategory_Name"))%>
                    </td>
                    <td align="right" class="td_admin_4">
                        <%#func.dbNullFix(Container.DataItem("Country_Name"))%>
                    </td>
                    <td align="center" class="td_admin_4">
                        <%#func.dbNullFix(Container.DataItem("Series_Name"))%>
                    </td>
                    
                        
                     <td align="center" class="td_admin_4">
                        <%#func.dbNullDate(Container.DataItem("LastEnter_Date"),"dd/MM/yyyy HH:mm")%>
                    </td>
                    <td align="center" class="td_admin_4">
                        <%#func.fixnumeric(Container.DataItem("countEnters"))%>
                    </td>
                       <td align="center" class="td_admin_4">
                        <asp:PlaceHolder ID="phAppleal16735" Runat=server Visible=False>
<img src="../../images/<%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16735")),"v.png","notok_icon.gif")%>" border="0" height="20">
                   </asp:PlaceHolder>
                    </td>
                   <td align="center" class="td_admin_4" id="blockExistsAppleal16504_<%#Container.DataItem("Log_Id")%>" <%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16504")),"style=background-color:#ccffcc","")%>>
                        <asp:PlaceHolder ID="phAppleal16504" Runat=server Visible=False>
                           <img id="imgExistsAppleal16504_<%#Container.DataItem("Log_Id")%>" src="../../images/<%#iif(func.dbnullBool(Container.Dataitem("ExistsAppleal16504")),"v.png","notok_icon.gif")%>" border="0" height="20">
                        </asp:PlaceHolder>
                    </td> 
                   <td align="center" colspan="2" class="td_admin_4">
                      <asp:PlaceHolder ID="phChkSendAppleal" Runat=server Visible=False>
                       <input type="checkbox" id="chkSendAppleal_<%#Container.DataItem("Log_Id")%>"  name="chkSendAppleal" value="<%#Container.DataItem("Log_Id")%>" onclick="chkToSendAppleal('<%#Container.DataItem("Log_Id")%>')">
                </asp:PlaceHolder>
                </td>
                </tr>
            </AlternatingItemTemplate>
        </asp:Repeater>
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
    </td>
    </tr>
    </table>
   
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
</body>
</html>

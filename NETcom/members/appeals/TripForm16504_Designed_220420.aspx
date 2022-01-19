<%@ Page Language="vb" AutoEventWireup="false" EnableViewState="true" Codebehind="TripForm16504_Designed_220420.aspx.vb" Inherits="bizpower_pegasus2018.TripForm16504"  validateRequest="false"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<title>טופס מתעניין בטיול</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
		<link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
		<script src="//code.jquery.com/ui/1.11.4/jquery-ui.min.js"></script>
		<script language="javascript" type="text/javascript" src="../../../Scripts/CalendarPopupH.js"
			charset="windows-1255"></script>
		<script language="javascript" type="text/javascript" charset="windows-1255">
	var cal = new CalendarPopup();
		</script>
		<script>
function refreshContactsPr(contactObj,quest_id)
{
	if ((event.keyCode==13) || (event.keyCode == 9))
		return false;
	contName = new String(contactObj.value);
	if(contName.length > 0)
	{
		window.document.getElementById("frameContactsPr").src = "contacts_privateQ16504.asp?cont_name=" + contName + "&quest_id=" + quest_id;	
		window.document.getElementById("frameContactsPr").style.display = "inline";
	}
	else
		window.document.getElementById("frameContactsPr").style.display = "none";
			
}

		</script>
		<script>

//alert("ready")
$('#Form1').submit( function( event ) 
{
alert("gf0")

        $("#Form1").submit();
});
 
   
function getNumbers(){
			var ch=event.keyCode;
			event.returnValue =(ch >= 48 && ch <= 57);
		}




		</script>
		<!--<script language="javascript" type="text/javascript" src="../../../Scripts/CalendarPopupH.js"></script>-->
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
             margin-top: 3px  !important;  
    margin-bottom: 10px  !important;
    padding-right: 10px;
    padding-left: 10px;
    padding-bottom: 5px;
    padding-top: 5px;
   background-color: #f7f7f7 !important;
    font-size: 15px !important;
    text-align: right;
    width: 468px;
    line-height: 1.428571429;
    color: #333333;
    vertical-align: middle;
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
            font-size: 15px;
            direction: rtl;
        }
        
        .ui-menu .ui-menu-item-wrapper
        {
            position: relative;
            font-size: 15px;
            font-weight:normal;
            padding: 2px 1em 2px .1em;
        }
        .ui-widget.ui-widget-content
        {
            border: 1px solid #c5c5c5;
            background: #ffffff;
            max-height: 250px;
          /*  width: 460px;*/
            overflow-y: scroll;
            overflow-x: hidden;
            font-size: 13px;
        }
       .ui-menu-item-wrapper:hover
        {
   background-color: #cccccc !important;
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
                    .attr("required" ,"required")
                    .attr('id', 'input_' + $(this.element).get(0).id)
                    .attr('name', 'input_' + $(this.element).get(0).id)
                    //.css('width', $(this.element).width())
                    .addClass("custom-sSeries-input ui-widget ui-widget-content ui-state-default ui-corner-right")
                    .autocomplete({
                        delay: 0,
                        minLength: 0,
                        source: $.proxy(this, "_source"),
                        async: true,
                        position: { my: "right top", at: "right bottom" }
                    })
                    .tooltip({
                        classes: {
                            "ui-tooltip": "ui-state-highlight"
                        }
                    })

                    .on("click", function () {
                        $(this).val('');
                        
                    $(".ui-menu").css("max-width", $(".custom-tour_order_field").css("width"));
                    $(".ui-menu").css("width", $(".custom-tour_order_field").css("width"));
                    });
               
                    this._on( this.input,
                        {
                            autocompleteselect: function (event, ui) {
                                ui.item.option.selected = true;
                                this._trigger("select", event,
                                    {
                                        item: ui.item.option
                                    });
                                var getwid = this.element.width();
                                //$(event.target).width(getwid);


                                if ($(event.target).attr("id") == 'input_CRMCountry') {
                                $("#input_selSerias").val('')
                                    //$( "#tourTitle" ).text(ui.item.label)
                                 //   selectedTourId = $("#selSRSCODE").val()
                                       $("#selSerias").html(fillSelByCRMCountry($("#CRMCountry").val(), $("#selSerias").val()))
                    
                                }
                            }
                       
                    });
                },

                _createShowAllButton: function () {
                    var input = this.input,
                    wasOpen = false;

                    $("<a>")
                      .attr("tabIndex", -1)
                      .attr("title", "All")
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
                      }
                   );
                },
                _source: function (request, response) {
                    //var matcher = new RegExp($.ui.autocomplete.escapeRegex(request.term), "i");
                     var matcher = new RegExp("^" + $.ui.autocomplete.escapeRegex(request.term), "i");
					response(this.element.children("option").map(function () {
                        var text = $(this).text();
                        if (this.value && (!request.term || matcher.test(text)))
                            return {
                                label: text,
                                value: text,
                                option: this
                            };
                    })
                );
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
                //.val("")
                //.attr("title", value + " didn't match any item")
                .tooltip("open");
                    this.element.val("");
                    this._delay(function () {
                        this.input.tooltip("close").attr("title", "");
                    }, 2500);
                    // this.input.autocomplete("instance").term = "";
                },


                _destroy: function () {
                    this.wrapper.remove();
                    this.element.show();
                }
            });
            $("#CRMCountry").sSeries();
            $("#toggle").on("click", function () {
                $("#CRMCountry").toggle();
            });
        
            $("#selSerias").sSeries();
            $("#toggle").on("click", function () {
                $("#selSerias").toggle();
            });
       
      

        });
              function fillSelByCRMCountry(COUNTRYID, SERID)
        {
    
         htmlOptionsList=""
            $.ajax({
                type: "POST",
                url: "getSeriasByCountry.aspx",
                data: { COUNTRYID: COUNTRYID, SERID: SERID },
                success: function(msg) {

                    if(msg != "")
                    {
            //      alert(msg)
                        htmlOptionsList=msg
                    }
                },
                async:   false
            });
            return htmlOptionsList
        }
		</script>
		<style>
  body
  {
  BACKGROUND-COLOR:#ffffff;
  font-size:15px !important
  }
   td
  {
  font-size:15px !important;
  padding-top: 4px  !important;
  padding-bottom: 4px  !important;
  padding-right: 0px  !important;
  padding-left: 0px  !important;
  }
  
  td.form
  {
  font-size:15px !important;
  color: #005790 !important;
  padding-top: 15px  !important
  }
  .formfield_title
  {
  font-size:15px !important
  }
  .input_form
  {
  margin-top: 3px  !important;  
    margin-bottom: 10px  !important;
    padding-right: 10px;
    padding-left: 10px;
    padding-bottom: 5px;
    padding-top: 5px;
   background-color: #f7f7f7;
    font-size: 15px !important;
    text-align: right;
    width: 100%;
    line-height: 1.428571429;
    color: #333333;
    vertical-align: middle;
    border: 1px solid #cccccc
  } 
  .textarea_form
  {
  margin-top: 3px  !important;  
    margin-bottom: 10px  !important;
    padding-right: 10px;
    padding-left: 10px;
    background-color: #f7f7f7;
    font-size: 16px;
    text-align: right;
    width: 100%;
    padding: 8px 12px;
    font-size: 14px;
    line-height: 1.428571429;
    color: #333333;
    vertical-align: middle;
    border: 1px solid #cccccc
  }
  .title_form
  {
	BACKGROUND-COLOR: #005790 !important;
  font-size:18px !important;
  }
  .subtitle_form
  {
	BACKGROUND-COLOR: #4e89c6 !important;
  font-size:18px !important;
  COLOR:#ffffff !important;
  padding-right:10px !important
  }
  .form_submit
  {   
    width: 100px;
    height:40px;
    font-size:15px;
    font-weight:bold;
    COLOR:#ffffff !important;
  border-radius: 3px !important;
	BACKGROUND-COLOR: #005790 !important;
  }
  .form_submit:hover
  {  
	BACKGROUND-COLOR: #4e89c6 !important;
  }
		</style>
	</head>
	<body>
		<form id="Form1" method="post" runat="server">
			<input type="hidden" id="pTitle" name="pTitle" value="<%=pTitle%>">
			<table width="100%" border="0" cellspacing="1" cellpadding="3" ID="Table2">
				<tr>
					<td class="title_form" style="font-size:14pt;line-height:24px" width="100%" align="center">טופס 
						מתעניין בטיול
					</td>
				</tr>
				<tr>
					<td align="center">
						<table border="0" align="center" width="100%">
							<tr>
								<td colspan="2" height="2" width="100%"></td>
							</tr>
							<%if contactID="" or request.querystring("isNewContact")="1" then%>
							<tr>
								<td align="right" class="subtitle_form" width="100%" colspan="2" class="title_form">פרטי 
									איש קשר</td>
							</tr>
							<tr>
								<td align="right" colspan="2">
								<div style="position:relative">
									<div style="position:absolute; top:0px;left:0px">
									<iframe style="display:none" name="frameContactsPr" id="frameContactsPr" src="contacts_privateQ16504.asp"
												ALLOWTRANSPARENCY="true" height="140" width="550" marginwidth="0" marginheight="0" hspace="0"
												vspace="0" scrolling="yes" frameborder="1"></iframe><input type=hidden name="companyId" id="companyID" value="<%=companyID%>">
											<input type=hidden name="contactID" id="contactID" value="<%=contactID%>">
										</div>
								</div>
											</td>
							</tr>
							<tr>
								<td align="right" width="100%"><input type=text dir="rtl" name="contact_name" value="<%=contactName%>" maxlength=50 style="width:400px" class=input_form onKeyUp="refreshContactsPr(this,'16504')" onKeyDown="refreshContactsPr(this,'16504')"  required ></td>
								<td align="right" nowrap class="form_title" valign="middle">
									&nbsp;<b>שם מלא</b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
							</tr>
							<tr>
								<td align="right"><input type=text dir="ltr" name="phone" value="" maxlength="20" value="<%=contactPhone%>" style="width:400px" class="input_form" ID="phone" required>
								</td>
								<td align="right" nowrap class="form_title" valign="middle">
									&nbsp;<b>טלפון</b>&nbsp;</td>
							</tr>
							<tr>
								<td align="right"><input type=text dir="ltr" name="cellular" value="<%=contactCellular%>" style="width:400px" maxlength=50 class="input_form" ID="cellular" required>
								</td>
								<td align="right" nowrap class="form_title" valign="middle">
									&nbsp;<b>טלפון נייד</b>&nbsp;</td>
							</tr>
							<tr>
								<td align="right"><input type=text name="email" dir=ltr value="<%=contactEmail%>" style="width:400px" maxlength=50 class="input_form" ID="email" dir="ltr" required>
								</td>
								<td align="right" nowrap class="form_title" valign="middle">
									&nbsp;<b>Email</b>&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
				<%end if%>
				<tr>
					<td width="100%" colspan="2" class="subtitle_form" align="right" valign="top" dir="ltr"
						style="padding-right:15px;"><span dir="rtl"><b>שאלות חובה במהלך השיחה המופנות ללקוח 
								מתעניין / מתעניין</b></span>
					</td>
				</tr>
				<tr>
					<td width="100%" class="formfield_title" align="right" valign="middle" dir="ltr"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>לאיזה 
								יעד נסיעה הינך מתעניין?</b></span>
					</td>
				</tr>
				<tr>
					<td align="right" width="100%" dir="rtl">
						<div class="ui-widget" >
							<select runat="server" id="CRMCountry" class="input_form" dir="rtl" name="CRMCountry"  style="width:480px" required>
							</select>
						</div>
					</td>
				</tr>
				<tr>
					<td width="100%" class="formfield_title" align="right" valign="middle" dir="ltr">
						<span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>סדרה </b></span>
					</td>
				</tr>
				<tr>
					<td align="right" width="100%" dir="rtl">
						<div class="ui-widget" >
							<select runat="server" id="selSerias" class="input_form" dir="rtl" name="selSerias" style="width:480px">
							</select>
						</div>
					</td>
				</tr>
				<tr>
					<td width="100%" class="formfield_title" align="right" valign="middle" dir="ltr"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>באיזה 
								חודש תרצה לנסוע? </b></span>
					</td>
				</tr>
				<tr>
					<td width="100%" align="right">
						<select runat="server" id="field40109" dir="rtl" name="field40109" class="input_form" style="width:220px"
							required>
							<option></option>
							<option value="ינואר">ינואר</option>
							<option value="פברואר">פברואר</option>
							<option value="מרץ">מרץ</option>
							<option value="אפריל">אפריל</option>
							<option value="מאי">מאי</option>
							<option value="יוני">יוני</option>
							<option value="יולי">יולי</option>
							<option value="אוגוסט">אוגוסט</option>
							<option value="ספטמבר">ספטמבר</option>
							<option value="אוקטובר">אוקטובר</option>
							<option value="נובמבר">נובמבר</option>
							<option value="דצמבר">דצמבר</option>
							<option value="לא שאלתי">לא שאלתי</option>
						</select>
					</td>
				</tr>
				<tr>
					<td width="100%" class="form_title" align="right" valign="top" dir="ltr"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>האם 
								נסעת עם פגסוס בעבר? </b></span>
					</td>
				</tr>
				<tr>
					<td  align="right" colspan="2">
						<fieldset style="border:1px solid #cccccc; padding:5px; width:210px;max-width:auto">
						<table border="0" cellspacing="1" cellpadding="1" dir="rtl">
								<tr>
									<td valign="middle"><input type="radio" id="field40167_0" name="field40167" value="כן" <%if v40167="כן" then%> checked="checked" <%end if%> required></td>
									<td class="form" dir="rtl" valign="middle">כן</td>
									<td width="20">&nbsp;</td>
									<td valign="middle"><input type="radio" id="field40167_1" name="field40167" value="לא"></td>
									<td class="form" dir="rtl" valign="middle">לא</td>
									<td width="20">&nbsp;</td>
									<td valign="middle"><input type="radio" id="field40167_2" name="field40167" value="לא שאלתי"></td>
									<td class="form" dir="rtl" valign="middle">לא שאלתי</td>
								</tr>
							</tbody>
						</table>
						</fieldset>
					</td>
				</tr>
				<tr>
					<td width="100%" class="form_title" align="right" valign="top" dir="ltr"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>כמה 
								מטיילים תהיו ? </b></span>
					</td>
				</tr>
				<tr>
					<td width="100%" align="right">
						<input dir="rtl" type="text" class="input_form" name="field40846" id="field40846" maxlength="50"
							style="width:220px" required onKeyPress="return getNumbers(this)" ">
					</td>
				</tr>
				<tr>
					<td width="100%" class="form_title" align="right" valign="top" dir="ltr"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font><%=TitleCompetitor%>
							</b></span>
					</td>
				</tr>
				<tr>
					<td width="100%" align="right">
							<select runat="server" id="selCompetitor" dir="rtl" class="input_form" name="selCompetitor" style="width: 460px" required>
										</select>
					</td>
				</tr>
				<tr>
					<td width="100%" class="form_title" align="right" valign="top" dir="ltr"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>תוכן 
								הפנייה </b></span>
					</td>
				</tr>
				<tr>
					<td width="100%" align="right">
						<textarea dir="rtl" class="textarea_form" rows="5" cols="30" style="width: 460px" id="field40103"
							name="field40103" required></textarea>
					</td>
				</tr>
				<tr>
					<td width="100%" class="form_title" align="right" valign="top" dir="ltr"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>מתי 
								כדי לחזור אליו במעקב - (נא להחמיר לטובת פגסוס) </b></span>
					</td>
				</tr>
				<tr>
					<td align="right" valign="middle" width="100%" class="form" bgcolor="#ffffff" dir="ltr">
						<img class="hand" src="../../images/calend.gif" id="imgcal" name="imgcal" border="0"
							onclick="cal.select(window.document.getElementById(&quot;field40170&quot;), &quot;field40170&quot;, &quot;dd/MM/yyyy&quot;); return false;">&nbsp;
						<input type="text" dir="ltr" class="input_form" maxlength="10" id="field40170" name="field40170" style="width:220px" value="<%=String.Format("{0:dd/MM/yyyy}", dateBack)%>" onclick="cal.select(this,this.id,&quot;dd/MM/yyyy&quot;); return false;" readonly="" >
					</td>
				</tr>
				<tr>
					<td width="100%" class="form_title" align="right" valign="top" dir="ltr">
						<font color="red"><span id="word19" name="word19"><!--שדות חובה--> * - שדות 
								חובה&nbsp;-&nbsp;*</span></font></td>
				</tr>
				<tr>
					<td width="100%" class="form" align="center" valign="top">
						<input type="submit" id="btnS" value="שלח" name="btnS" class="form_submit">
					</td>
				</tr>
			</table>
			</td> </tr> </table>
		</form>
	</body>
</html>

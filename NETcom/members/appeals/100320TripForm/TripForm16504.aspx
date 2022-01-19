<%@ Page Language="vb" AutoEventWireup="false" EnableViewState="true" Codebehind="TripForm16504.aspx.vb" Inherits="bizpower_pegasus2018.TripForm16504"  validateRequest="false"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>TripForm16504</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    	<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
    	  <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
      <link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css">
     <script src="//code.jquery.com/ui/1.11.4/jquery-ui.min.js"></script>
     <script language="javascript" type="text/javascript" src="../../../Scripts/CalendarPopupH.js"></script>
<script language="javascript" type="text/javascript">
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
    <script language="javascript" type="text/javascript" src="../../../Scripts/CalendarPopupH.js"></script>
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
                    .attr("required" ,"required")
                    .attr('id', 'input_' + $(this.element).get(0).id)
                    .attr('name', 'input_' + $(this.element).get(0).id)
                    .css('width', $(this.element).width())
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
                        //$(this).val('');
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
                                $(event.target).width(getwid);


                                if ($(event.target).attr("id") == 'input_CRMCountry') {
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
                    var matcher = new RegExp($.ui.autocomplete.escapeRegex(request.term), "i");
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

  
  </head>
  <body>
    <form id="Form1" method="post" runat="server">
    <input type="hidden" id="pTitle" name="pTitle" value="<%=pTitle%>">
    	<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#F0F0F0" ID="Table2">
    	<tr><td align =center><table border=0 align=center width =100%>
	<tr>
						<td colspan="2" height="2" width="100%" style="bgcolor:#808080"></td>
					</tr>
					<%if contactID="" then%>
					<tr><td align="right" class="title_form" width="100%" colspan="2" class="title_form">פרטי איש קשר</td></tr>
					
		<tr>
		<td align="right" colspan=2 bgcolor=#C9C9C9 style="border-bottom: 1px solid #808080;padding-right:15px;padding-left:15px">
		<table cellpadding=2 cellspacing=0 border=0 width=100% align="right">
		<tr><td colspan=3 height=10 nowrap></td></tr>
		<tr></tr><td width=100% align="right" rowspan=5 valign=top>
			<iframe style="display:none" name="frameContactsPr" id="frameContactsPr" src="contacts_privateQ16504.asp" ALLOWTRANSPARENCY=true height="140" width=550 marginwidth="0" marginheight="0" hspace="0" vspace="0" scrolling="yes" frameborder="1" ></iframe>  
			<input type=hidden name="companyId" id="companyID" value="<%=companyID%>">
			<input type=hidden name="contactID" id="contactID" value="<%=contactID%>">			
			</td>
			</tr>
		<tr><td  align="right"><input type=text dir="rtl" name="contact_name" value="<%'=vFix(contactName)%>" maxlength=50 style="width:200" class=Form onKeyUp="refreshContactsPr(this,'16504')" onKeyDown="refreshContactsPr(this,'16504')"  ></td>
		<td align="right" nowrap class="form_title" valign=top>					
			&nbsp;<b>שם מלא</b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
		</tr> 			    
	
		<tr><td  align="right"><input type=text dir="ltr" name="phone" value="" maxlength="20" value="<%'=vFix(contactPhone)%>" style="width:100" class="Form" ID="phone">  </td>
		<td align="right" nowrap class="form_title" valign=top>					
			&nbsp;<b>טלפון</b>&nbsp;</td>
		</tr>
		<tr><td  align="right"><input type=text dir="ltr" name="cellular" value="<%'=vFix(contactCellular)%>" style="width:100" maxlength=20 class="Form" ID="cellular">    </td>
		<td align="right" nowrap class="form_title" valign=top>					
			&nbsp;<b>טלפון נייד</b>&nbsp;</td>
		</tr>
		<tr><td  align="right"><input type=text name="email" dir=ltr value="<%'=vFix(contactEmail)%>" style="width:200" maxlength=50 class="Form" ID="email">     </td>
		<td align="right" nowrap class="form_title" valign=top>					
			&nbsp;<b>Email</b>&nbsp;</td>
		</tr> 		 			 			    
	
		</table>
		</td></tr>
		<%end if%>
					<tr>
						<td width="100%" colspan="2" class="title_form" bgcolor="#DADADA" align="right" valign="top"
							dir="ltr" style="padding-right:15px;"><span dir="rtl"><b>שאלות חובה במהלך השיחה המופנות 
									ללקוח מתעניין / מתעניין</b></span>
						</td>
					</tr>
					<tr>
						<td width="100%" class="form" bgcolor="#DADADA" align="right" valign="top" dir="ltr"
							style="padding-right:15px;"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>לאיזה יעד נסיעה הינך מתעניין?</b></span>
						</td>
					</tr>
					<tr>
						<td width="100%" align="right" colspan=2>
							<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#F0F0F0" ID="Table2">
								<tr>
									<td align="right" width="100%" dir="rtl" style="padding-right:15px;">
									   <div class="ui-widget" style="width: 100%">
                                <select runat="server" id="CRMCountry" dir="rtl" name="CRMCountry" required >
                                </select>
                            </div>
									
									</td>
								</tr>

	<tr>
						<td width="100%" class="form" bgcolor="#DADADA" align="right" valign="top" dir="ltr"
							style="padding-right:15px;"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>סדרה
									</b></span>
						</td>
					</tr>
					<tr>
						<td width="100%" align="right" colspan=2>
							<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#F0F0F0" ID="Table2">
								<tr>
									<td align="right" width="100%" dir="rtl" style="padding-right:15px;">
									   <div class="ui-widget" style="width: 100%">
                                <select runat="server" id="selSerias" dir="rtl" name="selSerias" style="width:180px">
                                </select>
                            </div>
							</td>
								</tr>
							</table>
						</td>
					</tr>
						<tr>
						<td width="100%" class="form" bgcolor="#DADADA" align="right" valign="top" dir="ltr"
							style="padding-right:15px;"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>באיזה חודש תרצה לנסוע?
									</b></span>
						</td>
					</tr>
						<tr>
						<td width="100%" align="right" colspan=2>
							<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#F0F0F0" ID="Table2">
								<tr>
									<td align="right" width="100%" dir="rtl" style="padding-right:15px;">
									   <div class="ui-widget" style="width: 100%">
                                <select runat="server" id="field40109" dir="rtl" name="field40109" style="width:180px" required >
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
                            </div>
							</td>
								</tr>
							</table>
						</td>
					</tr>
						<tr>
						<td width="100%" class="form" bgcolor="#DADADA" align="right" valign="top" dir="ltr"
							style="padding-right:15px;"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>האם נסעת עם פגסוס בעבר?
									</b></span>
						</td>
					</tr>
		<tr>
						<td width="100%" align="right" colspan=2>
					
	<table border="0" cellspacing="1" cellpadding="1" dir="rtl">
<tbody><tr>
<td width="5">&nbsp;</td><td valign="middle"><input type="radio" id="field40167_0" name="field40167" value="כן" <%if v40167="כן" then%> checked="checked" <%end if%> required></td><td class="form" dir="rtl" valign="middle">כן</td>
<td width="5">&nbsp;</td><td valign="middle"><input type="radio" id="field40167_1" name="field40167" value="לא"></td><td class="form" dir="rtl" valign="middle" >לא</td>
<td width="5">&nbsp;</td><td valign="middle"><input type="radio" id="field40167_2" name="field40167" value="לא שאלתי"></td><td class="form" dir="rtl" valign="middle">לא שאלתי</td>
</tr>
</tbody></table>
	</td></TR>
					<tr>
						<td width="100%" class="form" bgcolor="#DADADA" align="right" valign="top" dir="ltr"
							style="padding-right:15px;"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>כמה מטיילים תהיו ?
									</b></span>
						</td>
					</tr>
		
						<tr>
						<td width="100%" align="right" colspan=2>
<input dir="rtl" type="text" class="Form" name="field40846" id="field40846" maxlength="100" style="width:50" required onKeyPress="return getNumbers(this)" oninvalid="this.setCustomValidity('אנא מלא שדה חובה')">	</td></tr>

					<tr>
						<td width="100%" class="form" bgcolor="#DADADA" align="right" valign="top" dir="ltr"
							style="padding-right:15px;"  oninvalid="this.setCustomValidity('אנא מלא שדה חובה')"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>תוכן הפנייה
									</b></span>
						</td>
					</tr>
		
						<tr>
						<td width="100%" align="right" colspan=2>
				<textarea dir="rtl" class="txt" rows="5" cols="30" style="width: 220px" id="field40103" name="field40103" required oninvalid="this.setCustomValidity('אנא מלא שדה חובה')"></textarea>
	
</td>
					</tr>
			



			<tr>
						<td width="100%" class="form" bgcolor="#DADADA" align="right" valign="top" dir="ltr"
							style="padding-right:15px;"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font><%=TitleCompetitor%>
									</b></span>
						</td>
					</tr>
						<tr>
						<td width="100%" align="right" colspan=2>
							<table width="100%" border="0" cellspacing="1" cellpadding="3" bgcolor="#F0F0F0" ID="Table2">
								<tr>
									<td align="right" width="100%" dir="rtl" style="padding-right:15px;">
									   <div class="ui-widget" style="width: 100%">
                                <select runat="server" id="selCompetitor" dir="rtl" name="selCompetitor" required >
                                </select>
                            </div>
									
									</td>
								</tr></table></td></tr>

				
<tr>
						<td width="100%" class="form" bgcolor="#DADADA" align="right" valign="top" dir="ltr"
							style="padding-right:15px;"><span dir="rtl"><b><font color="red">&nbsp;*&nbsp;</font>מתי כדי לחזור אליו במעקב - (נא להחמיר לטובת פגסוס)
									</b></span>
						</td>
					</tr>
					<tr>		
				<td align="right" valign="middle" width="100%" class="form" bgcolor="#f0f0f0" style="padding-right:15px;" dir="ltr">					
					<img class="hand" src="../../images/calend.gif" id="imgcal" name="imgcal" border="0" onclick="cal.select(window.document.getElementById(&quot;field40170&quot;), &quot;field40170&quot;, &quot;dd/MM/yyyy&quot;); return false;">&nbsp;
					<input type="text" dir="ltr" class="passw" maxlength="10" id="field40170" name="field40170" size="10" value="<%=String.Format("{0:dd/MM/yyyy}", dateBack)%>" onclick="cal.select(this,this.id,&quot;dd/MM/yyyy&quot;); return false;" readonly="" >
				</td>
			</tr>
			<tr>
						<td width="100%" class="form" bgcolor="#DADADA" align="right" valign="top" dir="ltr"
							style="padding-right:15px;">		<font color="red"><span id="word19" name="word19"><!--שדות חובה-->* - שדות חובה&nbsp;-&nbsp;*</span></font></td></tr>
				<tr>
						<td width="100%" class="form" bgcolor="#DADADA" align="right" valign="top" dir="ltr"
							style="padding-right:15px;">								 
							<input type="submit" id="btnS" value ="שלח" name="btnS">
		</td></tr>		
					</table>
					</td></tr></table>
    </form>

  </body>
</html>

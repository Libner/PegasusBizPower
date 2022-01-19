<!DOCTYPE html>
<html>
	<head>
		<title>Bizpower - כלים חזקים אונליין</title>
		<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
		<link rel="stylesheet" type="text/css" href="../../includeCss/adminStyleNew.css">
		<link rel="stylesheet" type="text/css" href="../../includeCss/font-awesome.min.css">
		<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
		<script src="../../include/CalendarPopup.js" charset="windows-1255"></script>
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
		<link rel="stylesheet" href="/resources/demos/style.css">
		<script src="https://ajax.googleapis.com/ajax/libs/webfont/1.4.7/webfont.js"></script>
		<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.6.4/jquery.min.js"></script>
		<style>
  .custom-sSeries {
    position: relative;
    display: inline-block;
    padding: 0;
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
     /*width:58px;*/     
     width:100%;
    direction:rtl;
     height:26px;
    background: #f5f5f5;
    font-size: 16px;
    -moz-border-radius: 10px;
    -webkit-border-radius: 10px;
    border-radius: 10px;
    border: none;
    padding: 5px 5px;
    box-shadow: inset 0 2px 3px rgba( 0, 0, 0, 0.1 );
    clear: both;
}
.button_select_arrow
{
	font-family: 'Open Sans Hebrew';
    padding: 5px 5px;   
    text-decoration:none;
    -webkit-border-radius: 4px;
    -moz-border-radius: 4px;
    border-radius: 4px;
    background: #f5f5f5;
    font-size: 16px;
   
    box-shadow: inset 0 2px 3px rgba( 0, 0, 0, 0.1 );
    -webkit-transition-duration: 0.2s;
    -moz-transition-duration: 0.2s;
    transition-duration: 0.2s;
    -webkit-user-select:none;
    -moz-user-select:none;
    -ms-user-select:none;
    user-select:none;
    height: 36px;
    width: 36px !important;
}
.custom-sSeries-input:focus {
    background: #fff;
    box-shadow: 0 0 0 3px #E0E0E0, inset 0 2px 3px rgba( 0, 0, 0, 0.2 ), 0px 5px 5px rgba( 0, 0, 0, 0.15 );
    outline: none;
}
 
 .ui-button, .ui-button:link, .ui-button:visited, .ui-button:hover, .ui-button:active {
    text-decoration: none;
    width:16px;
}
  
  .ui-widget {
    	font-family: 'Open Sans Hebrew' !important;
		font-size:12pt;
		font-weight:600;
    color:navy;
    direction:rtl;
    overflow-x:hidden;
    -ms-overflow-x:hidden
}

.ui-menu .ui-menu-item-wrapper {
    position: relative;
    font-size:15px;
    padding: 2px 1em 2px .1em;
}
 
.ui-widget.ui-widget-content {
    border: 1px solid #c5c5c5;
    background:#ffffff;
     max-height: 250px;
     /*width:58px;*/     
     width:auto;
     overflow-y: scroll;
     overflow-x:hidden;
   
}

.plata_search
{
    background:#a5a5a5;
    -webkit-border-radius: 10px;
    -moz-border-radius: 10px;
    border-radius: 10px;
   padding:20px 20px 20px 60px;
    box-shadow: 0 0 0 3px #E0E0E0, inset 0 2px 3px rgba( 0, 0, 0, 0.2 ), 0px 5px 5px rgba( 0, 0, 0, 0.15 );
    outline: none;
    -webkit-transition-duration: 0.2s;
    -moz-transition-duration: 0.2s;
    transition-duration: 0.2s;
    -webkit-user-select:none;
    -moz-user-select:none;
    -ms-user-select:none;
    user-select:none;
}
.field_search{
		padding-left: 5px;
		font-family: 'Open Sans Hebrew' !important;
		font-size:14pt;
		font-weight:700;
		color:#ffffff;
}
.plata_results{
    background:#ffffff;
    -webkit-border-radius: 10px;
    -moz-border-radius: 10px;
    border-radius: 10px;
   padding:3px;
    box-shadow: 0 0 0 3px #E0E0E0, inset 0 2px 3px rgba( 0, 0, 0, 0.2 ), 0px 5px 5px rgba( 0, 0, 0, 0.15 );
    outline: none;
    -webkit-transition-duration: 0.2s;
    -moz-transition-duration: 0.2s;
    transition-duration: 0.2s;
    -webkit-user-select:none;
    -moz-user-select:none;
    -ms-user-select:none;
    user-select:none;
}
.plata_description{
    background:#f5f5f5;
    -webkit-border-radius: 10px;
    -moz-border-radius: 10px;
    border-radius: 10px;
   padding:5px;
    box-shadow: 0 0 0 3px #E0E0E0, inset 0 2px 3px rgba( 0, 0, 0, 0.2 ), 0px 5px 5px rgba( 0, 0, 0, 0.15 );
    outline: none;
    -webkit-transition-duration: 0.2s;
    -moz-transition-duration: 0.2s;
    transition-duration: 0.2s;
    -webkit-user-select:none;
    -moz-user-select:none;
    -ms-user-select:none;
    user-select:none;
}
.thParam
	{
		padding: 5px;
		font-family: 'Open Sans Hebrew' !important;
		font-size:12pt;
		font-weight:700;
	}

.paramName
	{
		padding: 5px;
		font-family: 'Open Sans Hebrew' !important;
		font-size:12pt;
		font-weight:700;
	}
  
.paramVal
	{
		padding: 5px ;
		font-family: 'Open Sans Hebrew' !important;
		font-size:12pt;
		font-weight:400;
	}
.paramDesc
	{
		padding: 5px ;
		font-family: 'Open Sans Hebrew' !important;
		font-size:12pt;
		font-weight:400;
	}
	
.tourTitle
	{
		font-family: 'Open Sans Hebrew' !important;
		font-size:14pt;
		font-weight:700;
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
          .attr( 'id', 'input_'+$(this.element).get(0).id)
          .css( 'width', $(this.element).width())
          .addClass( "custom-sSeries-input ui-widget ui-widget-content ui-state-default ui-corner-right" )
          .autocomplete({
            delay: 0,
            minLength: 0,
            source: $.proxy( this, "_source" ),
			position: { my : "right top", at: "right bottom" }
          })
          .tooltip({
            classes: {
              "ui-tooltip": "ui-state-highlight"
            }
          })
          
           .on( "click", function() {
            $(this).val('');
            });
 
        this._on( this.input, {
          autocompleteselect: function( event, ui ) {
            ui.item.option.selected = true;
            this._trigger( "select", event, {
              item: ui.item.option
            });
            var getwid = this.element.width(); // <<<<<<<<<<<< HERE
            $(event.target).width(getwid); 
            /*added by Mila to change competitor title*/            
           // alert($("#selTour").val())
            //alert($("#selCompetitor").val())
            
            /*added by Mila to change competitors select for pegasus tour*/
            if ($(event.target).attr("id")=='input_selTour')
			{	
				//$( "#tourTitle" ).text(ui.item.label)
				$( "#selCompetitor" ).html(fillCompetitorsSelByTour($("#selTour").val(),$("#selCompetitor").val()))
			}
			results=getScreen($("#selTour").val(),$("#selCompetitor").val())
			if (results != '')
			{
				$("#blockSearchResults").html(getScreen($("#selTour").val(),$("#selCompetitor").val()))
				$("#blockSearchResults").show()
			}
			else
			{
				$("#blockSearchResults").html('')
				$("#blockSearchResults").hide()
			}
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
          .addClass( "custom-sSeries-toggle ui-corner-right button_select_arrow" )
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
          var textForSearch
          /*added by Mila to change competitor title*/
          itemId=$(this).parent().prop('id')
           if (itemId=='selTour')
				textForSearch=text
		  if (itemId=='selCompetitor')
		  {		
				var TermsForSearch=getTermsForSearch($( this ).val())				
				textForSearch=text + ' ' +TermsForSearch
          }
          if ( this.value && ( !request.term || matcher.test(textForSearch) ) )
          {
            return {
              label: text,
              value: text,
              option: this
            };
            }
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
 
    $( "#selCompetitor" ).sSeries();
    $( "#toggle" ).on( "click", function() {
      $( "#selCompetitor" ).toggle();
    });
     
    $( "#selTour" ).sSeries();
    $( "#toggle" ).on( "click", function() {
      $( "#selTour" ).toggle();
    });
    
  } );
		</script>
		<script>
 $(document).ready(function () {
    
	$("[id*='paramDesc_']").on("mouseover",function() {
	//alert($(this).height())
	//alert($(this).parent().find("[id*='paramDescAll_']").height())
	if($(this).height()<$(this).parent().find("[id*='paramDescAll_']").height())
	{
		$(this).parent().find("[id*='paramDescAll_']").toggle();
	}
	});
	$("[id*='paramDescAll_']").on("mouseout",function() {
	//alert($(this).height())
	
		$(this).toggle()
	});	
})
function openDesc(el)
{
	$(el).css('height','auto');
	$(el).parent().find("div[id*='iconDescOpen_']").hide()
	$(el).parent().find("a[id*='iconDescClose_']").show()
}
function hideDesc(el)
{
	$(el).css('height','28px');
	$(el).parent().find("div[id*='iconDescOpen_']").show()
	$(el).parent().find("a[id*='iconDescClose_']").hide()
}			
function getTermsForSearch(competitorId)
{
	TermsForSearch=""
	$.ajax({
		type: "POST",
		url:"getCompetitorSearchTerms.asp",
		data: {competitorId:competitorId},
		success: function(msg) {
			if(msg != "")
			{
				TermsForSearch=msg
			}
		},
		async:   false
	});
	return TermsForSearch
}

function  getScreen(tourId,competitorId)
 {
	htmlScreensList=""
	$.ajax({
		type: "POST",
		url:"getScreenToViewBySearchTerms.asp",
		data: {tourId: tourId,competitorId:competitorId},
		success: function(msg) {
			if(msg != "")
			{
			//alert(msg)
				htmlScreensList=msg
			}
		},
		async:   false
	});
	return htmlScreensList
}

function fillCompetitorsSelByTour(tourId,competitorId)
{
	htmlOptionsList=""
	$.ajax({
		type: "POST",
		url:"getCompetitorsToViewByTour.asp",
		data: {tourId: tourId,competitorId:competitorId},
		success: function(msg) {
			if(msg != "")
			{
			//alert(msg)
				htmlOptionsList=msg
			}
		},
		async:   false
	});
	return htmlOptionsList
}
		</script>
	</head>
	<body class="body_admin">
		<table width="780" cellspacing="1" cellpadding="2" align="center" border="0" ID="Table3"
			dir="ltr">
			<!--site Competitor-->
			<tr>
				<TD align="right">
					<div class="ui-widget" style="width:100%">
						<Select id="selTour" name="selTour" dir="rtl">
							<option value="0"></option>
							<option value="240">טיולים מאורגנים לאירופה - ולנסיה, ברצלונה והפירנאים
							</option>
							<option value="120" selected>טיולים מאורגנים לאירופה - טיול מאורגן לספרד - דרום 
								המדינה, אנדלוסיה וגיברלטר עם מדריד
							</option>
							<option value="36">טיולים מאורגנים לאירופה - טיול מאורגן לספרד - צפון, חבל הבאסקים 
								והפירנאים
							</option>
							<option value="102">טיולים מאורגנים לאירופה - טיול מאורגן לספרד- צפון מערב, חבל 
								הבאסקים והפירנאים
							</option>
							<option value="31">טיולים מאורגנים לאירופה - טיול מאורגן לפורטוגל בטיול מקיף לאורך 
								ולרוחב
							</option>
						</Select>
					</div>
				</TD>
				<td align="right" nowrap dir="rtl" class="td_admin_4_right_align"><b>&nbsp;שם 
						טיול&nbsp;</b></td>
			</tr>
			<tr>
				<TD align="right" dir="rtl" width="100%" class="title_form">
					<div class="ui-widget" style="width:auto">
						<Select id="selCompetitor" name="selCompetitor" dir="rtl">
							<option value="0"></option>
							<option value="4">רימון
							</option>
							<option value="2" selected>שטיח מעופף
							</option>
						</Select>
					</div>
				</TD>
				<td align="right" nowrap dir="rtl" class="title_form"><b>&nbsp;שם מתחרה&nbsp;</b></td>
			</tr>
			<tr>
				<td height="15" nowrap></td>
			</tr>
		</table>
	</body>
</html>

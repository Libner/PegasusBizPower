<!DOCTYPE html>
<html>
	<head>
		<title>Bizpower - כלים חזקים אונליין</title>
		<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
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
    font-family: Arial;
    font-size: 15px;
    font-weight: bold;
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
				$( "#selCompetitor" ).html(fillCompetitorsSelByTour($("#selTour").val(),$("#selCompetitor").val()))
			}
			$("#blockSearchResults").html(getScreen($("#selTour").val(),$("#selCompetitor").val()))
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
		<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table4">
			<tr>
				<td>
					<table border="0" align="center" width="100%" cellspacing="0" cellpadding="2" class="table_admin_2"
						ID="Table10">
						<tr>
							<td align="left" nowrap="nowrap" class="td_title_2" nowrap><a href="screensView.asp" class="button_admin_1"><i class="fa fa-arrow-left"></i></a></td>
							<td width="100%" align="center" class="td_admin_3"><b>מסך השוואה</b></td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td align="center" valign="top">
					<table border="0" class="edit_page_table_style" cellpadding="0" cellspacing="0" width="100%"
						ID="Table8">
						<tr>
							<td height="15" nowrap></td>
						</tr>
						<tr>
							<td width="100%">
								<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" ID="Table1">
									<tr>
										<td width="100%">
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
																<option value="157">טיולים מאורגנים לאירופה - פורטוגל בטיול קלאסי
																</option>
																<option value="156">טיולים מאורגנים לאירופה - פורטוגל עם מדיירה
																</option>
																<option value="122">טיולים מאורגנים לאירופה - פורטוגל עם מדריד
																</option>
																<option value="98">טיולים מאורגנים לאירופה - איטליה – הקרנבלים אזור האגמים ( גרדה ) 
																	וטוסקנה
																</option>
																<option value="47">טיולים מאורגנים לאירופה - איטליה – רומא ודרומה החוף האמלפיטיני
																</option>
																<option value="237">טיולים מאורגנים לאירופה - איטליה, פוליה, לה מרקה ואומבריה
																</option>
																<option value="35">טיולים מאורגנים לאירופה - דרום איטליה וסיציליה
																</option>
																<option value="30">טיולים מאורגנים לאירופה - צפון איטליה טוסקנה וליגוריה
																</option>
																<option value="29">טיולים מאורגנים לאירופה - טיול מאורגן לרומניה
																</option>
																<option value="28">טיולים מאורגנים לאירופה - טיול מאורגן לרומניה בטיול קולע ומלא
																</option>
																<option value="133">טיולים מאורגנים לאירופה - יוון - סלוניקי וצפון יוון
																</option>
																<option value="158">טיולים מאורגנים לאירופה - יוון בטיול מורחב עם האי קורפו יואנינה 
																	ומטאורה
																</option>
																<option value="100">טיולים מאורגנים לאירופה - יוון בטיול מיוחד עם הצפון כולל 
																	סלוניקי ומטאורה
																</option>
																<option value="32">טיולים מאורגנים לאירופה - יוון בטיול קלאסי עם יואנינה ומטאורה
																</option>
																<option value="145">טיולים מאורגנים לאירופה - נורבגיה - טיול מיוחד
																</option>
																<option value="59">טיולים מאורגנים לאירופה - סקנדינביה - הטיול הקלאסי
																</option>
																<option value="110">טיולים מאורגנים לאירופה - דרום צרפת - פרובאנס והריביירה הצרפתית
																</option>
																<option value="128">טיולים מאורגנים לאירופה - פרובאנס והריביירה הצרפתית
																</option>
																<option value="74">טיולים מאורגנים לאירופה - פריז- עיר האורות והזוהר
																</option>
																<option value="208">טיולים מאורגנים לאירופה - צרפת - דורדון, טולוז, והפירנאים 
																	הצרפתים
																</option>
																<option value="136">טיולים מאורגנים לאירופה - צרפת כולל בריטני ונורמנדי והלואר
																</option>
																<option value="155">טיולים מאורגנים לאירופה - הקרנבל בטנריף
																</option>
																<option value="200">טיולים מאורגנים לאירופה - טנריף - האיים הקנריים
																</option>
																<option value="123">טיולים מאורגנים לאירופה - אלבניה מקדוניה ומונטנגרו
																</option>
																<option value="57">טיולים מאורגנים לאירופה - בולגריה בטיול יחודי
																</option>
																<option value="143">טיולים מאורגנים לאירופה - סלובניה קרואטיה ומונטנגרו
																</option>
																<option value="60">טיולים מאורגנים לאירופה - קרואטיה סלובניה - טיול כוכב במלון 4 
																	כוכבים
																</option>
																<option value="61">טיולים מאורגנים לאירופה - קרואטיה סלובניה בוסניה מונטנגרו וסרביה
																</option>
																<option value="56">טיולים מאורגנים לאירופה - אירלנד וצפון אירלנד - האי הירוק
																</option>
																<option value="149">טיולים מאורגנים לאירופה - אירלנד קלאסי
																</option>
																<option value="239">טיולים מאורגנים לאירופה - אנגליה במסלול מיוחד
																</option>
																<option value="112">טיולים מאורגנים לאירופה - סקוטלנד
																</option>
																<option value="65">טיולים מאורגנים לאירופה - סקוטלנד אירלנד - כולל צפון אירלנד
																</option>
																<option value="37">טיולים מאורגנים לאירופה - איסלנד המסלול המקיף - במסגרת בראש אחר
																</option>
																<option value="126">טיולים מאורגנים לאירופה - איסלנד- טיול עומק לאי הקרח והאש - 
																	במסגרת בראש אחר
																</option>
																<option value="89">טיולים מאורגנים לאירופה - טיול מאורגן לאיסלנד מסלול חורף - 
																	במסגרת בראש אחר
																</option>
																<option value="121">טיולים מאורגנים לאירופה - פניני איסלנד - במסגרת בראש אחר
																</option>
																<option value="64">טיולים מאורגנים לאירופה - אוסטריה ושוויץ כוכבי הרי האלפים - טיול 
																	כוכב כפול **
																</option>
																<option value="2">טיולים מאורגנים לאירופה - אתונה בחורף עם טברנות וטיולים
																</option>
																<option value="33">טיולים מאורגנים לאירופה - ברלין - דרזדן פוטסדאם ולייפציג
																</option>
																<option value="50">טיולים מאורגנים לאירופה - טיול לוינה (צפון אוסטריה) בודפשט 
																	וברטיסלבה
																</option>
																<option value="92">טיולים מאורגנים לאירופה - לפלנד - טיול סאפארי חורף
																</option>
																<option value="127">טיולים מאורגנים לאירופה - עמק הריין , לוקסמבורג , שטרסבורג , 
																	היער השחור
																</option>
																<option value="52">טיולים מאורגנים לאירופה - צ'כיה וצפון אוסטריה (עם וינה)
																</option>
																<option value="86">טיולים מאורגנים לאירופה - קסמיי היער השחור
																</option>
																<option value="53">טיולים מאורגנים לאירופה - אוקראינה
																</option>
																<option value="146">טיולים מאורגנים לאירופה - טיול שייט נהרות ברוסיה בין מוסקבה 
																	וסנט פטרבורג
																</option>
																<option value="34">טיולים מאורגנים לאירופה - פולין המתחדשת
																</option>
																<option value="119">טיולים מאורגנים לאירופה - צפון פולין והארצות הבלטיות - 
																	ליטא,לטביה,אסטוניה
																</option>
																<option value="40">טיולים מאורגנים לאירופה - רוסיה - סנט פטרבורג ומוסקבה
																</option>
																<option value="101">טיולים מאורגנים לאירופה - רוסיה - סנט פטרבורג ומוסקבה - קצר
																</option>
																<option value="68">טיולים מאורגנים לאירופה - תורכיה - הטיול הקלאסי המקיף
																</option>
																<option value="81">טיולים מאורגנים לאסיה - טיול לתאילנד בטיול מקיף - רגולר
																</option>
																<option value="26">טיולים מאורגנים לאסיה - טיול לתאילנד מלא וגדוש - פרימיום
																</option>
																<option value="45">טיולים מאורגנים לאסיה - בייג'ינג, בטיול מלא וגדוש
																</option>
																<option value="140">טיולים מאורגנים לאסיה - דרך המשי- קירגיזסטן ומערב סין - במסגרת 
																	בראש אחר
																</option>
																<option value="211">טיולים מאורגנים לאסיה - סין והונג קונג כולל בייג'ין,טונגלי 
																	סוג'ו,שנגחאי,שיאן גווילין ומקאו
																</option>
																<option value="46">טיולים מאורגנים לאסיה - סין כולל 3 החבלים סצ'ואן יונאן עם ארץ 
																	השנגרילה וגואייז'ו
																</option>
																<option value="95">טיולים מאורגנים לאסיה - סין כולל חבל גווילין, הונג קונג ומקאו
																</option>
																<option value="21">טיולים מאורגנים לאסיה - סין כולל חבל סצ'ואן הונג קונג ומקאו
																</option>
																<option value="139">טיולים מאורגנים לאסיה - סין ממלכת המזרח- הונג קונג ומקאו
																</option>
																<option value="20">טיולים מאורגנים לאסיה - סין מקיף כולל החבלים סצ'ואן ויונאן 
																	המורחב עם ארץ השנגרילה
																</option>
																<option value="80">טיולים מאורגנים לאסיה - סין שנגחאי והאקספו
																</option>
																<option value="23">טיולים מאורגנים לאסיה - סין, בייג'ין, גווילין, טונגלי, סוג'ו, 
																	שנגחאי
																</option>
																<option value="22">טיולים מאורגנים לאסיה - פניני סין, מקאו והונג קונג
																</option>
																<option value="63">טיולים מאורגנים לאסיה - פניני סין, מקאו והונג קונג
																</option>
																<option value="114">טיולים מאורגנים לאסיה - פסטיבל הקרח בחרבין- צפון סין - במסגרת 
																	בראש אחר
																</option>
																<option value="150">טיולים מאורגנים לאסיה - ארמניה בטיול מקיף ומיוחד
																</option>
																<option value="24">טיולים מאורגנים לאסיה - גאורגיה בטיול מקיף מלא גדוש ומיוחד
																</option>
																<option value="58">טיולים מאורגנים לאסיה - גאורגיה וארמניה - בטיול מקיף ומיוחד
																</option>
																<option value="131">טיולים מאורגנים לאסיה - גאורגיה כולל חבל סוונאטי המסתורי
																</option>
																<option value="111">טיולים מאורגנים לאסיה - דרום ויאטנם עם דלתת המקונג
																</option>
																<option value="27">טיולים מאורגנים לאסיה - וייאטנם קמבודיה
																</option>
																<option value="78">טיולים מאורגנים לאסיה - וייטנאם וקמבודיה במסלול מקיף ויחודי
																</option>
																<option value="182">טיולים מאורגנים לאסיה - וייטנאם קמבודיה
																</option>
																<option value="186">טיולים מאורגנים לאסיה - דרום קוריאה ודרום יפן - במסגרת פגסוס 
																	"בראש אחר"
																</option>
																<option value="84">טיולים מאורגנים לאסיה - יפן חדש וגם ישן כולל סיאול
																</option>
																<option value="93">טיולים מאורגנים לאסיה - יפן טיול מקיף - עם הירושימה ומיאג'ימה
																</option>
																<option value="18">טיולים מאורגנים לאסיה - יפן טיול מקיף - עם הירושימה וקויאסאן 
																	בטיול מיוחד
																</option>
																<option value="17">טיולים מאורגנים לאסיה - יפן עם הירושימה ומיאג'ימה
																</option>
																<option value="19">טיולים מאורגנים לאסיה - יפן, חדש וגם ישן
																</option>
																<option value="97">טיולים מאורגנים לאסיה - יפן, שונה ומיוחד
																</option>
																<option value="135">טיולים מאורגנים לאסיה - הודו - קומבה מהלה אל עומק תת היבשת - 
																	במסגרת בראש אחר
																</option>
																<option value="87">טיולים מאורגנים לאסיה - הודו ונפאל - בטיול מלא וגדוש
																</option>
																<option value="222">טיולים מאורגנים לאסיה - הודו -כולל את רישיקש
																</option>
																<option value="79">טיולים מאורגנים לאסיה - טיול מאורגן לדרום הודו כולל גואה - 
																	במסגרת בראש אחר
																</option>
																<option value="4">טיולים מאורגנים לאסיה - טיול מאורגן להודו
																</option>
																<option value="94">טיולים מאורגנים לאסיה - נפאל במסלול ייחודי - במסגרת בראש אחר
																</option>
																<option value="109">טיולים מאורגנים לאסיה - נפאל ובהוטן במסלול ייחודי - במסגרת בראש 
																	אחר
																</option>
																<option value="152">טיולים מאורגנים לאסיה - עם פגסוס להודו - בטיול מלא וגדוש
																</option>
																<option value="75">טיולים מאורגנים לאסיה - צפון הודו כולל רכס ההימאליה ונפאל
																</option>
																<option value="116">טיולים מאורגנים לאסיה - צפון מזרח הודו - סיקים ונפאל במסלול 
																	ייחודי במסגרת בראש אחר
																</option>
																<option value="25">טיולים מאורגנים לאסיה - אוזבקיסטן במסלול גיאוגרפי מקיף
																</option>
																<option value="39">טיולים מאורגנים לאסיה - אינדונזיה - מסע לארכיפלג מופלא - במסגרת 
																	בראש אחר
																</option>
																<option value="77">טיולים מאורגנים לאסיה - בורמה - מיאנמר - במסגרת בראש אחר
																</option>
																<option value="5">טיולים מאורגנים לאסיה - טיול מאורגן לסרי לנקה
																</option>
																<option value="76">טיולים מאורגנים לאסיה - מונגוליה- ארץ הנוודים - במסגרת בראש אחר
																</option>
																<option value="69">טיולים מאורגנים לאסיה - מזרח תורכיה בטיול גיאוגרפי
																</option>
																<option value="44">טיולים מאורגנים לאסיה - סיביר ומונגוליה - במסגרת בראש אחר
																</option>
																<option value="108">טיולים מאורגנים לאסיה - סין ומיאנמאר בתקופת פסטיבל מונאו 
																	ג'ינגפו - במסגרת בראש אחר
																</option>
																<option value="218">טיולים מאורגנים לאסיה - פיליפינים - המיטב
																</option>
																<option value="160">טיולים מאורגנים לאסיה - פיליפינים פרימיום - במסגרת בראש אחר
																</option>
																<option value="241">טיולים מאורגנים לאסיה - קירגיזסטן וקזחסטן - במסגרת בראש אחר
																</option>
																<option value="124">טיולים מאורגנים לאמריקה - הרי הרוקי הקנדיים ושייט באלסקה 
																	באוניית פאר
																</option>
																<option value="38">טיולים מאורגנים לאמריקה - טיול מאורגן לאלסקה- טיול יבשתי והרי 
																	הרוקי הקנדים
																</option>
																<option value="99">טיולים מאורגנים לאמריקה - טיול מאורגן לארה"ב - חוף מערבי, ניו 
																	יורק ושייט לריביירה המקסיקנית
																</option>
																<option value="210">טיולים מאורגנים לאמריקה - טיול מאורגן לארה"ב החוף המזרחי עם 
																	מיאמי אורלנדו
																</option>
																<option value="103">טיולים מאורגנים לאמריקה - טיול מאורגן לארה"ב וקנדה החוף המזרחי
																</option>
																<option value="73">טיולים מאורגנים לאמריקה - טיול מאורגן לארה"ב לחוף המערבי + שייט 
																	תענוגות לכיוון מכסיקו
																</option>
																<option value="43">טיולים מאורגנים לאמריקה - טיול מאורגן לארצות הברית וקנדה מהחוף 
																	המערבי לחוף המזרחי
																</option>
																<option value="70">טיולים מאורגנים לאמריקה - טיול מאורגן לקנדה - מזרח ומערב, כולל 
																	הרי הרוקי
																</option>
																<option value="62">טיולים מאורגנים לאמריקה - מזרח קנדה וארצות הברית ניו אינגלנד- 
																	טיול סתיו ושלכת
																</option>
																<option value="141">טיולים מאורגנים לאמריקה - אקוודור ואיי הגלאפגוס - כולל הג'ונגל 
																	האמזוני - במסגרת בראש אחר
																</option>
																<option value="154">טיולים מאורגנים לאמריקה - ארגנטינה וברזיל
																</option>
																<option value="153">טיולים מאורגנים לאמריקה - ארגנטינה וברזיל - הצעה מיוחדת
																</option>
																<option value="220">טיולים מאורגנים לאמריקה - ארגנטינה צ'ילה וברזיל
																</option>
																<option value="161">טיולים מאורגנים לאמריקה - ארגנטינה צ'ילה ואיגואסו
																</option>
																<option value="49">טיולים מאורגנים לאמריקה - ארגנטינה צ'ילה וברזיל
																</option>
																<option value="132">טיולים מאורגנים לאמריקה - ארגנטינה צ'ילה וברזיל - הצעה מיוחדת
																</option>
																<option value="83">טיולים מאורגנים לאמריקה - ארגנטינה צ'ילה וברזיל כולל פראטי
																</option>
																<option value="11">טיולים מאורגנים לאמריקה - ברזיל -איגואסו ,סלבדור ריו עם בואנוס 
																	איירס
																</option>
																<option value="117">טיולים מאורגנים לאמריקה - ברזיל ארגנטינה צ'ילה ופרו
																</option>
																<option value="90">טיולים מאורגנים לאמריקה - ברזיל וארגנטינה
																</option>
																<option value="106">טיולים מאורגנים לאמריקה - ברזיל ובואנוס איירס כולל שייט באיים 
																	הטרופיים
																</option>
																<option value="12">טיולים מאורגנים לאמריקה - ברזיל מקיף כולל האמאזונאס עם בואנוס 
																	איירס
																</option>
																<option value="107">טיולים מאורגנים לאמריקה - ברזיל צ'ילה וארגנטינה טיול המשלב את 
																	הקרחונים ואת שמורת הפאיינה
																</option>
																<option value="55">טיולים מאורגנים לאמריקה - פרו בוליביה - מסע אל רכס האנדים - 
																	במסגרת בראש אחר
																</option>
																<option value="159">טיולים מאורגנים לאמריקה - פרו בוליביה ארגנטינה צי'לה וברזיל - 
																	הצעה מיוחדת
																</option>
																<option value="72">טיולים מאורגנים לאמריקה - פרו, בוליביה, צ'ילה (מורחב), ארגנטינה, 
																	וברזיל
																</option>
																<option value="115">טיולים מאורגנים לאמריקה - צ'ילה ארגנטינה וברזיל
																</option>
																<option value="189">טיולים מאורגנים לאמריקה - קולומביה - במסגרת בראש אחר
																</option>
																<option value="232">טיולים מאורגנים לאמריקה - קולומביה ופנמה - במסגרת בראש אחר
																</option>
																<option value="13">טיולים מאורגנים לאמריקה - ריו דה ז'נרו מפליי האיגואסו ובואנוס 
																	איירס
																</option>
																<option value="148">טיולים מאורגנים לאמריקה - שייט אל חצי האי אנטרקטיקה - במסגרת 
																	בראש אחר
																</option>
																<option value="51">טיולים מאורגנים לאמריקה - גואטמלה,אל סלבדור,קופאן שבהונדורס וחבל 
																	הצ'ייאפס במכסיקו
																</option>
																<option value="134">טיולים מאורגנים לאמריקה - מכסיקו גואטמלה וקוסטה ריקה - טיול 
																	איכות במסלול ייחודי
																</option>
																<option value="225">טיולים מאורגנים לאמריקה - מקסיקו וגואטמלה
																</option>
																<option value="48">טיולים מאורגנים לאמריקה - מקסיקו עם גואטמלה ובליז
																</option>
																<option value="130">טיולים מאורגנים לאמריקה - קובה - הטיול המקיף כולל הדרום
																</option>
																<option value="14">טיולים מאורגנים לאמריקה - קובה וקוסטה ריקה
																</option>
																<option value="138">טיולים מאורגנים לאמריקה - קובה קוסטה ריקה ופנמה
																</option>
																<option value="9">טיולים מאורגנים לאפריקה - טיול למרוקו - טיול מלא וממצה
																</option>
																<option value="10">טיולים מאורגנים לאפריקה - טיול למרוקו והסהרה בדגש גיאוגרפי
																</option>
																<option value="125">טיולים מאורגנים לאפריקה - נמיביה – סודות המדבר - במסגרת בראש 
																	אחר
																</option>
																<option value="96">טיולים מאורגנים לאפריקה - נמיביה בוטסואנה ומפלי ויקטוריה - 
																	במסגרת בראש אחר
																</option>
																<option value="144">טיולים מאורגנים לאפריקה - נמיביה, בוטסואנה ו-מפלי ויקטוריה 
																	REGULAR - במסגרת בראש אחר
																</option>
																<option value="7">טיולים מאורגנים לאפריקה - דרום אפריקה ומפלי ויקטוריה ושמורת CHOBE 
																	בבוצואנה
																</option>
																<option value="118">טיולים מאורגנים לאפריקה - דרום אפריקה כולל סאן סיטי
																</option>
																<option value="113">טיולים מאורגנים לאפריקה - דרום אפריקה עם מפלי ויקטוריה וסאן 
																	סיטי
																</option>
																<option value="54">טיולים מאורגנים לאפריקה - אתיופיה - ערי הצפון ושבטי הדרום - 
																	במסגרת בראש אחר
																</option>
																<option value="88">טיולים מאורגנים לאפריקה - טיול מאורגן לקניה - ספארי - במסגרת 
																	בראש אחר
																</option>
																<option value="85">טיולים מאורגנים לאפריקה - טנזניה - שמורת הטבע המופלאות - במסגרת 
																	בראש אחר
																</option>
																<option value="82">טיולים מאורגנים לפסיפיק - ניו זילנד וסידני
																</option>
																<option value="42">טיולים מאורגנים לפסיפיק - ניו-זילנד
																</option>
																<option value="41">טיולים מאורגנים לפסיפיק - ניו-זילנד אוסטרליה - טיול איכות במסלול 
																	מקיף וייחודי
																</option>
																<option value="217">טיולי משפחות - ארה"ב - החוף המערבי עם ניו יורק
																</option>
																<option value="91">טיולי משפחות - לפלנד - טיול ספארי חורף - במסגרת בראש אחר
																</option>
																<option value="104">טיולי משפחות - סלובניה וקרואטיה טיול טבע והרפתקה
																</option>
																<option value="66">טיולי משפחות - פריס בריסל ואמסטרדם
																</option>
																<option value="71">טיולי משפחות - קסמיי היער השחור בטיול משפחות
																</option>
																<option value="137">טיולי משפחות - רומניה - טיול ספארי חורף - במסגרת בראש אחר
																</option>
																<option value="67">טיולי משפחות - תאילנד בטיול משפחות אקזוטי
																</option>
																<option value="105">טיול מאורגן בראש אחר - המשלחת לממלכות נסתרות בהימלאיה
																</option>
																<option value="206">טיולים קבוצות סגורות - אוסטריה
																</option>
																<option value="176">טיולים קבוצות סגורות - איטליה
																</option>
																<option value="198">טיולים קבוצות סגורות - איסלנד
																</option>
																<option value="190">טיולים קבוצות סגורות - אירלנד האי הירוק
																</option>
																<option value="204">טיולים קבוצות סגורות - אירלנד וסקוטלנד
																</option>
																<option value="175">טיולים קבוצות סגורות - אלבניה מונטנגרו מקדוניה
																</option>
																<option value="236">טיולים קבוצות סגורות - אנגליה
																</option>
																<option value="212">טיולים קבוצות סגורות - בולגריה
																</option>
																<option value="196">טיולים קבוצות סגורות - ברלין
																</option>
																<option value="185">טיולים קבוצות סגורות - היער השחור - שטרסבורג
																</option>
																<option value="167">טיולים קבוצות סגורות - וינה ובודפשט
																</option>
																<option value="184">טיולים קבוצות סגורות - טנריף
																</option>
																<option value="178">טיולים קבוצות סגורות - יוון
																</option>
																<option value="195">טיולים קבוצות סגורות - לפלנד
																</option>
																<option value="209">טיולים קבוצות סגורות - סיציליה
																</option>
																<option value="173">טיולים קבוצות סגורות - ספרד
																</option>
																<option value="215">טיולים קבוצות סגורות - סקנדינביה
																</option>
																<option value="203">טיולים קבוצות סגורות - פולין
																</option>
																<option value="171">טיולים קבוצות סגורות - פורטוגל
																</option>
																<option value="172">טיולים קבוצות סגורות - פורטוגל ומדירה
																</option>
																<option value="231">טיולים קבוצות סגורות - פורטוגל וספרד
																</option>
																<option value="174">טיולים קבוצות סגורות - צרפת
																</option>
																<option value="207">טיולים קבוצות סגורות - קיסמי היער השחור
																</option>
																<option value="202">טיולים קבוצות סגורות - קרואטיה וסלובניה
																</option>
																<option value="242">טיולים קבוצות סגורות - קרואטיה סלובניה בוסניה מונטנגרו סרביה
																</option>
																<option value="229">טיולים קבוצות סגורות - רומניה
																</option>
																<option value="214">טיולים קבוצות סגורות - רוסיה
																</option>
																<option value="233">טיולים קבוצות סגורות - אתיופיה
																</option>
																<option value="238">טיולים קבוצות סגורות - דרום אפריקה
																</option>
																<option value="165">טיולים קבוצות סגורות - מרוקו
																</option>
																<option value="166">טיולים קבוצות סגורות - מרוקו והסהרה
																</option>
																<option value="179">טיולים קבוצות סגורות - אוזבקיסטן
																</option>
																<option value="205">טיולים קבוצות סגורות - גיאורגיה
																</option>
																<option value="216">טיולים קבוצות סגורות - גיאורגיה וארמניה
																</option>
																<option value="213">טיולים קבוצות סגורות - דרום קוריאה
																</option>
																<option value="221">טיולים קבוצות סגורות - הודו
																</option>
																<option value="170">טיולים קבוצות סגורות - וייטנאם
																</option>
																<option value="227">טיולים קבוצות סגורות - וייטנאם קמבודיה ותאילנד
																</option>
																<option value="181">טיולים קבוצות סגורות - יפן
																</option>
																<option value="187">טיולים קבוצות סגורות - יפן עם הונג קונג ומקאו
																</option>
																<option value="219">טיולים קבוצות סגורות - סין בטיול מקיף
																</option>
																<option value="192">טיולים קבוצות סגורות - סין והונג קונג
																</option>
																<option value="201">טיולים קבוצות סגורות - סין ויפן
																</option>
																<option value="183">טיולים קבוצות סגורות - סרילנקה
																</option>
																<option value="194">טיולים קבוצות סגורות - פיליפינים
																</option>
																<option value="199">טיולים קבוצות סגורות - פניני סין
																</option>
																<option value="164">טיולים קבוצות סגורות - תאילנד
																</option>
																<option value="193">טיולים קבוצות סגורות - ארגנטינה צילה וברזיל
																</option>
																<option value="224">טיולים קבוצות סגורות - פרו צילה ארגנטינה וברזיל
																</option>
																<option value="162">טיולים קבוצות סגורות - פרו, אקוודור
																</option>
																<option value="188">טיולים קבוצות סגורות - מקסיקו וגואטמלה
																</option>
																<option value="223">טיולים קבוצות סגורות - מקסיקו וגואטמלה וקוסטה ריקה
																</option>
																<option value="197">טיולים קבוצות סגורות - קובה
																</option>
																<option value="230">טיולים קבוצות סגורות - קובה וגוטאמלה
																</option>
																<option value="168">טיולים קבוצות סגורות - קובה וקוסטה ריקה
																</option>
																<option value="169">טיולים קבוצות סגורות - קובה קוסטה ריקה ופנמה
																</option>
																<option value="228">טיולים קבוצות סגורות - קוסטה ריקה
																</option>
																<option value="234">טיולים קבוצות סגורות - אלסקה והרי הרוקיס הקנדיים
																</option>
																<option value="191">טיולים קבוצות סגורות - ארה"ב וקנדה מזרח ומערב
																</option>
																<option value="177">טיולים קבוצות סגורות - ארה"ב ושייט בקריבים
																</option>
																<option value="180">טיולים קבוצות סגורות - ניו זילנד ואוסטרליה
																</option>
																<option value="226">טיולים קבוצות סגורות - ניו זילנד והונג קונג
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
										</td>
									</tr>
									<tr>
										<td width="100%">
											<div id="blockSearchResults">
											</div>
										</td>
									</tr>
									<!--site Competitor-->
								</table>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</body>
</html>

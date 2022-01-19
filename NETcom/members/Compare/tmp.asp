<!DOCTYPE html>
<html>
	<head>
		<title>Bizpower - ���� ����� �������</title>
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
          .attr( "title", "���" )
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
							<td width="100%" align="center" class="td_admin_3"><b>��� ������</b></td>
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
																<option value="240">������ �������� ������� - ������, ������� ���������
																</option>
																<option value="120" selected>������ �������� ������� - ���� ������ ����� - ���� 
																	������, �������� �������� �� �����
																</option>
																<option value="36">������ �������� ������� - ���� ������ ����� - ����, ��� ������� 
																	���������
																</option>
																<option value="102">������ �������� ������� - ���� ������ �����- ���� ����, ��� 
																	������� ���������
																</option>
																<option value="31">������ �������� ������� - ���� ������ �������� ����� ���� ����� 
																	������
																</option>
																<option value="157">������ �������� ������� - ������� ����� �����
																</option>
																<option value="156">������ �������� ������� - ������� �� ������
																</option>
																<option value="122">������ �������� ������� - ������� �� �����
																</option>
																<option value="98">������ �������� ������� - ������ � �������� ���� ������ ( ���� ) 
																	�������
																</option>
																<option value="47">������ �������� ������� - ������ � ���� ������ ���� ����������
																</option>
																<option value="237">������ �������� ������� - ������, �����, �� ���� ��������
																</option>
																<option value="35">������ �������� ������� - ���� ������ ��������
																</option>
																<option value="30">������ �������� ������� - ���� ������ ������ ��������
																</option>
																<option value="29">������ �������� ������� - ���� ������ �������
																</option>
																<option value="28">������ �������� ������� - ���� ������ ������� ����� ���� ����
																</option>
																<option value="133">������ �������� ������� - ���� - ������� ����� ����
																</option>
																<option value="158">������ �������� ������� - ���� ����� ����� �� ��� ����� ������� 
																	�������
																</option>
																<option value="100">������ �������� ������� - ���� ����� ����� �� ����� ���� 
																	������� �������
																</option>
																<option value="32">������ �������� ������� - ���� ����� ����� �� ������� �������
																</option>
																<option value="145">������ �������� ������� - ������� - ���� �����
																</option>
																<option value="59">������ �������� ������� - ��������� - ����� ������
																</option>
																<option value="110">������ �������� ������� - ���� ���� - ������� ��������� �������
																</option>
																<option value="128">������ �������� ������� - ������� ��������� �������
																</option>
																<option value="74">������ �������� ������� - ����- ��� ������ ������
																</option>
																<option value="208">������ �������� ������� - ���� - ������, �����, ��������� 
																	�������
																</option>
																<option value="136">������ �������� ������� - ���� ���� ������ �������� ������
																</option>
																<option value="155">������ �������� ������� - ������ ������
																</option>
																<option value="200">������ �������� ������� - ����� - ����� �������
																</option>
																<option value="123">������ �������� ������� - ������ ������� ���������
																</option>
																<option value="57">������ �������� ������� - ������� ����� �����
																</option>
																<option value="143">������ �������� ������� - ������� ������� ���������
																</option>
																<option value="60">������ �������� ������� - ������� ������� - ���� ���� ����� 4 
																	������
																</option>
																<option value="61">������ �������� ������� - ������� ������� ������ �������� ������
																</option>
																<option value="56">������ �������� ������� - ������ ����� ������ - ��� �����
																</option>
																<option value="149">������ �������� ������� - ������ �����
																</option>
																<option value="239">������ �������� ������� - ������ ������ �����
																</option>
																<option value="112">������ �������� ������� - �������
																</option>
																<option value="65">������ �������� ������� - ������� ������ - ���� ���� ������
																</option>
																<option value="37">������ �������� ������� - ������ ������ ����� - ������ ���� ���
																</option>
																<option value="126">������ �������� ������� - ������- ���� ���� ��� ���� ���� - 
																	������ ���� ���
																</option>
																<option value="89">������ �������� ������� - ���� ������ ������� ����� ���� - 
																	������ ���� ���
																</option>
																<option value="121">������ �������� ������� - ����� ������ - ������ ���� ���
																</option>
																<option value="64">������ �������� ������� - ������� ������ ����� ��� ������ - ���� 
																	���� ���� **
																</option>
																<option value="2">������ �������� ������� - ����� ����� �� ������ �������
																</option>
																<option value="33">������ �������� ������� - ����� - ����� ������� ��������
																</option>
																<option value="50">������ �������� ������� - ���� ����� (���� �������) ������ 
																	���������
																</option>
																<option value="92">������ �������� ������� - ����� - ���� ������ ����
																</option>
																<option value="127">������ �������� ������� - ��� ����� , ��������� , �������� , 
																	���� �����
																</option>
																<option value="52">������ �������� ������� - �'��� ����� ������� (�� ����)
																</option>
																<option value="86">������ �������� ������� - ����� ���� �����
																</option>
																<option value="53">������ �������� ������� - ��������
																</option>
																<option value="146">������ �������� ������� - ���� ���� ����� ������ ��� ������ 
																	���� �������
																</option>
																<option value="34">������ �������� ������� - ����� �������
																</option>
																<option value="119">������ �������� ������� - ���� ����� ������� ������� - 
																	����,�����,�������
																</option>
																<option value="40">������ �������� ������� - ����� - ��� ������� �������
																</option>
																<option value="101">������ �������� ������� - ����� - ��� ������� ������� - ���
																</option>
																<option value="68">������ �������� ������� - ������ - ����� ������ �����
																</option>
																<option value="81">������ �������� ����� - ���� ������� ����� ���� - �����
																</option>
																<option value="26">������ �������� ����� - ���� ������� ��� ����� - �������
																</option>
																<option value="45">������ �������� ����� - ����'���, ����� ��� �����
																</option>
																<option value="140">������ �������� ����� - ��� ����- ��������� ����� ��� - ������ 
																	���� ���
																</option>
																<option value="211">������ �������� ����� - ��� ����� ���� ���� ����'��,������ 
																	���'�,������,���� ������� �����
																</option>
																<option value="46">������ �������� ����� - ��� ���� 3 ������ ��'��� ����� �� ��� 
																	�������� �������'�
																</option>
																<option value="95">������ �������� ����� - ��� ���� ��� �������, ���� ���� �����
																</option>
																<option value="21">������ �������� ����� - ��� ���� ��� ��'��� ���� ���� �����
																</option>
																<option value="139">������ �������� ����� - ��� ����� �����- ���� ���� �����
																</option>
																<option value="20">������ �������� ����� - ��� ���� ���� ������ ��'��� ������ 
																	������ �� ��� ��������
																</option>
																<option value="80">������ �������� ����� - ��� ������ �������
																</option>
																<option value="23">������ �������� ����� - ���, ����'��, �������, ������, ���'�, 
																	������
																</option>
																<option value="22">������ �������� ����� - ����� ���, ���� ����� ����
																</option>
																<option value="63">������ �������� ����� - ����� ���, ���� ����� ����
																</option>
																<option value="114">������ �������� ����� - ������ ���� ������- ���� ��� - ������ 
																	���� ���
																</option>
																<option value="150">������ �������� ����� - ������ ����� ���� ������
																</option>
																<option value="24">������ �������� ����� - ������� ����� ���� ��� ���� ������
																</option>
																<option value="58">������ �������� ����� - ������� ������� - ����� ���� ������
																</option>
																<option value="131">������ �������� ����� - ������� ���� ��� ������� �������
																</option>
																<option value="111">������ �������� ����� - ���� ������ �� ���� ������
																</option>
																<option value="27">������ �������� ����� - ������� �������
																</option>
																<option value="78">������ �������� ����� - ������� �������� ������ ���� ������
																</option>
																<option value="182">������ �������� ����� - ������� �������
																</option>
																<option value="186">������ �������� ����� - ���� ������ ����� ��� - ������ ����� 
																	"���� ���"
																</option>
																<option value="84">������ �������� ����� - ��� ��� ��� ��� ���� �����
																</option>
																<option value="93">������ �������� ����� - ��� ���� ���� - �� �������� �����'���
																</option>
																<option value="18">������ �������� ����� - ��� ���� ���� - �� �������� �������� 
																	����� �����
																</option>
																<option value="17">������ �������� ����� - ��� �� �������� �����'���
																</option>
																<option value="19">������ �������� ����� - ���, ��� ��� ���
																</option>
																<option value="97">������ �������� ����� - ���, ���� ������
																</option>
																<option value="135">������ �������� ����� - ���� - ����� ���� �� ���� �� ����� - 
																	������ ���� ���
																</option>
																<option value="87">������ �������� ����� - ���� ����� - ����� ��� �����
																</option>
																<option value="222">������ �������� ����� - ���� -���� �� ������
																</option>
																<option value="79">������ �������� ����� - ���� ������ ����� ���� ���� ���� - 
																	������ ���� ���
																</option>
																<option value="4">������ �������� ����� - ���� ������ �����
																</option>
																<option value="94">������ �������� ����� - ���� ������ ������ - ������ ���� ���
																</option>
																<option value="109">������ �������� ����� - ���� ������ ������ ������ - ������ ���� 
																	���
																</option>
																<option value="152">������ �������� ����� - �� ����� ����� - ����� ��� �����
																</option>
																<option value="75">������ �������� ����� - ���� ���� ���� ��� �������� �����
																</option>
																<option value="116">������ �������� ����� - ���� ���� ���� - ����� ����� ������ 
																	������ ������ ���� ���
																</option>
																<option value="25">������ �������� ����� - ��������� ������ �������� ����
																</option>
																<option value="39">������ �������� ����� - ��������� - ��� �������� ����� - ������ 
																	���� ���
																</option>
																<option value="77">������ �������� ����� - ����� - ������ - ������ ���� ���
																</option>
																<option value="5">������ �������� ����� - ���� ������ ���� ����
																</option>
																<option value="76">������ �������� ����� - ��������- ��� ������� - ������ ���� ���
																</option>
																<option value="69">������ �������� ����� - ���� ������ ����� ��������
																</option>
																<option value="44">������ �������� ����� - ����� ��������� - ������ ���� ���
																</option>
																<option value="108">������ �������� ����� - ��� �������� ������ ������ ����� 
																	�'����� - ������ ���� ���
																</option>
																<option value="218">������ �������� ����� - ��������� - �����
																</option>
																<option value="160">������ �������� ����� - ��������� ������� - ������ ���� ���
																</option>
																<option value="241">������ �������� ����� - ��������� ������� - ������ ���� ���
																</option>
																<option value="124">������ �������� ������� - ��� ����� ������� ����� ������ 
																	������� ���
																</option>
																<option value="38">������ �������� ������� - ���� ������ ������- ���� ����� ���� 
																	����� ������
																</option>
																<option value="99">������ �������� ������� - ���� ������ ����"� - ��� �����, ��� 
																	���� ����� �������� ���������
																</option>
																<option value="210">������ �������� ������� - ���� ������ ����"� ���� ������ �� 
																	����� �������
																</option>
																<option value="103">������ �������� ������� - ���� ������ ����"� ����� ���� ������
																</option>
																<option value="73">������ �������� ������� - ���� ������ ����"� ���� ������ + ���� 
																	������� ������ ������
																</option>
																<option value="43">������ �������� ������� - ���� ������ ������ ����� ����� ����� 
																	������ ���� ������
																</option>
																<option value="70">������ �������� ������� - ���� ������ ����� - ���� �����, ���� 
																	��� �����
																</option>
																<option value="62">������ �������� ������� - ���� ���� ������ ����� ��� �������- 
																	���� ���� �����
																</option>
																<option value="141">������ �������� ������� - ������� ���� �������� - ���� ��'���� 
																	������� - ������ ���� ���
																</option>
																<option value="154">������ �������� ������� - �������� ������
																</option>
																<option value="153">������ �������� ������� - �������� ������ - ���� ������
																</option>
																<option value="220">������ �������� ������� - �������� �'��� ������
																</option>
																<option value="161">������ �������� ������� - �������� �'��� ��������
																</option>
																<option value="49">������ �������� ������� - �������� �'��� ������
																</option>
																<option value="132">������ �������� ������� - �������� �'��� ������ - ���� ������
																</option>
																<option value="83">������ �������� ������� - �������� �'��� ������ ���� �����
																</option>
																<option value="11">������ �������� ������� - ����� -������� ,������ ��� �� ������ 
																	�����
																</option>
																<option value="117">������ �������� ������� - ����� �������� �'��� ����
																</option>
																<option value="90">������ �������� ������� - ����� ���������
																</option>
																<option value="106">������ �������� ������� - ����� ������� ����� ���� ���� ����� 
																	��������
																</option>
																<option value="12">������ �������� ������� - ����� ���� ���� ��������� �� ������ 
																	�����
																</option>
																<option value="107">������ �������� ������� - ����� �'��� ��������� ���� ����� �� 
																	�������� ��� ����� �������
																</option>
																<option value="55">������ �������� ������� - ��� ������� - ��� �� ��� ������ - 
																	������ ���� ���
																</option>
																<option value="159">������ �������� ������� - ��� ������� �������� ��'�� ������ - 
																	���� ������
																</option>
																<option value="72">������ �������� ������� - ���, �������, �'��� (�����), ��������, 
																	������
																</option>
																<option value="115">������ �������� ������� - �'��� �������� ������
																</option>
																<option value="189">������ �������� ������� - �������� - ������ ���� ���
																</option>
																<option value="232">������ �������� ������� - �������� ����� - ������ ���� ���
																</option>
																<option value="13">������ �������� ������� - ��� �� �'��� ����� �������� ������� 
																	�����
																</option>
																<option value="148">������ �������� ������� - ���� �� ��� ��� ��������� - ������ 
																	���� ���
																</option>
																<option value="51">������ �������� ������� - �������,�� ������,����� ��������� ���� 
																	��'����� �������
																</option>
																<option value="134">������ �������� ������� - ������ ������� ������ ���� - ���� 
																	����� ������ ������
																</option>
																<option value="225">������ �������� ������� - ������ ��������
																</option>
																<option value="48">������ �������� ������� - ������ �� ������� �����
																</option>
																<option value="130">������ �������� ������� - ���� - ����� ����� ���� �����
																</option>
																<option value="14">������ �������� ������� - ���� ������ ����
																</option>
																<option value="138">������ �������� ������� - ���� ����� ���� �����
																</option>
																<option value="9">������ �������� ������� - ���� ������ - ���� ��� �����
																</option>
																<option value="10">������ �������� ������� - ���� ������ ������ ���� ��������
																</option>
																<option value="125">������ �������� ������� - ������ � ����� ����� - ������ ���� 
																	���
																</option>
																<option value="96">������ �������� ������� - ������ �������� ����� �������� - 
																	������ ���� ���
																</option>
																<option value="144">������ �������� ������� - ������, �������� �-���� �������� 
																	REGULAR - ������ ���� ���
																</option>
																<option value="7">������ �������� ������� - ���� ������ ����� �������� ������ CHOBE 
																	��������
																</option>
																<option value="118">������ �������� ������� - ���� ������ ���� ��� ����
																</option>
																<option value="113">������ �������� ������� - ���� ������ �� ���� �������� ���� 
																	����
																</option>
																<option value="54">������ �������� ������� - ������� - ��� ����� ����� ����� - 
																	������ ���� ���
																</option>
																<option value="88">������ �������� ������� - ���� ������ ����� - ����� - ������ 
																	���� ���
																</option>
																<option value="85">������ �������� ������� - ������ - ����� ���� �������� - ������ 
																	���� ���
																</option>
																<option value="82">������ �������� ������� - ��� ����� ������
																</option>
																<option value="42">������ �������� ������� - ���-�����
																</option>
																<option value="41">������ �������� ������� - ���-����� �������� - ���� ����� ������ 
																	���� �������
																</option>
																<option value="217">����� ������ - ���"� - ���� ������ �� ��� ����
																</option>
																<option value="91">����� ������ - ����� - ���� ����� ���� - ������ ���� ���
																</option>
																<option value="104">����� ������ - ������� �������� ���� ��� �������
																</option>
																<option value="66">����� ������ - ���� ����� ��������
																</option>
																<option value="71">����� ������ - ����� ���� ����� ����� ������
																</option>
																<option value="137">����� ������ - ������ - ���� ����� ���� - ������ ���� ���
																</option>
																<option value="67">����� ������ - ������ ����� ������ ������
																</option>
																<option value="105">���� ������ ���� ��� - ������ ������� ������ ��������
																</option>
																<option value="206">������ ������ ������ - �������
																</option>
																<option value="176">������ ������ ������ - ������
																</option>
																<option value="198">������ ������ ������ - ������
																</option>
																<option value="190">������ ������ ������ - ������ ��� �����
																</option>
																<option value="204">������ ������ ������ - ������ ��������
																</option>
																<option value="175">������ ������ ������ - ������ �������� �������
																</option>
																<option value="236">������ ������ ������ - ������
																</option>
																<option value="212">������ ������ ������ - �������
																</option>
																<option value="196">������ ������ ������ - �����
																</option>
																<option value="185">������ ������ ������ - ���� ����� - ��������
																</option>
																<option value="167">������ ������ ������ - ���� �������
																</option>
																<option value="184">������ ������ ������ - �����
																</option>
																<option value="178">������ ������ ������ - ����
																</option>
																<option value="195">������ ������ ������ - �����
																</option>
																<option value="209">������ ������ ������ - �������
																</option>
																<option value="173">������ ������ ������ - ����
																</option>
																<option value="215">������ ������ ������ - ���������
																</option>
																<option value="203">������ ������ ������ - �����
																</option>
																<option value="171">������ ������ ������ - �������
																</option>
																<option value="172">������ ������ ������ - ������� ������
																</option>
																<option value="231">������ ������ ������ - ������� �����
																</option>
																<option value="174">������ ������ ������ - ����
																</option>
																<option value="207">������ ������ ������ - ����� ���� �����
																</option>
																<option value="202">������ ������ ������ - ������� ��������
																</option>
																<option value="242">������ ������ ������ - ������� ������� ������ �������� �����
																</option>
																<option value="229">������ ������ ������ - ������
																</option>
																<option value="214">������ ������ ������ - �����
																</option>
																<option value="233">������ ������ ������ - �������
																</option>
																<option value="238">������ ������ ������ - ���� ������
																</option>
																<option value="165">������ ������ ������ - �����
																</option>
																<option value="166">������ ������ ������ - ����� ������
																</option>
																<option value="179">������ ������ ������ - ���������
																</option>
																<option value="205">������ ������ ������ - ��������
																</option>
																<option value="216">������ ������ ������ - �������� �������
																</option>
																<option value="213">������ ������ ������ - ���� ������
																</option>
																<option value="221">������ ������ ������ - ����
																</option>
																<option value="170">������ ������ ������ - �������
																</option>
																<option value="227">������ ������ ������ - ������� ������� �������
																</option>
																<option value="181">������ ������ ������ - ���
																</option>
																<option value="187">������ ������ ������ - ��� �� ���� ���� �����
																</option>
																<option value="219">������ ������ ������ - ��� ����� ����
																</option>
																<option value="192">������ ������ ������ - ��� ����� ����
																</option>
																<option value="201">������ ������ ������ - ��� ����
																</option>
																<option value="183">������ ������ ������ - �������
																</option>
																<option value="194">������ ������ ������ - ���������
																</option>
																<option value="199">������ ������ ������ - ����� ���
																</option>
																<option value="164">������ ������ ������ - ������
																</option>
																<option value="193">������ ������ ������ - �������� ���� ������
																</option>
																<option value="224">������ ������ ������ - ��� ���� �������� ������
																</option>
																<option value="162">������ ������ ������ - ���, �������
																</option>
																<option value="188">������ ������ ������ - ������ ��������
																</option>
																<option value="223">������ ������ ������ - ������ �������� ������ ����
																</option>
																<option value="197">������ ������ ������ - ����
																</option>
																<option value="230">������ ������ ������ - ���� ��������
																</option>
																<option value="168">������ ������ ������ - ���� ������ ����
																</option>
																<option value="169">������ ������ ������ - ���� ����� ���� �����
																</option>
																<option value="228">������ ������ ������ - ����� ����
																</option>
																<option value="234">������ ������ ������ - ����� ���� ������ �������
																</option>
																<option value="191">������ ������ ������ - ���"� ����� ���� �����
																</option>
																<option value="177">������ ������ ������ - ���"� ����� �������
																</option>
																<option value="180">������ ������ ������ - ��� ����� ���������
																</option>
																<option value="226">������ ������ ������ - ��� ����� ����� ����
																</option>
															</Select>
														</div>
													</TD>
													<td align="right" nowrap dir="rtl" class="td_admin_4_right_align"><b>&nbsp;�� 
															����&nbsp;</b></td>
												</tr>
												<tr>
													<TD align="right" dir="rtl" width="100%" class="title_form">
														<div class="ui-widget" style="width:auto">
															<Select id="selCompetitor" name="selCompetitor" dir="rtl">
																<option value="0"></option>
																<option value="4">�����
																</option>
																<option value="2" selected>���� �����
																</option>
															</Select>
														</div>
													</TD>
													<td align="right" nowrap dir="rtl" class="title_form"><b>&nbsp;�� �����&nbsp;</b></td>
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

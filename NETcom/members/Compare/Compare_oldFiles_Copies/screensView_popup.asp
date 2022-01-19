<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	If Request.QueryString("deleteId") <> nil Then
		delId = trim(Request.QueryString("deleteId"))
		If Len(delId) > 0 Then
			con.executeQuery("Delete From Compare_Screens Where Screen_Id = " & delId & ";Delete From Compare_Parameters_ToScreen Where Screen_Id = " & delId)
		End If
		'!!!!!!!!!delete from screens
		
		Response.Redirect "Screens.asp"
	End If
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 30 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(arr_title(0)) = arr_title(1)
			End If
		Next
	  End If
	  set rstitle = Nothing	  

%>
<!DOCTYPE html>
<html>
	<head>
		<!-- #include file="../../title_meta_inc.asp" -->
		<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
			<link rel="stylesheet" type="text/css" href="../../includeCss/adminStyleNew.css">
				<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
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
.field_search{
		padding: 5px 5px;
		font-family: 'Open Sans Hebrew' !important;
		font-size:14pt;
		font-weight:700;
}
.thParam
	{
		padding: 5px 5px;
		font-family: 'Open Sans Hebrew' !important;
		font-size:12pt;
		font-weight:700;
	}

.paramName
	{
		padding: 5px 0px;
		font-family: 'Open Sans Hebrew' !important;
		font-size:12pt;
		font-weight:700;
	}
  
.paramVal
	{
		padding: 5px 0px;
		font-family: 'Open Sans Hebrew' !important;
		font-size:12pt;
		font-weight:400;
	}
.paramDesc
	{
		padding: 5px 0px;
		font-family: 'Open Sans Hebrew' !important;
		font-size:12pt;
		font-weight:400;
	}
	
.titleList
	{	
		display:block;
		background-color:#e5e5e5;
		padding: 10px;
		font-family: 'Open Sans Hebrew' !important;
		font-size:14pt;
		font-weight:700;
	}
	.itemList
	{
		display:block;
		background-color:#f5f5f5;
		font-family: 'Open Sans Hebrew' !important;
		font-size:14pt;
		font-weight:400;
		text-decoration:none
	}
	a.itemList
	{
		padding: 10px;
		color: #000000;
		font-family: 'Open Sans Hebrew' !important;
		font-size:14pt;
		font-weight:400;
		text-decoration:none
	}
	a.itemList:hover
	{
		font-family: 'Open Sans Hebrew' !important;
		font-size:14pt;
		font-weight:400;
		text-decoration:none;
		background-color:#e5e5e5;
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
            //alert($("#selCompetitor").val())
			$("#blockSearchResults").html(getScreensList($("#selTour").val(),$("#selCompetitor").val()))
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
            
			if (input.attr("id")=='input_selCompetitor')
				$("#blockSearchResults").html(getScreensList($("#selTour").val(),''))
			if (input.attr("id")=='input_selTour')
				$("#blockSearchResults").html(getScreensList('',$("#selCompetitor").val()))
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
  

 function getScreensList(tourId,competitorId)
 {
	htmlScreensList=""
	$.ajax({
		type: "POST",
		url:"getScreensToViewBySearchTerms_popup.asp",
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
			//alert(msg)
				TermsForSearch=msg
			}
		},
		async:   false
	});
	return TermsForSearch
}

function openScreen(screenId)
{
	window.open("viewScreen_popup.asp?ScreenId=" + screenId,"openScreen","width="+$(window).width() +", height="+ $(window).height() +", top=0, left=0, scrollbars=no;toolbar=no;statusbar=no;resize=no");
			
}
						</script>
	</head>
	<body>
		<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>" ID="Table1">
			<tr>
				<td width=100% align="<%=align_var%>">
					<!--#include file="../../logo_top.asp"-->
				</td>
			</tr>
			<tr>
				<td width=100% align="<%=align_var%>">
					<%numOftab = 95%>
					<%numOfLink = 3%>
					<!--#include file="../../top_in.asp"-->
				</td>
			</tr>
			<tr>
				<td class="page_title" align="<%=align_var%>" valign="middle" width="100%">מסכי 
					השוואה</td>
			</tr>
			<tr>
				<td align="center" valign="top">
					<table border="0" cellpadding="0" cellspacing="0" width="100%" ID="Table8">
						<tr>
							<td height="25" nowrap></td>
						</tr>
						<tr>
							<td width="100%">
								<table border="0"  cellspacing="0" cellpadding="0" align="center" ID="Table2">
									<tr>
										<td width="100%">
											<table  cellspacing="0" cellpadding="5" align="center" border="0" ID="Table4"
												dir="ltr" class="title_form">
												<tr>
													<td height="5" colspan="2" nowrap class="title_form"></td>
												</tr>
												<!--site Competitor-->
												<%	
											If Len(Screen_Id) > 0 Then
												sqlstr="Select Tour_Id, Competitor_Id, Start_Date, End_Date From Compare_Screens Where Screen_Id = " & Screen_Id
												set rssub = con.getRecordSet(sqlstr)
												If not rssub.eof Then
													Tour_Id = cint(rssub("Tour_Id"))
													Competitor_Id = cint(rssub("Competitor_Id"))
													Start_Date = cdate(rssub("Start_Date"))
													End_Date = cdate(rssub("End_Date"))
												End If
												set rssub=Nothing
											End If
											
										%>
												<tr>
													<TD align="<%=align_var%>" dir="<%=dir_obj_var%>"  class="title_form">
														<div class="ui-widget" style="width:auto">
															<Select id="selTour" name="selTour" dir="<%=dir_obj_var%>">
																<option value="0"></option>
																<%
															
														'get tours' names from Pegasusisrael site
															sqlstr="SELECT Tours.Tour_Id, Tours.Category_Id, Tours.SubCategory_Id, Tours.Tour_Name, Tours.Tour_Title, "
															sqlstr=sqlstr & " Tours_Categories.Category_Name,  Tours_SubCategories.SubCategory_Name FROM Tours INNER JOIN Tours_Categories ON Tours.Category_Id = Tours_Categories.Category_Id INNER JOIN "
															sqlstr=sqlstr & "  Tours_SubCategories ON Tours.SubCategory_Id = Tours_SubCategories.SubCategory_Id"
															sqlstr=sqlstr & "  order by Tours_Categories.Category_Order,Tours_SubCategories.SubCategory_Order,Tours.Tour_Name"
															'sqlstr=sqlstr & "  Where Tour_Id = " & Tour_Id
															set rssub = conPegasus.Execute(sqlstr)
															do while not rssub.eof															
																selTour_Id= cint(rssub("Tour_Id"))
																selCategory_Name = trim(rssub("Category_Name"))
																selSubCategory_Name = trim(rssub("SubCategory_Name"))
																selTour_Name = trim(rssub("Tour_Name"))
																%>
																<option value="<%=selTour_Id%>" <%if selTour_Id =Tour_Id then%>selected<%end if%>><%=selCategory_Name%>
																	-
																	<%=selTour_Name%>
																</option>
																<%
																rssub.movenext
															loop
															set rssub=Nothing%>
															</Select>
														</div>
													</TD>
													<TD align="<%=align_var%>" dir="<%=dir_obj_var%>" width="100%" class="title_form field_search"><b>&nbsp;שם 
															טיול&nbsp;</b></TD>
												</tr>
												<tr>
													<TD align="<%=align_var%>" dir="<%=dir_obj_var%>" width="100%" class="title_form" style="padding-left:30px">
														<div class="ui-widget" style="width:auto">
															<Select id="selCompetitor" name="selCompetitor" dir="<%=dir_obj_var%>">
																<option value="0"></option>
																<%
															
														'get Competitor' names 
															sqlstr="SELECT Competitor_Id,Competitor_Name from Compare_Competitors  order by Competitor_Name"
															'sqlstr=sqlstr & "  Where Tour_Id = " & Tour_Id
															set rssub = con.getRecordSet(sqlstr)
															do while not rssub.eof															
																selCompetitor_Id= cint(rssub("Competitor_Id"))																
																selCompetitor_Name = trim(rssub("Competitor_Name"))
																%>
																<option value="<%=selCompetitor_Id%>"  <%if selCompetitor_Id =Competitor_Id then%>selected<%end if%>><%=selCompetitor_Name%>
																</option>
																<%
																rssub.movenext
															loop
															set rssub=Nothing%>
															</Select>
														</div>
													</TD>
													<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>" class="title_form field_search"><b>&nbsp;שם 
															מתחרה&nbsp;</b></td>
												</tr>
												<tr>
													<td height="5" colspan="2" nowrap class="title_form"></td>
												</tr>
											</table>
										</td>
									</tr>
									<tr>
										<td height="10"></td>
									</tr>
									<tr>
										<td height="20"></td>
									</tr>
									<tr>
										<td width="100%" valign="top" align="center">
											<div id="blockSearchResults">
												<table cellspacing="1" cellpadding="2" border="0" bgcolor="#ffffff" ID="Table3">
													<%		'get tours' names from Pegasusisrael site
									sqlstr="SELECT distinct Compare_Screens.Tour_Id, " & pegasusDBName & ".dbo.Tours.Tour_Name, "
									sqlstr=sqlstr &  pegasusDBName & ".dbo.Tours_Categories.Category_Name,  " & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Name, "
									sqlstr=sqlstr &  pegasusDBName & ".dbo.Tours_Categories.Category_Order," & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Order, " & pegasusDBName & ".dbo.Tours.Tour_Order"
									sqlstr=sqlstr & " FROM   Compare_Screens INNER JOIN "
									sqlstr=sqlstr &  pegasusDBName & ".dbo.Tours ON Compare_Screens.Tour_Id = " & pegasusDBName & ".dbo.Tours.Tour_Id "
									sqlstr=sqlstr & " INNER JOIN " & pegasusDBName & ".dbo.Tours_Categories ON " & pegasusDBName & ".dbo.Tours.Category_Id = " & pegasusDBName & ".dbo.Tours_Categories.Category_Id INNER JOIN "
										sqlstr=sqlstr & pegasusDBName & ".dbo.Tours_SubCategories ON " & pegasusDBName & ".dbo.Tours.SubCategory_Id = " & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Id "
									sqlstr=sqlstr & " where datediff(dd,Compare_Screens.[Start_Date],getdate())>=0 and datediff(dd,Compare_Screens.[End_Date],getdate())<=0 "
										sqlstr=sqlstr & "order by " & pegasusDBName & ".dbo.Tours_Categories.Category_Order," & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Order, " & pegasusDBName & ".dbo.Tours.Tour_Order"
									set rs_Tours = con.GetRecordSet(sqlstr)
									do while not rs_Tours.eof
    										Tour_Id = trim(rs_Tours("Tour_Id"))
												Category_Name = trim(rs_Tours("Category_Name"))
												SubCategory_Name = trim(rs_Tours("SubCategory_Name"))
												Tour_Name = trim(rs_Tours("Tour_Name"))
											
										%>
													<tr>
														<td  align="<%=align_var%>" class="titleList">
															<b>
																<%=Category_Name%>
																- <font color="navy">
																	<%=Tour_Name%>
																</font></b>&nbsp;</td>
													</tr>
													<%
											sqlstr= "Select Screen_Id,Competitor_Name,Start_Date,End_Date  from Compare_Screens "
											sqlstr=sqlstr & "  inner join Compare_Competitors on Compare_Competitors.Competitor_Id=Compare_Screens.Competitor_Id where Tour_Id=" & Tour_Id 
											sqlstr=sqlstr & " and datediff(dd,Compare_Screens.[Start_Date],getdate())>=0 and datediff(dd,Compare_Screens.[End_Date],getdate())<=0 "
											sqlstr=sqlstr & " order by Competitor_Name,End_Date desc, Start_Date desc"
											set rs_Comp = con.GetRecordSet(sqlstr)
											do while not rs_Comp.eof
    											Screen_Id = trim(rs_Comp("Screen_Id"))
												Competitor_Name = trim(rs_Comp("Competitor_Name"))
												Start_Date = trim(rs_Comp("Start_Date"))
												End_Date = trim(rs_Comp("End_Date"))
										%>
													<tr>
														<td align="<%=align_var%>"  class="itemList">
															<a href="" onclick="openScreen('<%=Screen_Id%>');return false"  class="itemList">
																<b>
																	<%=Competitor_Name%>
																</b></a>
														</td>
													</tr>
													<%rs_Comp.MoveNext
										loop
										rs_Comp.close
										set rs_Comp=nothing
									rs_Tours.MoveNext
								loop
							set rs_Tours=nothing%>
												</table>
											</div>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td height="10" nowrap></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
	</body>
</html>

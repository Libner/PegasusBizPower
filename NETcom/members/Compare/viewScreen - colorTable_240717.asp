<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%  

    	
	If Request.QueryString("TourId") <> nil Then
		if cint(trim(Request.QueryString("TourId")))>0 then
			Tour_Id = cint(Request.QueryString("TourId"))
		end if
	end if
	If Request.QueryString("ScreenId") <> nil Then
		if cint(trim(Request.QueryString("ScreenId")))>0 then
			Screen_Id = cint(Request.QueryString("ScreenId"))
		end if
	end if
	If Request.QueryString("CompetitorId") <> nil Then
		if cint(trim(Request.QueryString("CompetitorId")))>0 then
			Competitor_Id = cint(Request.QueryString("CompetitorId"))
		end if
	end if

	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 32 Order By word_id"				
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
	  
	  sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	  set rsbuttons = con.getRecordSet(sqlstr)
	  If not rsbuttons.eof Then
		arrButtons = ";" & rsbuttons.getString(,,",",";")		
		arrButtons = Split(arrButtons, ";")
	  End If
	  set rsbuttons=nothing	 	  
%>
<!DOCTYPE html>
<html>
	<head>
		<!-- #include file="../../title_meta_inc.asp" -->
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
		font-size:11pt;
    padding: 2px 1em 2px .1em;
}
 .ui-menu-item-wrapper:hover {
		font-size:11pt;
		font-weight:600;
    color:navy;
    border:0px;
		background-color:rgba( 0, 0, 0, 0.15 );
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
				$("#blockSearchResults").html(results)
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
          var logo
          logo=''
          /*added by Mila to change competitor title*/
          itemId=$(this).parent().prop('id')
           if (itemId=='selTour')
				textForSearch=text
		  if (itemId=='selCompetitor')
		  {		
				var TermsForSearch=getTermsForSearch($( this ).val())				
				textForSearch=text + ' ' +TermsForSearch				
				logo=getImgLogo($( this ).val())
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
 /*
 $(document).ready(function () {
    //alert('uu')
   // alert($("[id*='paramDesc_']").first().attr('id'))
	
	$("[id*='paramDesc_']").on("mouseover",function() {
	alert($(this).height())
	alert($(this).parent().find("[id*='paramDescAll_']").height())
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
*/
function openDesc(el)
{
	/*$(el).css('height','auto');
	$(el).parent().find("div[id*='iconDescOpen_']").hide()
	$(el).parent().find("a[id*='iconDescClose_']").show()
	*/
	if(el.height()<$(el).parent().find("[id*='paramDescAll_']").height())
	{
		if(el.parent().find("[id*='paramDescAll_']").css('display')=='none')
		{	el.parent().find("[id*='paramDescAll_']").css('display','block')}
	}
}
function hideDesc(el)
{
	/*$(el).css('height','28px');
	$(el).parent().find("div[id*='iconDescOpen_']").show()
	$(el).parent().find("a[id*='iconDescClose_']").hide()*/	
	
		$(el).css('display','none');
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
function getImgLogo(competitorId)
{
	ImgLogo=""
	$.ajax({
		type: "POST",
		url:"getCompetitorLogo.asp",
		data: {competitorId:competitorId},
		success: function(msg) {
			if(msg != "")
			{
				ImgLogo=" " + msg
			}
		},
		async:   false
	});
	return ImgLogo
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

 $(document).ready(function () {
  $(".ui-menu-item-wrapper").each(function() {
         $(this).on( "click", function() {
          alert('66')
          })   
}) 
})   
							</script>
	</head>
	<body class="body_admin">
		<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table4">
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
				<td class="page_title" align="<%=align_var%>" valign="middle" width="100%">���� 
					������</td>
			</tr>
			<tr>
				<td align="center" valign="top">
					<table border="0" cellpadding="0" cellspacing="0" width="100%" ID="Table8">
						<tr>
							<td height="25" nowrap></td>
						</tr>
						<tr>
							<td width="100%" align="center">
								<table border="0" cellspacing="0" cellpadding="0" align="center" ID="Table1">
									<tr>
										<td align="center">
											<div class="plata_search">
												<table cellspacing="0" cellpadding="5" align="center" border="0" ID="Table3" dir="ltr"
													width="800">
													<tr>
														<td height="5" colspan="2"></td>
													</tr>
													<!--site Competitor-->
													<%	
											If Len(Screen_Id) > 0 Then
												sqlstr="Select Tour_Id, Compare_Screens.Competitor_Id, Start_Date, End_Date From Compare_Screens  Where Screen_Id = " & Screen_Id
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
														<TD align="<%=align_var%>" dir="<%=dir_obj_var%>" width="100%">
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
																if selTour_Id =Tour_Id then
																	Tour_Name = trim(rssub("Tour_Name"))
																end if
																%>
																	<option value="<%=selTour_Id%>" <%if selTour_Id =Tour_Id then%>selected<%end if%>><%=selCategory_Name%>&nbsp;-&nbsp;<%=selTour_Name%></option>
																	<%
																rssub.movenext
															loop
															set rssub=Nothing%>
																</Select>
															</div>
														</TD>
														<TD align="<%=align_var%>" dir="<%=dir_obj_var%>" class="field_search"><b>&nbsp;�� 
																����&nbsp;</b></TD>
													</tr>
													<tr>
														<TD align="<%=align_var%>" dir="<%=dir_obj_var%>">
															<div class="ui-widget" style="width:auto">
															<%
															
														'get Competitor' names 
														sqlstr="SELECT distinct m.Competitor_Id,m.Competitor_Name, min(m.ord)"
														sqlstr=sqlstr & " from ( select row_number() OVER (ORDER BY Competitor_Name,End_Date desc, Start_Date desc) as ord,"
														sqlstr=sqlstr & " Compare_Screens.Competitor_Id,Competitor_Name"
														sqlstr=sqlstr & " from Compare_Screens  inner join Compare_Competitors on Compare_Competitors.Competitor_Id=Compare_Screens.Competitor_Id "
														sqlstr=sqlstr & " where  "
														if Tour_Id<>0 then
															sqlstr=sqlstr & " Tour_Id=" & Tour_Id & " and "
														end if
														sqlstr=sqlstr & " datediff(dd,Compare_Screens.[Start_Date],getdate())>=0 "
														sqlstr=sqlstr & " and datediff(dd,Compare_Screens.[End_Date],getdate())<=0 "
														sqlstr=sqlstr & " ) as m"
														sqlstr=sqlstr & " group by  m.Competitor_Id,m.Competitor_Name"
														sqlstr=sqlstr & " order by  min(m.ord)"
														
														
														
															'sqlstr="SELECT Compare_Screens.Competitor_Id,Competitor_Name from Compare_Screens "
															'sqlstr=sqlstr & " inner join Compare_Competitors on Compare_Competitors.Competitor_Id=Compare_Screens.Competitor_Id  where"
															'if Tour_Id<>0 then
															'sqlstr=sqlstr & " Tour_Id=" & Tour_Id & " and "
															'end if
															'sqlstr=sqlstr & "  datediff(dd,Compare_Screens.[Start_Date],getdate())>=0 and datediff(dd,Compare_Screens.[End_Date],getdate())<=0 "
															'sqlstr=sqlstr & "  order by Competitor_Name,End_Date desc, Start_Date desc"
															
															set rssub = con.getRecordSet(sqlstr)
															'response.write (sqlstr)
															%>
																<Select id="selCompetitor" name="selCompetitor" dir="<%=dir_obj_var%>">
																	<option value="0"></option>
																	
															<%do while not rssub.eof															
																selCompetitor_Id= cint(rssub("Competitor_Id"))																
																selCompetitor_Name = trim(rssub("Competitor_Name"))															
																if selCompetitor_Id =Competitor_Id then
																	Competitor_Name = trim(rssub("Competitor_Name"))
																end if
																%>
																	<option value="<%=selCompetitor_Id%>"  <%if selCompetitor_Id =Competitor_Id then%>selected<%end if%>><%=selCompetitor_Name%></option>
																	<%
																rssub.movenext
															loop
															set rssub=Nothing%>
																</Select>
															</div>
														</TD>
														<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>" class="field_search"><b>&nbsp;�� 
																�����&nbsp;</b></td>
													</tr>
													<tr>
														<td height="5" colspan="2" nowrap></td>
													</tr>
												</table>
											</div>
										</td>
									</tr>
									<tr>
										<td height="35" nowrap></td>
									</tr>
									<tr>
										<td width="100%">
											<div id="blockSearchResults" class="plata_results" <%if Screen_Id=0 then%>style="display:none"<%end if%>>
												<%if Competitor_Id>0 then
												sqlstr= "Select Competitor_Logo from  Compare_Competitors  whereCompetitor_Id=" & selCompetitor_Id 
			set rs_Comp = con.GetRecordSet(sqlstr)
			if not rs_Comp.eof then
				Competitor_Logo=rs_Comp("Competitor_Logo")
				if Competitor_Logo<>"" then
					set fso=server.CreateObject("Scripting.FileSystemObject") ' open a new FILE object!
					fileString= Server.MapPath("../../../download/competitors/"& Competitor_Logo ) 'deleting of existing file
					if not fso.FileExists(fileString) then
						Competitor_Logo=""
					end if
				end if	
			end if
			rs_Comp.close
			set rs_Comp=nothing
												
												%>
												<table cellspacing="1" cellpadding="2" width="100%" align="center" border="0" ID="Table3">
													<tr>
		<td  dir="<%=dir_obj_var%>"  align="<%=align_var%>">
			<table border="0" width="100%" cellpadding="0" cellspacing="0" dir="ltr" style="padding-right:10px" ID="Table2">
				<tr>
					<td colspan="2" align="<%=align_var%>" dir="<%=dir_obj_var%>"  class="tourTitle" style="color:navy"><%=Tour_Name%></td>
				</tr>
				<tr>
					<td colspan="2" height="5"></td>
				</tr>
				<tr>
					<td  dir="<%=dir_obj_var%>" width="100%" align="<%=align_var%>" class="tourTitle"><%=Competitor_Name%></td>
					<td align="<%=align_var%>"><%if Competitor_Logo<>"" then%>
						<img src="../../../download/competitors/<%=Competitor_Logo%>" border="0"  style="margin-left:10px;max-height:40px">
						<%end if%>
					</td>
				</tr>
				<tr>
					<td colspan="2" height="5"></td>
				</tr>
			
			</table>
		</td>
	</tr>
													<tr>
														<td align="right" nowrap bgcolor="#ffffff" valign="top" colspan="2">
															<table border="0" cellpadding="7" cellspacing="3" ID="Table6" dir="<%=dir_obj_var%>" width="100%" style="border-collapse:separate">
																<tr>
																	<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" class="thParam" >
																		������ �����</th>
																	<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" class="thParam" style="background-color:#99cc00;">
																		�����</th>
																	<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" class="thParam" style="background-color:#ffcc33;">
																		<!--�����--><div id="competitorTitle"><%=Competitor_Name%></div>
																	</th>
																	<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" class="thParam" >
																		&nbsp;�����?&nbsp;</th>
																	<th  dir="<%=dir_obj_var%>" align="<%=align_var%>"class="thParam" >
																		���� �� ������</th>
																</tr>
																<%'get parameters
												dim numParameters
												numParameters=0
												dim numFixedParameters
												numFixedParameters=0
													sqlstr="Select  Parameter_Id, Parameter_Name, Field_Type, isFixedName,Parameter_Icon from Compare_Parameters order by Parameter_Order"
													set rssub = con.getRecordSet(sqlstr)
													
												isVariabledParameters=false
												do while not rssub.eof
													numParameters=numParameters + 1
													
													Parameter_Id = cint(rssub("Parameter_Id"))
													Parameter_Name = trim(rssub("Parameter_Name"))
													Field_Type = cint(rssub("Field_Type"))
													isFixedName = cbool(rssub("isFixedName"))
													Parameter_Icon=  trim(rssub("Parameter_Icon"))
													if Screen_Id>0 then
														sqlstr="SELECT  Parameter_Value, Parameter_Advantages, Advantages_Description "
														sqlstr=sqlstr & " FROM  Compare_Parameters_ToScreen where Screen_Id=" & Screen_Id & " and Parameter_Id=" & Parameter_Id
														set rssubSCR = con.getRecordSet(sqlstr)
														if not	rssubSCR.eof then
															Parameter_Value_Screen=	trim(rssubSCR("Parameter_Value"))
															Parameter_Advantages=	trim(rssubSCR("Parameter_Advantages"))	
															isAdvantages=cbool(Parameter_Advantages)
															Advantages_Description=	trim(rssubSCR("Advantages_Description"))
															isParameterInScreen=true												
														else
															Parameter_Value_Screen=""	
															Parameter_Advantages=""	
															isAdvantages=true	
															Advantages_Description=""										
															isParameterInScreen=false
														end if		
														rssubSCR.close
														set rssubSCR =nothing	
													else								
														isParameterInScreen=false
													end if
													Parameter_Value_Tour=""
														Parameter_Name_Variabled=""
													if Tour_Id>0  then
														sqlstr="SELECT  Parameter_Value, Parameter_Name_Variabled  FROM  Compare_Parameters_ToTour where Tour_Id=" & Tour_Id & " and Parameter_Id=" & Parameter_Id
														set rssubTR = con.getRecordSet(sqlstr)	
														if not	rssubTR.eof then
															Parameter_Value_Tour=	trim(rssubTR("Parameter_Value"))
															if not isFixedName then
																Parameter_Name_Variabled=	trim(rssubTR("Parameter_Name_Variabled"))
															end if
														end if
														rssubTR.close
														set rssubTR =nothing			
													end if
													%>
																<%if Parameter_Value_Screen<>"" or Advantages_Description<>"" then%>
																<tr style="background-color:#f5f5f5" onmouseover="this.style.backgroundColor ='#e5e5e5'"
																	onmouseout="this.style.backgroundColor='#f5f5f5'">
																	<td>
																		<table width="100%" cellpadding="0" cellspacing="0">
																			<tr>
																				<td  class="paramName"align="right" dir="rtl" <%if Parameter_Advantages<>"" and isAdvantages then%> style="color:green"<%else%> style="color:red"<%end if%>>
																					<%=Parameter_Icon%>
																				</td>
																				<td class="paramName" align="right" dir="rtl">
																					<%if isFixedName then%>
																					<%=Parameter_Name%>
																					<%else%>
																					<%=Parameter_Name_Variabled%>
																					<%end if%>
																				</td>
																			</tr>
																		</table>
																	</td>
																	<td class="paramVal" align="<%=align_var%>"  dir="<%=dir_obj_var%>"   <%if Parameter_Advantages<>"" and isAdvantages then%> style="background-color:#f1ffef"<%else%> style="background-color:#fff0e2"<%end if%>>
																		<%=Parameter_Value_Tour%>
																	</td>
																	<td class="paramVal" align="<%=align_var%>"  dir="<%=dir_obj_var%>"  style="background-color:#fffff1">
																		<%=Parameter_Value_Screen%>
																	</td>
																	<td class="paramVal" align="center"  <%if Parameter_Advantages<>"" and isAdvantages then%> style="background-color:#e5ffe2"<%else%> style="background-color:#fff0e2"<%end if%>>
																		<%if Parameter_Advantages<>"" and isAdvantages then%>
																		<i class="fa fa-thumbs-up fa-flip-horizontal" aria-hidden="true" style="font-size:24px;color:green">
																		</i>
																		<%end if%>
																		<%if Parameter_Advantages<>"" and not isAdvantages then%>
																		<i class="fa fa-thumbs-down fa-flip-horizontal" aria-hidden="true" style="font-size:24px;color:red">
																		</i>
																		<%end if%>
																	</td>
																	<td class="paramDesc" align="<%=align_var%>"  dir="<%=dir_obj_var%>" >
																		<!--<div id="Desc_<%=Parameter_Id%>" style="position:relative;max-width:400px;min-height:28px;height:28px;overflow-y:hidden;white-space:normal;padding-left:10px"
																	onmouseover="openDesc(this)"><%=breaks(Advantages_Description)%>
																		<%if Advantages_Description<>"" then%>
																		<div id="iconDescOpen_<%=Parameter_Id%>" style="position:absolute;bottom:0px;left:0px;margin-bottom:3px"><i  class="fa fa-chevron-circle-down" aria-hidden="true" style="color:rgba( 0, 0, 0, 0.3 )"></i></div>
																		<%end if%>
																		<%if Advantages_Description<>"" then%>
																		<a id="iconDescClose_<%=Parameter_Id%>" style="position:absolute;bottom:0px;left:0px;display:none" href="" onclick="hideDesc($(this).parent());return false"><i  class="fa fa-chevron-circle-up" aria-hidden="true" style="color:rgba( 0, 0, 0 )"></i></a>
																		<%end if%>
																</div>
																-->
																		<div style="position:relative">
																			<div id="paramDesc_<%=Parameter_Id%>" 
																	style="position:relative;max-width:400px;min-height:28px;height:28px;overflow-y:hidden;white-space:normal;padding-left:10px;z-index:1"
																	onmouseover="openDesc($(this))"><%=breaks(Advantages_Description)%>
																			</div>
																			<%if Advantages_Description<>"" then%>
																			<div id="paramDescAll_<%=Parameter_Id%>" class="plata_description"
																	style="position:absolute;width:420px;height:auto;top:-10px;left:-10px;white-space:normal;padding-left:10px;display:none;z-index:1000"
																	onmouseout="hideDesc($(this))"><%=breaks(Advantages_Description)%>
																			</div>
																			<%end if%>
																		</div>
																	</td>
																</tr>
																<%end if%>
																<%rssub.movenext
												loop
												set rssub=Nothing
												
												%>
															</table>
														</td>
													</tr>
												</table>
												<%end if%>
											</div>
										</td>
									</tr>
									<tr>
										<td height="10"></td>
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

<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%session.LCID=1037

	If Request.QueryString("deleteId") <> nil Then
		delId = trim(Request.QueryString("deleteId"))
		If Len(delId) > 0 Then
			con.executeQuery("Delete From Compare_Screens Where Screen_Id = " & delId & ";Delete From Compare_Parameters_ToScreen Where Screen_Id = " & delId)
		End If
		'!!!!!!!!!delete from screens
		
		'Response.Redirect "ScreensAdmin.asp"
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
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
			<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
			<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
				<link rel="stylesheet" href="/resources/demos/style.css">
					<script src="https://ajax.googleapis.com/ajax/libs/webfont/1.4.7/webfont.js"></script>
					<script src="http://ajax.googleapis.com/ajax/libs/jquery/1.6.4/jquery.min.js"></script>
<style>
  .custom-sSeries {
    position: relative;
    display: inline-block; 
  }
  .custom-sSeries-toggle {
    position: absolute;
    top: 0;
    bottom: 0;
    margin-left: -1px;
    padding: 0; 
    height: 32px;
    width: auto;
  }
  .custom-sSeries-input {
    margin: 0;
     /*width:58px;*/ 
    direction:rtl;
    padding: 2px 5px;
     height:26px;
  }
 .ui-button, .ui-button:link, .ui-button:visited, .ui-button:hover, .ui-button:active {
    text-decoration: none;
    width:26px;
    /* height:32px; bluto*/
     height:26px;
}
  .ui-widget {
    font-family: Arial;
    font-size: 15px;
    font-weight: bold;
    color:navy;
    direction:rtl;
}

.ui-menu .ui-menu-item-wrapper {
    position: relative;
    font-size:15px;
    padding: 2px 1em 2px .1em;
     width:600px;
}
 
.ui-widget.ui-widget-content {
    border: 1px solid #c5c5c5;
    background:#ffffff;
     max-height: 250px;
     /*width:58px;*/
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
            //$(this).val('');
            });
 
        this._on( this.input, {
          autocompleteselect: function( event, ui ) {
            ui.item.option.selected = true;
            this._trigger( "select", event, {
              item: ui.item.option
            });
           
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
          .addClass( "custom-sSeries-toggle ui-corner-right" )
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
            return {
              label: text,
              value: text,
              option: this
            };
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
		url:"getScreensBySearchTerms.asp",
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

function addScreenBySearch()
{
	tourId=$("#selTour").val();
	competitorId=$("#selCompetitor").val();
	document.location.href="addScreen.asp?TourId=" + tourId + "&CompetitorId=" + competitorId + "&sTourId=" + tourId + "&sCompetitorId=" + competitorId
}
					</script>
					<SCRIPT LANGUAGE="javascript">
	function checkDelete()
	{

		<%
			If trim(lang_id) = "1" Then
				str_confirm = "?��� ������ ����� �� ������"
			Else
				str_confirm = "Are you sure want to delete the Competitor ?"
			End If   
		%>
		
		return window.confirm("<%=str_confirm%>");
		
	}
					</SCRIPT>
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
					<%numOfLink =1%>
					<!--#include file="../../top_in.asp"-->
				</td>
			</tr>
			<tr>
				<td class="page_title" align="<%=align_var%>" valign="middle" width="100%">���� 
					������</td>
			</tr>
			<tr>
				<td align="<%=align_var%>" valign=top align=center>
					<table cellpadding="0" cellspacing="0" width="100%" border="0" ID="Table2">
						<tr>
							<td>
								<table cellspacing="0" cellpadding="0" align="center" border="0" ID="Table6">
									<tr>
										<td height="15" nowrap></td>
									</tr>
									<tr>
										<td bgcolor="DarkGray">
											<table cellspacing="0" cellpadding="0" align="center" border="0" ID="Table7">
												<tr>
													<td colspan="2" align="center" class="title_form" style="background-color:DarkGray">����� 
														���� ������</td>
												</tr>
												<tr>
													<td bgcolor="#ffffff">
														<table cellspacing="1" cellpadding="2" align="center" border="0" bordercolor="#ffffff"
															ID="Table5">
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
											
											else	
											'back to search results		
												If Request.querystring("stourId")<>"" Then
													Tour_Id = cint(Request.querystring("stourId"))
												End If
														
												If Request.querystring("scompetitorId")<>"" Then
													Competitor_Id = cint(Request.querystring("scompetitorId"))
												End If
											End If
										%>
															<tr>
																<TD align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_form" style="padding-left:30px">
																	<div class="ui-widget" style="width:100%">
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
																			<option value="<%=selTour_Id%>" <%if selTour_Id =Tour_Id then%>selected<%end if%>><%=selTour_Name%></option>
																			<%
																rssub.movenext
															loop
															set rssub=Nothing%>
																		</Select>
																	</div>
																</TD>
																<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>" class="title_form"><b>&nbsp;�� 
																		����&nbsp;</b></td>
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
																<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>" class="title_form"><b>&nbsp;�� 
																		�����&nbsp;</b></td>
															</tr>
														</table>
													</td>
												</tr>
												<tr>
													<td colspan="2" bgcolor="DarkGray" height="10"></td>
												</tr>
												<tr>
													<td colspan="2" bgcolor="DarkGray" align="center"><a class="button_edit_1" style="width:100px;" href="" onclick="addScreenBySearch();return false"><span id="Span4" name="word6"><!--����� �����-->����� ��� ������</span></a></td>
												</tr>
												<tr>
													<td colspan="2" bgcolor="DarkGray" height="10"></td>
												</tr>
											</table>
										</td>
									</tr>
									<tr>
										<td height="20"></td>
									</tr>
									<tr>
										<td width="100%" valign="top" align="right">
											<div id="blockSearchResults">
												<table width="100%" cellspacing="1" cellpadding="2" border="0" bgcolor="#ffffff" ID="Table3">
													<tr>
														<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="word3" name="word3"><!--�����--><%=arrTitles(3)%></span>&nbsp;</td>
														<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="word4" name="word4"><!--�����--><%=arrTitles(4)%></span>&nbsp;</td>
														<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="Span1" name="word4">�� �����</span>&nbsp;</td>
														<td width="80" nowrap align="center" class="title_sort">&nbsp;<span id="Span2" name="word4">������</span>&nbsp;</td>
														<td width="100%"  align="<%=align_var%>" class="title_sort">&nbsp;<span id="word5" name="word5"><!--�����-->�����</span>&nbsp;</td>
													</tr>
													<%		'get tours' names from Pegasusisrael site
									sqlstr="SELECT distinct Compare_Screens.Tour_Id, " & pegasusDBName & ".dbo.Tours.Tour_Name, "
									sqlstr=sqlstr &  pegasusDBName & ".dbo.Tours_Categories.Category_Name,  " & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Name, "
									sqlstr=sqlstr &  pegasusDBName & ".dbo.Tours_Categories.Category_Order," & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Order, " & pegasusDBName & ".dbo.Tours.Tour_Order"
									sqlstr=sqlstr & " FROM   Compare_Screens INNER JOIN "
									sqlstr=sqlstr &  pegasusDBName & ".dbo.Tours ON Compare_Screens.Tour_Id = " & pegasusDBName & ".dbo.Tours.Tour_Id "
									sqlstr=sqlstr & " INNER JOIN " & pegasusDBName & ".dbo.Tours_Categories ON " & pegasusDBName & ".dbo.Tours.Category_Id = " & pegasusDBName & ".dbo.Tours_Categories.Category_Id INNER JOIN "
										sqlstr=sqlstr & pegasusDBName & ".dbo.Tours_SubCategories ON " & pegasusDBName & ".dbo.Tours.SubCategory_Id = " & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Id "
										if Tour_Id<>0 or Competitor_ID <>0 then
											swhere =  " where "
											if Tour_Id<>0 then
											swhere=swhere & " Compare_Screens.Tour_Id =" & Tour_Id
											end if
											if Competitor_ID<>0 then
												if swhere <> " where " then
												swhere=swhere & " AND "
												end if
												swhere=swhere & " Compare_Screens.Competitor_ID =" & Competitor_ID
											end if
											sqlstr=sqlstr & swhere
										end if
										sqlstr=sqlstr & "order by " & pegasusDBName & ".dbo.Tours_Categories.Category_Order," & pegasusDBName & ".dbo.Tours_SubCategories.SubCategory_Order, " & pegasusDBName & ".dbo.Tours.Tour_Order"
									set rs_Tours = con.GetRecordSet(sqlstr)
									do while not rs_Tours.eof
    										TourId = trim(rs_Tours("Tour_Id"))
												Category_Name = trim(rs_Tours("Category_Name"))
												SubCategory_Name = trim(rs_Tours("SubCategory_Name"))
												Tour_Name = trim(rs_Tours("Tour_Name"))
											
										%>
													<tr>
														<th nowrap colspan="2" align="center" bgcolor="#e6e6e6">
															<a class="button_edit_1" style="width:100;" href="addScreen.asp?tourId=<%=Tour_Id%>">
																<span id="Span3" name="word6"><!--����� �����-->����� ��� ������</span></a></th>
														<th colspan="3" align="<%=align_var%>" bgcolor="#e6e6e6">
															<b><%=Tour_Name%>
															</b>&nbsp;</th>
													</tr>
													<%
											sqlstr= "Select Screen_Id,Competitor_Name,Start_Date,End_Date  from Compare_Screens "
											sqlstr=sqlstr & "  inner join Compare_Competitors on Compare_Competitors.Competitor_Id=Compare_Screens.Competitor_Id where Tour_Id=" & TourId 
											if Competitor_ID<>0 then
												sqlstr=sqlstr & " AND  Compare_Competitors.Competitor_ID =" & Competitor_ID
											end if
											sqlstr=sqlstr & " order by Competitor_Name,End_Date desc, Start_Date desc"
											set rs_Comp = con.GetRecordSet(sqlstr)
											do while not rs_Comp.eof
    											Screen_Id = trim(rs_Comp("Screen_Id"))
												Competitor_Name = trim(rs_Comp("Competitor_Name"))
												Start_Date = trim(rs_Comp("Start_Date"))
												End_Date = trim(rs_Comp("End_Date"))
										%>
										<tr>
											<td align="center" bgcolor="#e6e6e6" nowrap><a href="screensAdmin.asp?deleteId=<%=Screen_Id%><%if Tour_ID<>"" then%>&stourId=<%=Tour_ID%><%end if%><%if competitor_Id<>"" then%>&scompetitorId=<%=competitor_ID%><%end if%>" ONCLICK="return checkDelete(<%=countComp %>)"><IMG SRC="../../images/delete_icon.gif" BORDER="0"></a></td>
											<td align="center" bgcolor="#e6e6e6" nowrap><a href="addScreen.asp?TourID=<%=Tour_Id%>&screenID=<%=Screen_Id%><%if Tour_ID<>"" then%>&stourId=<%=Tour_ID%><%end if%><%if competitor_Id<>"" then%>&scompetitorId=<%=competitor_ID%><%end if%>"><IMG SRC="../../images/edit_icon.gif" BORDER=0 alt="<%=alt_edit%>"></a></td>
											<td align="<%=align_var%>" bgcolor="#e6e6e6" nowrap <%if datediff("d",date(),Start_Date)>0 or datediff("d",date(),End_Date)<0 then%>style="color:#a9a9a9"<%end if%>><b><%= cdate(End_Date)%></b>&nbsp;</td>
											<td align="<%=align_var%>" bgcolor="#e6e6e6" nowrap <%if datediff("d",date(),Start_Date)>0 or datediff("d",date(),End_Date)<0 then%>style="color:#a9a9a9"<%end if%>><b><%=cdate(Start_Date)%></b>&nbsp;</td>
											<td align="<%=align_var%>" bgcolor="#e6e6e6" width="100%" <%if datediff("d",date(),Start_Date)>0 or datediff("d",date(),End_Date)<0 then%>style="color:#a9a9a9"<%end if%>><b><%=Competitor_Name%></b>&nbsp;</td>
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
							<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
								<table cellpadding="1" cellspacing="1" width="100%" ID="Table4">
									<tr>
										<td align="<%=align_var%>" height="18" nowrap></td>
									</tr>
									<%if Request.querystring("stourId")<>"" or Request.querystring("scompetitorId")<>"" then
										squery="?"
										if Request.querystring("stourId")<>"" then
											squery=squery & "stourId=" & Request.querystring("stourId")
										end if
										if Request.querystring("scompetitorId")<>"" then
											if squery<>"?" then
											squery=squery & "&"
											end if
											squery=squery & "scompetitorId=" & Request.querystring("scompetitorId")
										end if
									end if
									%>
									<tr>
										<td nowrap align="center"><a class="button_edit_1" style="width:100;" href="addScreen.asp<%=squery%>"><span id="word6" name="word6"><!--����� �����-->����� ��� ������</span></a></td>
									</tr>
									<tr>
										<td nowrap align="center"><a class="button_edit_1" style="width:100;background-color:DarkGray" href="screensAdmin.asp"><span id="Span5" name="word6"><!--����� �����-->��� ���</span></a></td>
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

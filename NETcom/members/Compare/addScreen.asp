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
	
    If Request.QueryString("add") <> nil Then
		if Request.form("selTour")<>"" and Request.form("selCompetitor")<>"" and Request.form("Start_Date")<>"" and Request.form("End_Date")<>""then 'only joined to tour and Competitor
			'insert/update=====================================
			Tour_Id=cint(Request.form("selTour"))
			Competitor_Id=cint(Request.form("selCompetitor"))
			Start_Date=cdate(Request.form("Start_Date"))
			End_Date=cdate(Request.form("End_Date"))
			
			if Tour_Id>0 and Competitor_Id>0 then 'only joined to tour and Competitor
				isValid=true
				'check if exists screen for the same tour and the same Competitor in the same interval:  the new and old dates are overlaped
				prevScreen_Id=0
				sqlstr= "set dateformat dmy;select Screen_Id from Compare_Screens where Tour_Id=" & Tour_Id & " and Competitor_Id=" & Competitor_Id & " and "
				sqlstr= sqlstr & " not ( datediff(dd,End_Date,'" & Start_Date & "')>0 or datediff(dd,Start_Date,'" & End_Date & "')<0)"
				If Screen_Id > 0 Then 
					sqlstr= sqlstr & " and Screen_Id<>" & Screen_Id
				end if
				'23/08/2017','23/07/2017'<0
				set rs_tmp = con.getRecordSet(sqlstr)
				if not rs_tmp.eof then
						prevScreen_Id = rs_tmp("Screen_Id")
						isValid=false
				end if
				set rs_tmp = Nothing
				if isValid then 'dates are valid
					'get parameters
					If Screen_Id = 0 Then ' add Screen
						sqlstr = "Insert into Compare_Screens (Tour_Id, Competitor_Id, Start_Date, End_Date,Modify_Date,USER_ID,USER_NAME) values ("
						sqlstr=sqlstr &  Tour_Id & ","  & Competitor_Id & ",'"  & Start_Date & "','"  & End_Date & "',getDate()," & request.Cookies("bizpegasus")("UserId") & ",'" & request.Cookies("bizpegasus")("UserName") & "')"
						'response.Write("<br>1: " & sqlstr)
						con.executeQuery(sqlstr)
						
						sqlstr="SELECT top 1 Screen_Id  from Compare_Screens  order by Screen_Id desc"
						set rs_tmp = con.getRecordSet(sqlstr)
						if not rs_tmp.eof then
								Screen_Id = rs_tmp("Screen_Id")
						end if
						set rs_tmp = Nothing
						'response.Write("<br>Screen_Id after insert: " & Screen_Id)
						
					Else ' update Screen
						sqlstr="Update Compare_Screens set Tour_Id = " & Tour_Id & ""
						sqlstr=sqlstr & ",Start_Date = '" & Start_Date & "'"
						sqlstr=sqlstr & ",End_Date = '" & End_Date & "'"
						sqlstr=sqlstr & ",Modify_Date = getDate()"
						sqlstr=sqlstr & ",USER_ID = " & request.Cookies("bizpegasus")("UserId") 
						sqlstr=sqlstr & ",USER_NAME = '" & request.Cookies("bizpegasus")("UserName") & "'"
						sqlstr=sqlstr & " Where Screen_Id = " & Screen_Id
					'response.Write("<br>2: " & sqlstr)
							con.executeQuery(sqlstr)
					End If
					If Screen_Id > 0 Then ' add parameters
						sqlstr="delete from Compare_Parameters_ToScreen  Where Screen_Id = " & Screen_Id
						con.executeQuery(sqlstr)
						sqlstr="delete from Compare_Parameters_ToTour  Where Tour_Id = " & Tour_Id
						con.executeQuery(sqlstr)	
						
						sqlstr="Select  Parameter_Id, isFixedName from Compare_Parameters order by Parameter_Order"
						set rssub = con.getRecordSet(sqlstr)
						do while not rssub.eof							
							Parameter_Id = cint(rssub("Parameter_Id"))
							isFixedName = cbool(rssub("isFixedName"))
							if not isFixedName then 'add variable name to params for tour
								if request.Form("Parameter_Name_Variabled_" & Parameter_Id)<>"" then
									sqlstr = "Insert into Compare_Parameters_ToTour (Tour_Id, Parameter_Id,Parameter_Name_Variabled, Parameter_Value) values ("
									sqlstr=sqlstr &  Tour_Id & ","  & Parameter_Id & ",'"  & sfix(request.Form("Parameter_Name_Variabled_" & Parameter_Id)) & "','"  & sfix(request.Form("Parameter_Value_Pegasus_" & Parameter_Id)) & "')"
					'response.Write("<br>3: " & sqlstr)
										con.executeQuery(sqlstr)
								end if
							else 'fixed
									sqlstr = "Insert into Compare_Parameters_ToTour (Tour_Id, Parameter_Id,Parameter_Name_Variabled, Parameter_Value) values ("
									sqlstr=sqlstr &  Tour_Id & ","  & Parameter_Id & ",NULL,'"  & sfix(request.Form("Parameter_Value_Pegasus_" & Parameter_Id)) & "')"
					'response.Write("<br>4: " & sqlstr)
										con.executeQuery(sqlstr)
							end if 
							
							if request.Form("Parameter_Value_Screen_" & Parameter_Id)<>"" or request.Form("Parameter_Advantages_" & Parameter_Id)<>"" then
								if request.Form("Parameter_Advantages_" & Parameter_Id)="" then
									strAdvantages="NULL"
								else
									strAdvantages=request.Form("Parameter_Advantages_" & Parameter_Id)
								end if
								sqlstr = "Insert into Compare_Parameters_ToScreen (Screen_Id, Parameter_Id,Parameter_Value,Parameter_Advantages,Advantages_Description) values ("
								sqlstr=sqlstr &  Screen_Id & ","  & Parameter_Id & ",'"  & sfix(request.Form("Parameter_Value_Screen_" & Parameter_Id)) & "',"  & strAdvantages & ",'"  & sfix(request.Form("Advantages_Description_" & Parameter_Id)) & "')"
						'response.Write("<br>5: " & sqlstr)
								con.executeQuery(sqlstr)
							end if
							rssub.movenext
						loop
						'tour parameters
						
					end if
				end if
			End If
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
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">

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
  }
  .custom-sSeries-toggle {
    position: absolute;
    top: 0;
    bottom: 0;
    margin-left: -1px;
    padding: 0;   
     width:auto;
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
    width:16px;
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
            if ($(event.target).attr("id")=='input_selCompetitor')
				$("#competitorTitle").text(ui.item.label);
            /*added by Mila to change parameters for pegasus tour*/
            if ($(event.target).attr("id")=='input_selTour')
			{
				$("#tourTitle").text(ui.item.label);
				fillTourParameters(ui.item.option.value);
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
          if ( this.value && ( !request.term || matcher.test(text) ) )
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
  $(document).ready(function () {
    
	$("input[id*='Parameter_Value_Screen_']").keyup(function() {
		if ($(this).val() != "")
		{
			if($(this).parent().parent().find("[id*='Parameter_Advantages_1']").prop( "checked", false ) && $(this).parent().parent().find("[id*='Parameter_Advantages_0']").prop( "checked", false ))
			{
				$(this).parent().parent().find("[id*='tdAdvantages_1_']").css('background-color','#ccffcc')
					$(this).parent().parent().find("[id*='Parameter_Advantages_1']").prop( "checked", true );
			}
		}
	});
	 $("input[id*='Parameter_Advantages_1']").click(function () {
			
		if(this.checked )
		{
			$(this).parent().css('background-color','#ccffcc')
			if($(this).parent().parent().find("[id*='Parameter_Advantages_0']").prop( "checked"))
			{
				$(this).parent().parent().find("[id*='Parameter_Advantages_0']").prop( "checked", false );
				$(this).parent().parent().find("[id*='Parameter_Advantages_0']").parent().css('background-color','')
			}
		}
		else
		{
			$(this).parent().css('background-color','')
			//$(this).parent().parent().find("[id*='Parameter_Advantages_0']").prop( "checked", true );
			//$(this).parent().parent().find("[id*='Parameter_Advantages_0']").parent().css('background-color','#ff6666')
		}
	});
	$("input[id*='Parameter_Advantages_0']").click(function () {
				
			if(this.checked )
			{
				$(this).parent().css('background-color','#ff6666')
				if($(this).parent().parent().find("[id*='Parameter_Advantages_1']").prop( "checked"))
				{
					$(this).parent().parent().find("[id*='Parameter_Advantages_1']").prop( "checked", false );
					$(this).parent().parent().find("[id*='Parameter_Advantages_1']").parent().css('background-color','')
				}
			}
			else
			{
				$(this).parent().css('background-color','')
				//$(this).parent().parent().find("[id*='Parameter_Advantages_1']").prop( "checked", true );
				//$(this).parent().parent().find("[id*='Parameter_Advantages_1']").parent().css('background-color','#ccffcc')
			}
		});
           
      }) 
   
   //submit form=========================================   
		function checkForm()
		{
			var isFormValid
			isFormValid=true
			
			if($("#selTour").val()=='0')
			{
				isFormValid=false
				window.alert("!!נא לבחור טיול");
				$("#input_selTour").get(0).focus();;
				return false;
			}
			if($("#selCompetitor").val()=='0')
			{
				isFormValid=false
				window.alert("!!נא לבחור מתחרה");
				$("#input_selCompetitor").get(0).focus();
				return false;
			}
			if($("#Start_Date").val()=='' || $("#End_Date").val()=='')
			{
				isFormValid=false
				window.alert("!!נא לבחור תוקף מסך השוואה");
				if($("#Start_Date").val()=='')
						$("#Start_Date").get(0).focus();
				else
						$("#End_Date").get(0).focus();
				return false;
			}
			//alert(isFormValid)
			if ( isFormValid)
			{	// tour+competitor date validation 
						
				if(!isValidScreenDates())
				{
					isFormValid=false
					window.alert("מסך השוואה כבר קיים בתאריכים שהזנתה");
					$("#Start_Date").get(0).focus();
					return false;
				}	
			}	
			
			if ( isFormValid)
			{	
					//check if is entered name for variabled parameters
				$("input[id*='Parameter_Name_Variabled_']").each(function () {				
					if ($(this).val() == '' && ($(this).parent().parent().find("[id*='Parameter_Value_Pegasus_']").val()!='' || $(this).parent().parent().find("[id*='Parameter_Value_Screen_']").val()!=''))
					{
								isFormValid=false
								window.alert("!!נא הכנס שם הפרמטר ");
								$(this).parent().parent().find("[id*='Parameter_Name_Variabled_']").get(0).focus();
								return false;
					}
				});
			}	
			if ( isFormValid)
			{	
				//check if is entered both values for pegasus and competitor 
				isParamatersChoosen=false
			
				$("input[id*='Parameter_Value_Screen_']").each(function () {
					if ($(this).val() != '')
					{
						isParamatersChoosen=true
						/*
						if($(this).parent().parent().find("[id*='Parameter_Value_Pegasus_']").val()=='')
							{
								isFormValid=false
								window.alert("!!נא הכנס ערך הפרמטר לטיול של פגסוס ");
								$(this).parent().parent().find("[id*='Parameter_Value_Pegasus_']").get(0).focus();
								return false;
							}
						*/
					}
				});
				
				$("input[id*='Parameter_Value_Pegasus_']").each(function () {
					if ($(this).val() != '')
					{
						isParamatersChoosen=true
					}
				});
			}
			if ( isFormValid)
			{		//must choose at least one value 
				if (!isParamatersChoosen)
				{
					isFormValid=false
					window.alert("!!נא בחר פרמטרים להשוואה ");
					$("[id*='Parameter_Value_Screen_']:first").get(0).focus();
					return false;
				}
			}
			if (isFormValid)
			{
				document.form1.submit()	
				return true	
			}	
		}   
		
		function  isValidScreenDates()
		{
			screenId='<%=Screen_Id%>'
			tourId=$("#selTour").val()
			competitorId=$("#selCompetitor").val()
			startDate=$("#Start_Date").val()
			endDate=$("#End_Date").val()
			isValid=false
			$.ajax({
				type: "POST",
				url:"isScreenDatesValid.asp",
				data: {screenId:screenId, tourId: tourId, competitorId:competitorId, startDate:endDate, endDate:endDate},
				success: function(msg) {
					if(msg=="1")
					{
						isValid=true
					}
				},
				async:   false
			});
			return isValid
		}
		
		function  fillTourParameters(tourId)
		{
			$("input[id*='Parameter_Value_Pegasus_']").each(function () {
				$(this).val('')
			});
			
			$("input[id*='Parameter_Name_Variabled_']").each(function () {
				$(this).val('')
			});
			
			var strParamValues=getTourParametersValues(tourId)
			
			var arrParams=new Array
			arrParams=strParamValues.split("#|#") //array of strings each of them in format 'id~|~paramvalue~|~varaible param name', the array separeted with #|#
		
			for (i=0; i<arrParams.length; i++)
			{
				arrParam=new Array
				arrParam=arrParams[i].split("~|~")
				paramId=arrParam[0]
				paramVal=arrParam[1]
				paramVarName=arrParam[2]
				if($("#Parameter_Value_Pegasus_"+paramId))
				{
					$("#Parameter_Value_Pegasus_"+paramId).val(paramVal)
				}	
				if($("#Parameter_Name_Variabled_"+paramId))
				{
					$("#Parameter_Name_Variabled_"+paramId).val(paramVarName)
				}	
			}
		}
		function  getTourParametersValues(tourId)
		{
			paramsValues=""
				$.ajax({
				type: "POST",
				url:"getParametersValuesToTour.asp",
				data: {tourId: tourId},
				success: function(msg) {
					if(msg != "")
					{
						paramsValues=msg
					}
				},
				async:   false
			});
			return paramsValues
		}
				
		/*function  fillTourParameters()
		{
			$("input[id*='Parameter_Value_Tour_']").each(function () {
					paramId=$(this).attr("id").substring(('Parameter_Value_Tour_').length)
					paramVal=getTourParameterValue(paramId)
					$(this).val(paramVal)
				});
		}
		function  getTourParameterValue(paramId)
		{
			tourId=$("#selTour").val()
			paramVal=""
				$.ajax({
				type: "POST",
				url:"getParameterValToTour.asp",
				data: {tourId: tourId, paramId: paramId},
				success: function(msg) {
					if(msg != "")
					{
						paramVal=msg
					}
				},
				async:   false
			});
			return paramVal
		}*/
  </script>		
	</head>
	<body>
		<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>" ID="Table4">
			<tr>
				<td width=100% align="<%=align_var%>">
					<!--#include file="../../logo_top.asp"-->
				</td>
			</tr>
			<tr>
				<td width=100% align="<%=align_var%>">
					<%numOftab = 95%>
					<%numOfLink =1%>
<%topLevel2 = 101 'current bar ID in top submenu - added 03/10/2019%>
					<!--#include file="../../top_in.asp"-->
				</td>
			</tr>
			<tr>
				<td class="page_title" align="<%=align_var%>" valign="middle" width="100%" dir="rt"><%If Len(Screen_Id) > 0 Then%>עדכון<%Else%>הוספת<%End If%>&nbsp;מסך השוואה&nbsp;</td>
			</tr>
			<tr>
				<td align="<%=align_var%>" valign=top>
					<table border="0"  bgcolor="#E6E6E6" cellpadding="0" cellspacing="0" width="100%" dir="<%=dir_var%>" ID="Table8">
						
						<tr>
							<td height="15" nowrap></td>
						</tr>
						
						<tr>
							<td width="100%">
								<form name="form1" id="form1" action="addScreen.asp?add=1&TourID=<%=Tour_Id%>&screenID=<%=Screen_Id%><%if Request.querystring("stourId")<>"" then%>&stourId=<%=Request.querystring("stourId")%><%end if%><%if Request.querystring("scompetitorId")<>"" then%>&scompetitorId=<%=Request.querystring("scompetitorId")%><%end if%>" target="_self" method="post">
									<table dir="<%=dir_var%>" border="0" width="100%" cellspacing="0" cellpadding="0" align="center" ID="Table1">
										<tr>
											<td width="100%">
												<table width="480" cellspacing="1" cellpadding="2" align="center" border="0" ID="Table3">
													<!--site Competitor-->
													<%	
											Modify_USER_ID=0
											If Len(Screen_Id) > 0 Then
												sqlstr="Select Tour_Id, Compare_Screens.Competitor_Id, Start_Date, End_Date,Modify_Date,USER_ID,USER_NAME From Compare_Screens Where Screen_Id = " & Screen_Id
												set rssub = con.getRecordSet(sqlstr)
												If not rssub.eof Then
													Tour_Id = cint(rssub("Tour_Id"))
													Competitor_Id = cint(rssub("Competitor_Id"))
													Start_Date = cdate(rssub("Start_Date"))
													End_Date = cdate(rssub("End_Date"))
													if not isnull(rssub("Modify_Date")) then
														Modify_Date=cdate(rssub("Modify_Date"))
													end if
													if not isnull(rssub("USER_ID")) then
														Modify_USER_ID=rssub("USER_ID")
													end if
													if not isnull(rssub("USER_NAME")) then
														Modify_USER_NAME=rssub("USER_NAME")
													end if
												End If
												set rssub=Nothing
												if Modify_USER_ID>0 then 'get current user name
													sqlstr="Select FIRSTNAME,LASTNAME From USERS Where USER_ID = " & Modify_USER_ID
													set rssub = con.getRecordSet(sqlstr)
													If not rssub.eof Then
														Modify_USER_NAME= trim(rssub("FIRSTNAME")) & " " & trim(rssub("LASTNAME"))
													end if													
													set rssub=Nothing
												end if
											End If
											
										%>
										<tr>
											<td colspan="2" align="right" dir="rtl">
											<span style="font-size:11pt"><span style="font-weight:normal">עודכן</span> <%=Modify_Date%>&nbsp;<span style="font-weight:normal">ע''י</span> <%=Modify_USER_NAME%></span>
											</td>
										</tr>
						<tr>
							<td height="15" nowrap></td>
						</tr>
													<tr>
														<TD align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_form">
														<div class="ui-widget" style="width:100%">
														<Select id="selTour" name="selTour" dir="<%=dir_obj_var%>">
															<option value="0"></option>
															<%
															
														'get tours' names from Pegasusisrael site
															sqlstr="SELECT Tours.Tour_Id, Tours.Category_Id, Tours.SubCategory_Id, Tours.Tour_Name, Tours.Tour_Title, "
															sqlstr=sqlstr & " Tours_Categories.Category_Name,  Tours_SubCategories.SubCategory_Name, Series_Name FROM Tours INNER JOIN Tours_Categories ON Tours.Category_Id = Tours_Categories.Category_Id INNER JOIN "
															sqlstr=sqlstr & "  Tours_SubCategories ON Tours.SubCategory_Id = Tours_SubCategories.SubCategory_Id"
															sqlstr=sqlstr & " LEFT JOIN " &  Application("bizpowerDBName") & ".dbo.Series on Series.Series_Id=Tours.SeriasId "
															sqlstr=sqlstr & "  order by Tours_Categories.Category_Order,Tours_SubCategories.SubCategory_Order,Tours.Tour_Name"
															'sqlstr=sqlstr & "  Where Tour_Id = " & Tour_Id
															set rssub = conPegasus.Execute(sqlstr)
															do while not rssub.eof															
																selTour_Id= cint(rssub("Tour_Id"))
																selCategory_Name = trim(rssub("Category_Name"))
																selSubCategory_Name = trim(rssub("SubCategory_Name"))
																selTour_Name = trim(rssub("Tour_Name"))
																if  not rssub("Series_Name") is nothing then							
																	Series_Name = trim(rssub("Series_Name"))
																	selTour_Name=selTour_Name & " (" & Series_Name & ")"
																end if									
																if selTour_Id =Tour_Id then
																	Tour_Name = trim(rssub("Tour_Name"))
																end if
																%>
																<option value="<%=selTour_Id%>" <%if selTour_Id =Tour_Id then%>selected<%end if%>><%=selTour_Name%></option>
															<%
																rssub.movenext
															loop
															set rssub=Nothing%>
														</Select>
														</div>
														</TD>
														<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>" class="title_form"><b><font color="red">*</font>&nbsp;שם טיול&nbsp;</b></td>
													</tr>
													<tr>
														<TD align="<%=align_var%>" dir="<%=dir_obj_var%>" width="100%" class="title_form">
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
																if selCompetitor_Id =Competitor_Id then
																	Competitor_Name = trim(rssub("Competitor_Name"))
																end if
																%>
																<option value="<%=selCompetitor_Id%>"  <%if selCompetitor_Id =Competitor_Id then%>selected<%end if%>><%=selCompetitor_Name%> </option>
															<%
																rssub.movenext
															loop
															set rssub=Nothing%>
														</select>
</div>
</td>
														<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>" class="title_form"><b><font color="red">*</font>&nbsp;שם מתחרה&nbsp;</b></td>
													</tr>
													<tr>
														<td align="right" nowrap class="td_admin_4" valign="top">
															<table border="0" cellpadding="2" cellspacing="0" ID="Table2">
																<tr>
																	<TD align="<%=align_var%>" dir="<%=dir_obj_var%>"><a href="" onclick="cal1xx.select(document.getElementById('Start_Date'),'AsStart_Date','dd/MM/yyyy'); return false;"
																			id="AsStart_Date"><img src="../../images/calendar.gif" border="0" align="center"></a><input type="text" name="Start_Date" ID="Start_Date" size=20 value="<%=vFix(Start_Date)%>" readonly dir="<%=dir_obj_var%>"></TD>
																	<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>"><b><font color="red">*</font>&nbsp;מתאריך&nbsp;</b></td>
																</tr>
																<tr>
																	<TD align="<%=align_var%>" dir="<%=dir_obj_var%>"><a href="" onclick="cal1xx.select(document.getElementById('End_Date'),'AsStart_Date','dd/MM/yyyy'); return false;"
																			id="AsEnd_Date"><img src="../../images/calendar.gif" border="0" align="center"></a><input type="text" name="End_Date" ID="End_Date" size=20 value="<%=vFix(End_Date)%>" readonly dir="<%=dir_obj_var%>"></TD>
																	<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>"><b><font color="red">*</font>&nbsp;עד תאריך&nbsp;</b></td>
																</tr>
															</table>
														</td>
														<td align="<%=align_var%>" nowrap dir="<%=dir_obj_var%>"><b><font color="red">*</font>&nbsp;תוקף מסכי 
																השוואה&nbsp;</b></td>
													</tr>
													<tr>
														<td colspan="2" dir="<%=dir_obj_var%>" class="form_header"><div id="tourTitle"><%=Tour_Name%></div></td>
													</tr>
													<tr>
														<td align="right" nowrap bgcolor="#ffffff" valign="top" colspan="2">
															<table border="0" cellpadding="3" cellspacing="1" ID="Table6" dir="<%=dir_obj_var%>">
																<tr>
																	<th  dir="<%=dir_obj_var%>" align="<%=align_var%>">
																		השירות בטיול</th>
																	<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" style="background-color:#99cc00">
																		פגסוס</th>
																	<th  dir="<%=dir_obj_var%>" align="<%=align_var%>" style="background-color:#ffcc33">
																		<!--מתחרה--><div id="competitorTitle"><%=Competitor_Name%></div>
																	</th>
																	<th  dir="<%=dir_obj_var%>" align="<%=align_var%>">
																		יתרון?</th>
																	<th  dir="<%=dir_obj_var%>" align="<%=align_var%>">
																		הסבר על היתרון</th>
																</tr>
												<%'get parameters
												dim numParameters
												numParameters=0
												dim numFixedParameters
												numFixedParameters=0
													sqlstr="Select  Parameter_Id, Parameter_Name, Field_Type, isFixedName from Compare_Parameters order by Parameter_Order"
													set rssub = con.getRecordSet(sqlstr)
													
												isVariabledParameters=false
												do while not rssub.eof
													numParameters=numParameters + 1
													
													Parameter_Id = cint(rssub("Parameter_Id"))
													Parameter_Name = trim(rssub("Parameter_Name"))
													Field_Type = cint(rssub("Field_Type"))
													isFixedName = cbool(rssub("isFixedName"))
													if Screen_Id>0 then
														sqlstr="SELECT  Parameter_Value, Parameter_Advantages, Advantages_Description "
														sqlstr=sqlstr & " FROM  Compare_Parameters_ToScreen where Screen_Id=" & Screen_Id & " and Parameter_Id=" & Parameter_Id
														set rssubSCR = con.getRecordSet(sqlstr)
														if not	rssubSCR.eof then
															Parameter_Value_Screen=	trim(rssubSCR("Parameter_Value"))
															Parameter_Advantages=	trim(rssubSCR("Parameter_Advantages"))	
															if Parameter_Advantages<>"" then
																isAdvantages=cbool(Parameter_Advantages)
															else
																isAdvantages=true
															end if
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
													<%if not isVariabledParameters and isFixedName=false then 'end of fixed parameters and go to variabled parameters for defined screen
														isVariabledParameters=true%>
														<tr><th colspan="5" align="<%=align_var%>"  dir="<%=dir_obj_var%>" class="form_title">שירותים נוספים</th></tr>
													<%end if%>
																<tr>
																	<td align="<%=align_var%>"  dir="<%=dir_obj_var%>" class="td_admin_5">
																		<%if isFixedName then%>
																		<%=Parameter_Name%>
																		<%else%>
																		<input type="text" class="texts" name="Parameter_Name_Variabled_<%=Parameter_Id%>" ID="Parameter_Name_Variabled_<%=Parameter_Id%>" style="width:150px" value="<%=vFix(Parameter_Name_Variabled)%>" maxlength=150 dir="<%=dir_obj_var%>">
																		<%end if%>
																	</td>
																	<td align="<%=align_var%>"  dir="<%=dir_obj_var%>" class="td_admin_5"  style="background-color:#ccff99">
																		<input type="text" class="texts" name="Parameter_Value_Pegasus_<%=Parameter_Id%>" ID="Parameter_Value_Pegasus_<%=Parameter_Id%>"  style="width:200px" value="<%=vFix(Parameter_Value_Tour)%>" maxlength=150 dir="<%=dir_obj_var%>">
																	</td>
																	<td align="<%=align_var%>"  dir="<%=dir_obj_var%>" class="td_admin_5" style="background-color:#ffffcc">
																		<input type="text" class="texts" name="Parameter_Value_Screen_<%=Parameter_Id%>" ID="Parameter_Value_Screen_<%=Parameter_Id%>" style="width:200px" value="<%=vFix(Parameter_Value_Screen)%>" maxlength=150 dir="<%=dir_obj_var%>">
																	</td>
																	<td align="<%=align_var%>"  dir="<%=dir_obj_var%>" class="td_admin_5">
																		<table border="0" cellpadding="5" cellspacing="0" ID="Table7" width="100%" height="100%">
																			<tr>
																				<td width="50%" ID="tdAdvantages_1_<%=Parameter_Id%>" style="background-color:<%if Parameter_Advantages<>"" and isAdvantages then%>#ccffcc<%else%><%end if%>" nowrap>
																					<input type="checkbox" name="Parameter_Advantages_<%=Parameter_Id%>" ID="Parameter_Advantages_1_<%=Parameter_Id%>" value="1" <%if Parameter_Advantages<>"" and isAdvantages then%>checked<%end if%>>כן
																				</td>
																				<td width="50%"  ID="tdAdvantages_0_<%=Parameter_Id%>" style="background-color:<%if Parameter_Advantages<>"" and not isAdvantages then%>#ff6666<%else%><%end if%>" nowrap>
																					<input type="checkbox" name="Parameter_Advantages_<%=Parameter_Id%>" ID="Parameter_Advantages_0_<%=Parameter_Id%>" value="0" <%if Parameter_Advantages<>"" and not isAdvantages then%>checked<%end if%>>לא
																				</td>
																			</tr>
																		</table>
																	</td>
																	<td align="<%=align_var%>"  dir="<%=dir_obj_var%>" class="td_admin_5">
																		<textarea  class="texts" name="Advantages_Description_<%=Parameter_Id%>" ID="Advantages_Description_<%=Parameter_Id%>" rows=3  style="width:500px"><%=vFix(Advantages_Description)%></textarea>
																	</td>
																</tr>
																<%rssub.movenext
												loop
												set rssub=Nothing
												
												%>
															</table>
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr>
											<td bgcolor="#e6e6e6" height="10"></td>
										</tr>
										<!--site Competitor-->
										<tr>
											<td height="5" colspan="2" nowrap></td>
										</tr>
										<tr>
											<td height="5" nowrap colspan="2"></td>
										</tr>
										<tr>
											<td height="1" nowrap colspan="2" bgcolor="#808080"></td>
										</tr>
										<tr>
											<td height="10" nowrap colspan="2"></td>
										</tr>
										<tr>
											<td align="center" colspan="2">
												<table cellpadding="0" cellspacing="0" width="500" align="center" ID="Table9">
													<tr>
														<td width="50%" align="center">
															<input type=button class="but_menu" style="width:90px" onclick="document.location.href='screens.asp';" value="<%=arrButtons(2)%>" id=Button2 name=Button2></td>
														<td width="50" nowrap></td>
														<td width="50%" align="center"><input type=button class="but_menu" style="width:90px" onclick="javascript:return checkForm();" value="<%=arrButtons(1)%>" id=Button1 name=Button1></td>
													</tr>
												</table>
											</td>
										</tr>
									</table>
								</form>
							</td>
							<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
								<table cellpadding="1" cellspacing="1" width="100%" ID="Table10">
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
										<td nowrap align="center"><a class="button_edit_1" style="width:100;" href="screensAdmin.asp<%=squery%>"><span id="word6" name="word6"><!--הוספת מתחרה-->חזרה לרשימת מסכים</span></a></td>
									</tr>
								</table>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
					
		<SCRIPT type="text/javascript">
            <!--
            var cal1xx = new CalendarPopup('CalendarDiv');
                cal1xx.setYearSelectStartOffset(120);
	cal1xx.setYearSelectEndOffset(2);
	//cal1xx.setYearSelectStart(1910); //Added by Mila
                cal1xx.showNavigationDropdowns();
                cal1xx.offsetX = -50;
                cal1xx.offsetY = -40;
          
                //-->
		</SCRIPT>
		<DIV ID='CalendarDiv' STYLE='POSITION:absolute;VISIBILITY:hidden;BACKGROUND-COLOR:white;layer-background-color:white'></DIV>
	</body>
</html>

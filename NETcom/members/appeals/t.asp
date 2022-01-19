
<!--#include file="../../connect.asp"-->
 
<!--#include file="../../reverse.asp"-->


<!doctype html>
<html lang="en">
<head>
 <!--#include file="../checkWorker.asp"-->
 <%
	appId = trim(Request("appid"))
	quest_id = trim(Request("quest_id"))
	companyID = trim(Request("companyID"))
	contactID = trim(Request("contactID"))
	projectID = trim(Request("projectID")) 
	mechanismID = trim(Request("mechanismID"))
	if trim(request.QueryString ("Is_GroupsTour"))="1" then
	Is_GroupsTour=trim(request.QueryString ("Is_GroupsTour"))
	else
	Is_GroupsTour=0
	end if
	 lang_id =1	
	'טופס רשום לטיול
	ReservationId = trim(Request("ReservationId"))
	Country_CRM=0
	'response.Write ("ReservationId="& ReservationId)
	'response.end
'response.Write "Country_CRM="&Country_CRM
if trim(ReservationId)<>"" then
sqlstr="SELECT T.Tour_Id,Country_CRM,Insert_Date, TR.Reservation_Id	FROM " & pegasusDBName & ".dbo.Tour_Reservations TR LEFT OUTER JOIN " & pegasusDBName & ".dbo.Tours T ON TR.Tour_Id = T.Tour_Id	WHERE (Contact_Id = "& contactID &") and  Reservation_Id=" & ReservationId
set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			Country_CRM = trim(rs_pr("Country_CRM"))
		else
		Country_CRM=0
		End If
		set rs_pr = Nothing
end if
if not IsNumeric(Country_CRM) then Country_CRM=0
	title = false
	If trim(companyId) <> "" Then
		sqlstr = "Select company_Name from companies WHERE company_Id = " & companyId
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			companyName = trim(rs_pr("company_Name"))
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(contactID) <> "" Then
		sqlstr = "Select CONTACT_NAME from contacts WHERE contact_Id = " & contactID
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			contactName = trim(rs_pr(0))
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(projectID) <> "" Then
		sqlstr = "Select project_Name from projects WHERE project_Id = " & projectID
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			projectName = trim(rs_pr("project_Name"))
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	from = trim(Request("from"))
	If trim(from) = "" Then
		from = "pop_up"
	End If 	
	PathCalImage = "../../"
	
	arr_Status = session("arr_Status")
	
	sqlstr =  "Select Langu, PRODUCT_NAME, PRODUCT_DESCRIPTION, RESPONSIBLE, FILE_ATTACHMENT, ATTACHMENT_TITLE From Products WHERE PRODUCT_ID = " & quest_id
	'Response.Write sqlstr
	'Response.End
	set rsq = con.getRecordSet(sqlstr)
	If not rsq.eof Then		
		Langu = trim(rsq(0))
		productName = trim(rsq(1))
		PRODUCT_DESCRIPTION = trim(rsq(2))
		RESPONSIBLE = trim(rsq(3))
		attachment_file = trim(rsq(4))
		attachment_title = trim(rsq(5))		
	End If
	set rsq = Nothing	
	if Langu = "eng" then
		td_align = "left" : pr_language = "eng"
	else
		td_align = "right" : 	pr_language = "heb"
	end if
	
	userName = user_name 'cokie
	If IsNumeric(appid) Then
	sqlstr = "Exec dbo.get_appeals '','','','','" & OrgID & "','','','','','','','" & appid & "','',''"
'	Response.Write("sqlstr=" & sqlstr)
	set app = con.GetRecordSet(sqlstr)
	if not app.eof then	
	appeal_date = FormatDateTime(app("appeal_date"), 2) & " " & FormatDateTime(app("appeal_date"), 4)	
	quest_id = trim(app("questions_id"))
	companyID = app("company_ID")
	contactID = app("contact_ID")
	projectID = app("project_ID")
	mechanismID = app("mechanism_ID")
	companyName = app("company_Name")
	contactName = app("contact_Name")
	projectName = app("project_Name")
	userName = app("user_Name")
	appeal_status = app("appeal_status")
	private_flag = trim(app("private_flag"))
	UserIdOrderOwner = trim(app("User_Id_Order_Owner"))
	AppCountryID = trim(app("appeal_CountryId"))
	If trim(appeal_status) = "1" And trim(UserID) = trim(RESPONSIBLE) Then
		sqlstring="UPDATE appeals set appeal_status = '2' WHERE ORGANIZATION_ID = " & OrgID & " AND appeal_id = " & appid 	
		'Response.Write(sqlstring)
		'Response.End 
		con.ExecuteQuery(sqlstring) 
		appeal_status = "2"
		xmlFilePath = "../../../download/xml_forms/bizpower_forms.xml"
		'------ start deleting the new message from XML file ------
		set objDOM = Server.CreateObject("Microsoft.XMLDOM")
		objDom.async = false			
		if objDOM.load(server.MapPath(xmlFilePath)) then
				set objNodes = objDOM.getElementsByTagName("FORM")
				for j=0 to objNodes.length-1
					set objTask = objNodes.item(j)
					node_app_id = objTask.attributes.getNamedItem("APPEAL_ID").text										
					if trim(appId) = trim(node_app_id) Then					
						objDOM.documentElement.removeChild(objTask)
						exit for
					else
						set objTask = nothing
					end if
				next
				Set objNodes = nothing
				set objTask = nothing
				objDom.save server.MapPath(xmlFilePath)
			end if
			set objDOM = nothing
		' ------ end  deleting the new message from XML file ------
	End If	
	set app = nothing
	End If
	Else
		appeal_date = FormatDateTime(Now(), 2) & " " & FormatDateTime(Now(), 4)
	End If	
	
	If trim(appeal_status) = "1" And trim(UserID) = trim(RESPONSIBLE) Then
		sqlstring="UPDATE appeals set appeal_status = '2' WHERE ORGANIZATION_ID = " & OrgID & " AND appeal_id = " & appid 	
		'Response.Write(sqlstring)
		'Response.End 
		con.ExecuteQuery(sqlstring) 
		appeal_status = "2"
    End If			
	
	title = false
	If trim(companyId) <> "" Then
		sqlstr = "Select company_Name from companies WHERE company_Id = " & companyId
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			companyName = trim(rs_pr("company_Name"))
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(contactID) <> "" Then
		sqlstr = "Select CONTACT_NAME from contacts WHERE contact_Id = " & contactID
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			contactName = trim(rs_pr(0))
			title = true		
		End If
		set rs_pr = Nothing
	End If	
	
	If trim(projectID) <> "" Then
		sqlstr = "Select project_Name from projects WHERE project_Id = " & projectID
		set rs_pr = con.getRecordSet(sqlstr)
		If not rs_pr.eof Then
			projectName = trim(rs_pr("project_Name"))
			title = true		
		End If
		set rs_pr = Nothing
	End If	
%>
<%   
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 19 Order By word_id"				
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
	  set rsbuttons=nothing %>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  	<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<%If trim(Langu) = "eng" Then%>
<script language="javascript" type="text/javascript" src="../../../Scripts/CalendarPopup.js"></script>
<%Else%>
<script language="javascript" type="text/javascript" src="../../../Scripts/CalendarPopupH.js"></script>
<%End If%>
<script language="javascript" type="text/javascript">
	var cal = new CalendarPopup();
</script>
		<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
  <title>jQuery UI Autocomplete - Combobox</title>
  <link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
  <link rel="stylesheet" href="/resources/demos/style.css">
  <style>
  .custom-combobox {
    position: relative;
    display: inline-block;
  }
  .custom-combobox-toggle {
    position: absolute;
    top: 0;
    bottom: 0;
    margin-left: -1px;
    padding: 0;
  }
  .custom-combobox-input {
    margin: 0;
    padding: 5px 10px;
  }
  </style>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
  <script>
  $( function() {
    $.widget( "custom.combobox", {
      _create: function() {
        this.wrapper = $( "<span>" )
          .addClass( "custom-combobox" )
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
          .addClass( "custom-combobox-input ui-widget ui-widget-content ui-state-default ui-corner-left" )
          .autocomplete({
            delay: 0,
            minLength: 0,
            source: $.proxy( this, "_source" )
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
          },
 
          autocompletechange: "_removeIfInvalid"
        });
      },
 
      _createShowAllButton: function() {
        var input = this.input,
          wasOpen = false;
 
        $( "<a>" )
          .attr( "tabIndex", -1 )
          .attr( "title", "Show All Items" )
          .tooltip()
          .appendTo( this.wrapper )
          .button({
            icons: {
              primary: "ui-icon-triangle-1-s"
            },
            text: false
          })
          .removeClass( "ui-corner-all" )
          .addClass( "custom-combobox-toggle ui-corner-right" )
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
 
    $( "#combobox" ).combobox();
    $( "#toggle" ).on( "click", function() {
      $( "#combobox" ).toggle();
    });
  } );
  </script>
 
</head>
<body>
 
<div class="ui-widget">
  <label>Your preferred programming language: </label>
  <select id="combobox" NAME="combobox">
    <option value="">Select one...</option>
    <option value="ActionScript">ActionScript</option>
    <option value="AppleScript">AppleScript</option>
    <option value="Asp">Asp</option>
    <option value="BASIC">BASIC</option>
    <option value="C">C</option>
    <option value="C++">C++</option>
    <option value="Clojure">Clojure</option>
    <option value="COBOL">COBOL</option>
    <option value="ColdFusion">ColdFusion</option>
    <option value="Erlang">Erlang</option>
    <option value="Fortran">Fortran</option>
    <option value="Groovy">Groovy</option>
    <option value="Haskell">Haskell</option>
    <option value="Java">Java</option>
    <option value="JavaScript">JavaScript</option>
    <option value="Lisp">Lisp</option>
    <option value="Perl">Perl</option>
    <option value="PHP">PHP</option>
    <option value="Python">Python</option>
    <option value="Ruby">Ruby</option>
    <option value="Scala">Scala</option>
    <option value="Scheme">Scheme</option>
  </select>
</div>
 
 
</body>
</html>
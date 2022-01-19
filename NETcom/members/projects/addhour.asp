<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
<!--
   	
	function GetNumbers ()
	{
		var ch=event.keyCode;
		event.returnValue =(ch >= 48 && ch <= 57) || ch == 58 ;
	} 	
	
	function checkFields()
	{
		if(window.document.all("company_id").value == "")
		{
			<%
				If trim(lang_id) = "1" Then
					str_alert = "אנא בחר " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & " !!"
				Else
					str_alert = "Please choose " & trim(Request.Cookies("bizpegasus")("CompaniesOne")) & " !!"
				End If   
			%>			
			window.alert("<%=str_alert%>");				
			window.document.all("company_id").focus();
			return false;
		}
		if(window.document.all("project_id").value == "")
		{
			<%
				If trim(lang_id) = "1" Then
					str_alert = "!! אנא בחר " & trim(Request.Cookies("bizpegasus")("Projectone"))
				Else
					str_alert = "Please choose " & trim(Request.Cookies("bizpegasus")("Projectone")) & "!!"
				End If   
			%>			
			window.alert("<%=str_alert%>");			
			window.document.all("project_id").focus();
			return false;
		}
		if(window.document.all("mechanism_id").value == "")
		{
			<%
				If trim(lang_id) = "1" Then
					str_alert = "!! אנא בחר מנגנון"
				Else
					str_alert = "Please choose mechanism !!"
				End If   
			%>			
			window.alert("<%=str_alert%>");			
			window.document.all("mechanism_id").focus();
			return false;
		}
		if(window.document.all("hours").value == "")
		{
			<%
				If trim(lang_id) = "1" Then
					str_alert = "!! אנא מלא את שעות העבודה על ה " & trim(Request.Cookies("bizpegasus")("Projectone"))
				Else
					str_alert = "Please fill in the labor hours on " & trim(Request.Cookies("bizpegasus")("Projectone")) & "!!"
				End If   
			%>			
			window.alert("<%=str_alert%>");					
			window.document.all("hours").focus();
			return false;
		}
	}
	
function filterlist(selectobj) {

  //==================================================
  // PARAMETERS
  //==================================================

  // HTML SELECT object
  // For example, set this to document.myform.myselect
  this.selectobj = selectobj;

  // Flags for regexp matching.
  // "i" = ignore case; "" = do not ignore case
  // You can use the set_ignore_case() method to set this
  this.flags = 'i';

  // Which parts of the select list do you want to match?
  this.match_text = true;
  this.match_value = false;

  // You can set the hook variable to a function that
  // is called whenever the select list is filtered.
  // For example:
  // myfilterlist.hook = function() { }

  // Flag for debug alerts
  // Set to true if you are having problems.
  this.show_debug = false;

  //==================================================
  // METHODS
  //==================================================

  //--------------------------------------------------
  this.init = function() {
    // This method initilizes the object.
    // This method is called automatically when you create the object.
    // You should call this again if you alter the selectobj parameter.

    if (!this.selectobj) return this.debug('selectobj not defined');
    if (!this.selectobj.options) return this.debug('selectobj.options not defined');

    // Make a copy of the select list options array
    this.optionscopy = new Array();
    if (this.selectobj && this.selectobj.options) {
      for (var i=0; i < this.selectobj.options.length; i++) {

        // Create a new Option
        this.optionscopy[i] = new Option();

        // Set the text for the Option
        this.optionscopy[i].text = selectobj.options[i].text;

        // Set the value for the Option.
        // If the value wasn't set in the original select list,
        // then use the text.
        if (selectobj.options[i].value) {
          this.optionscopy[i].value = selectobj.options[i].value;
        } else {
          this.optionscopy[i].value = selectobj.options[i].text;
        }

      }
    }
  }

  //--------------------------------------------------
  this.reset = function() {
    // This method resets the select list to the original state.
    // It also unselects all of the options.

    this.set('');
  }


  //--------------------------------------------------
  this.set = function(pattern) {
    // This method removes all of the options from the select list,
    // then adds only the options that match the pattern regexp.
    // It also unselects all of the options.

    var loop=1, index=1, regexp, e;

    if (!this.selectobj) return this.debug('selectobj not defined');
    if (!this.selectobj.options) return this.debug('selectobj.options not defined');

    // Clear the select list so nothing is displayed
    this.selectobj.options.length = 1;    
 
    // Set up the regular expression.
    // If there is an error in the regexp,
    // then return without selecting any items.
    try {

      // Initialize the regexp
      regexp = new RegExp(pattern, this.flags);

    } catch(e) {

      // There was an error creating the regexp.

      // If the user specified a function hook,
      // call it now, then return
      if (typeof this.hook == 'function') {
        this.hook();
      } 
      
      return;
    }
    
	    
    // Loop through the entire select list and
    // add the matching items to the select list
    for (loop=1; loop < this.optionscopy.length; loop++) {

      // This is the option that we're currently testing
      var option = this.optionscopy[loop];

      // Check if we have a match
      if ((this.match_text && regexp.test(option.text)) ||
          (this.match_value && regexp.test(option.value))) {

        // We have a match, so add this option to the select list
        // and increment the index

        this.selectobj.options[index++] =
          new Option(option.text, option.value, false);

      }
    }    
     
    this.selectobj.selectedIndex = 0;
     
    // If the user specified a function hook,
    // call it now
    if (typeof this.hook == 'function') {
      this.hook();
    }      

  }

  //--------------------------------------------------
  this.set_ignore_case = function(value) {
    // This method sets the regexp flags.
    // If value is true, sets the flags to "i".
    // If value is false, sets the flags to "".

    if (value) {
      this.flags = 'i';
    } else {
      this.flags = '';
    }
  }


  //--------------------------------------------------
  this.debug = function(msg) {
    if (this.show_debug) {
      alert('FilterList: ' + msg);
    }
  }


  //==================================================
  // Initialize the object
  //==================================================
  this.init();

}	
//-->
</script> 
</head> 
<%
	UserID=trim(Request.Cookies("bizpegasus")("UserID"))
	OrgID=trim(trim(Request.Cookies("bizpegasus")("ORGID")))
	date_ = trim(Request("date"))
	companyId = trim(Request("companyId"))
	projectID = trim(Request("projectID"))
	end_date = Request("end_date")
	start_date = Request("start_date")
	If trim(Request("USER_ID")) <> "" Then
		UserID = trim(Request("USER_ID"))
	End If
	If trim(companyID)<>"" then   
		sqlStr = "SELECT company_name FROM companies WHERE company_id="& companyID 
		set pr=con.GetRecordSet(sqlStr)
		if not pr.EOF then	
			company_name = pr("company_name")
		end if
		set pr = Nothing
	End If	
	if Request("add")<>nil then		
		company_id = trim(Request.Form("company_id"))   
		project_id = trim(Request.Form("project_id"))
		mechanism_id = trim(Request.Form("mechanism_id"))  
		hours = trim(Request.Form("hours"))   
		If trim(hours) <> "" Then
			arrhours = split(hours,":")
			if arrhours(0) <> "" then
				minuts = cdbl(arrhours(0)) * 60
			else
				minuts = 0
			end if	
			if ubound(arrhours) > 0 then
				if arrhours(1) <> "" then
					minuts = minuts + cdbl(arrhours(1))
				end if	
			end if	
			
			sqlstr = "SET DATEFORMAT DMY; Insert Into hours (user_id,organization_id,date,project_id,company_id,mechanism_id,minuts) VALUES (" &_
			userID & "," & orgID & ",'" & date_ & "'," & project_id & "," & company_id & "," & mechanism_id & "," & minuts & ")"
			'Response.Write sqlstr
			'Response.End
			con.executeQuery(sqlstr)
	    
	%>
	<script language=javascript>
	<!--
		window.opener.document.location.href = window.opener.document.location.pathname + "?dtCurrentDate=<%=date_%>&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=UserID%>";
		window.close();	
	//-->
    </script>
  
	<%
	End If
end if
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 53 Order By word_id"				
	  set rstitle = con.getRecordSet(sqlstr)		
	  If not rstitle.eof Then
	    dim arrTitles()
		arrTitlesD = ";" & trim(rstitle.getString(,,",",";"))
		arrTitlesD = Split(trim(arrTitlesD), ";")		
		redim arrTitles(Ubound(arrTitlesD))
		For i=1 To Ubound(arrTitlesD)			
			arr_title = Split(arrTitlesD(i),",")			
			If IsArray(arr_title) And Ubound(arr_title) = 1 Then
			arrTitles(trim(arr_title(0))) = arr_title(1)
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
<body style="margin:0px;background:#E6E6E6" onload="window.focus();">
<table border="0" cellpadding="0" cellspacing="1" width="100%" dir="<%=dir_var%>">
<tr>
   <td bgcolor="#e6e6e6" align="left" valign="middle" nowrap>
	 <table width="100%" border="0" cellpadding="0" cellspacing="0" dir="<%=dir_var%>">
	  <tr><td class="page_title"><span id=word1 name=word1><!--הוספת שעות עבודה על--><%=arrTitles(1)%></span> <%=trim(Request.Cookies("bizpegasus")("Projectone"))	%>&nbsp;</td></tr>		   
   </table></td></tr>         
   <tr>    
    <td width="100%" valign="top" align="center">    
	<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1">	
	<form name="add_hour" id="add_hour" method=post action="addhour.asp?date=<%=date_%>&add=1&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=UserID%>">	
	<tr>
	<td align="<%=align_var%>" width="310">
	<INPUT NAME=regexp class="Form" onKeyUp="myfilter.set(this.value)" style="width:300" dir="rtl" ID="regexp" onChange="myfilter.set(this.form.regexp.value)"><br>
	<select dir="<%=dir_obj_var%>" name="company_id" style="width:300px;font-family:arial" onchange="document.location.href='addhour.asp?date=<%=date_%>&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=UserID%>&companyId='+this.value;">
	<%If trim(lang_id) = "1" Then%>
		<option value=""><%=String(20,"-")%> בחר <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%> <%=String(20,"-")%></option>
	<%Else%>
	    <option value=""><%=String(20,"-")%> Choose <%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%> <%=String(20,"-")%></option>
	<%End If%>	
		<%  sqlstr = "Select COMPANY_ID, COMPANY_NAME FROM COMPANIES Where ORGANIZATION_ID = "& OrgID & " Order BY PRIVATE_FLAG DESC,COMPANY_NAME"
			set rs_comp = con.getRecordSet(sqlstr)
			While not rs_comp.eof
		%>
		<option value="<%=rs_comp(0)%>" <%If trim(companyId) = trim(rs_comp(0)) Then%> selected <%End If%>><%=rs_comp(1)%></option>
		<%
			rs_comp.moveNext
			Wend
			set rs_comp = Nothing
		%>
	</select>
	<SCRIPT TYPE="text/javascript">
	<!--
	   var myfilter = new filterlist(document.add_hour.company_id);
	//-->
	</SCRIPT>
	</td>
	<td align="<%=align_var%>" valign=top><b><%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%></b>&nbsp;</td>
	</tr>
	<%If trim(companyId) <> "" Then%>
	<tr>		
	<td align="<%=align_var%>" width="310">
	<select dir="<%=dir_obj_var%>" name="project_id" style="width:300px;font-family:Arial" ID="project_id" onchange="document.location.href='addhour.asp?date=<%=date_%>&start_date=<%=start_date%>&end_date=<%=end_date%>&User_ID=<%=UserID%>&companyId=<%=companyID%>&projectID='+this.value;" >
	<%If trim(lang_id) = "1" Then%>
	<option value=""><%=String(20,"-")%> בחר <%=trim(Request.Cookies("bizpegasus")("Projectone"))%> <%=String(20,"-")%></option>
	<%Else%>
	<option value=""><%=String(20,"-")%> Choose <%=trim(Request.Cookies("bizpegasus")("Projectone"))%> <%=String(20,"-")%></option>
	<%End If%>
	<%  		
		sqlstr = "SET DATEFORMAT DMY; Select PROJECT_ID, PROJECT_NAME FROM PROJECTS WHERE company_id IN (0," & companyID &_
	    ") AND ORGANIZATION_ID = " & OrgID & " AND active = 1 AND rtrim(ltrim(status)) <> '3' Order BY PROJECT_NAME"
	    'Response.Write sqlstr
	    'Response.End
		set rs_proj = con.getRecordSet(sqlstr)
		Do While not rs_proj.eof
	%>
	<option value="<%=rs_proj(0)%>" <%If trim(projectID) = trim(rs_proj(0)) Then%> selected <%End If%>><%=rs_proj(1)%></option>
	<%
		rs_proj.moveNext
		Loop
		set rs_proj = Nothing
	%>
	</select></td>
	<td align="<%=align_var%>"><b><%=trim(Request.Cookies("bizpegasus")("Projectone"))%></b>&nbsp;</td>
	</tr>
	<%  End If
	    If trim(companyID) <> "" And trim(projectID) <> "" Then		
	%>
	<tr>		
	<td align="<%=align_var%>" width="310">
	<select dir="<%=dir_obj_var%>" name="mechanism_id" style="width:300px;font-family:Arial" ID="mechanism_id">
	<%If trim(lang_id) = "1" Then%>
	<option value=""><%=String(20,"-")%> בחר מנגנון <%=String(20,"-")%></option>
	<%Else%>
	<option value=""><%=String(18,"-")%> Choose mechanism <%=String(18,"-")%></option>
	<%End If    
		sqlstr = "Select mechanism_id, mechanism_name From mechanism Where company_id IN(0," & companyID &_
	    ") AND ORGANIZATION_ID = " & OrgID & " AND PROJECT_ID = " & projectID & " AND mechanism_id NOT IN ("&_
	    " Select DISTINCT mechanism_id FROM hours WHERE ORGANIZATION_ID = " & orgID &_
	    " And USER_ID = " & userID & " And COMPANY_ID = " & companyId & " AND date = '" & date_ & "'"&_
	    " And Project_ID = " & projectID & ")  Order BY mechanism_id"
	    'Response.Write sqlstr
	    'Response.End
		set rs_mech = con.getRecordSet(sqlstr)
		Do While not rs_mech.eof
	%>
	<option value="<%=rs_mech(0)%>"><%=rs_mech(1)%></option>
	<%
		rs_mech.moveNext
		Loop
		set rs_mech = Nothing
	%>
	</select></td>
	<td align="<%=align_var%>"><b><!--מנגנון--><%=arrTitles(3)%></b>&nbsp;</td>
	</tr>	
	<tr>
	<td align="<%=align_var%>"><input type=text class="texts" name="hours" id="hours" value="<%=hours_pr%>" onkeypress="GetNumbers();" maxlength=4 style="width:45"></td>		
	<td align="<%=align_var%>"><b><span id="word2" name=word2><!--שעות--><%=arrTitles(2)%></span></b>&nbsp;</td>
	</tr>
	<%End If%>	
	<tr><td colspan="2" height=20 nowrap></td></tr>
	<tr><td colspan="2">
	<table width=100% border=0 cellspacing=0>
	<tr><td width=48% align="center">
	<INPUT class="but_menu" style="width:90" type="button" value="<%=arrButtons(2)%>" id=button2 name=button2 onclick="window.close();"></td>
	<td width=4% nowrap></td>
	<td width=48% align="center">
	<INPUT class="but_menu" style="width:90" type="submit" value="<%=arrButtons(1)%>" id=Button1 name=Button1 onclick="return  checkFields()"></td>
	</tr>
	</table>
	</td></tr>
	</form>		
</table>
</td></tr></table>
</body>
</html>
<%
set con = nothing
%>
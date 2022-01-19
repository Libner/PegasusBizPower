<!--#INCLUDE FILE="../Netcom/reverse.asp"-->
<!--#include file="../Netcom/connect.asp"-->
<%quest_id = Decode(trim(Request.QueryString(Encode("P"))))	
	 appID = Decode(trim(Request.QueryString(Encode("appID")))) 
	
	 HTTP_REFERER = Request.ServerVariables("HTTP_REFERER")
	 MYreferer    = Request.ServerVariables("HTTP_HOST")  
    
	 arrPath=Split(HTTP_REFERER,"/")
	 If UBound(arrPath) > 1 Then
		referer=arrPath(2)
	 End If
	
	 pageId = "" : fromID = ""
	
	If InStr(HTTP_REFERER, "?") Then
        arrTemp = Split(HTTP_REFERER, "?")
        strTemp = arrTemp(UBound(arrTemp))       
		arrQuery = Split(strTemp & "&", "&")	
		
        For i=0 To Ubound(arrQuery) - 1        
			arrTmp1 = Split(trim(arrQuery(i)), "=")
			If trim(arrTmp1(0)) = "pageId" Then
				pageId = trim(arrTmp1(1))				
			End If
			If trim(arrTmp1(0)) = "fromID" Then
				fromID = trim(arrTmp1(1))				
			End If			
        Next        
	End If
	
	'response.Write "fromID=" & fromID & " pageId=" & pageId
	'response.End
	
	PathCalImage = "../netcom/"
	show_sale = "True"

	if trim(quest_id) = "" OR IsNumeric(quest_id) = false then
		show_sale = "False"
		reason = "SEKER NOT FOUND"		
	else
		sqlStr = "Select Product_Name, PRODUCT_TYPE, PRODUCT_DESCRIPTION, Langu, QUESTIONS_ID,USER_ID,"&_
		" DATE_START, DATE_END, ORGANIZATION_ID, ADD_CLIENT, FILE_ATTACHMENT, ATTACHMENT_TITLE " & _
		" FROM dbo.products WHERE (PRODUCT_ID=" & quest_id & ") AND (SHOW_INSITE = 1)"
		'Response.Write sqlStr
		'Response.End
		set rs_products = con.GetRecordSet(sqlStr)
		if not rs_products.eof then			
				PRODUCT_TYPE = rs_products("PRODUCT_TYPE")							
				OrgId = rs_products("ORGANIZATION_ID")
				PRODUCT_DESCRIPTION = rs_products("PRODUCT_DESCRIPTION")
				PRODUCT_NAME = trim(rs_products("Product_Name"))
				Langu = trim(rs_products("Langu"))
				addClient = trim(rs_products("ADD_CLIENT"))		
				attachment_file = trim(rs_products("FILE_ATTACHMENT"))
				attachment_title = trim(rs_products("ATTACHMENT_TITLE"))	
				if Langu = "eng" then
					dir_var = "ltr"
					dir_obj_var = "rtl"
					td_align = "left"
					pr_language = "eng"
					lang_id = "2"
				else
					dir_var = "rtl"
					dir_obj_var = "ltr"
					td_align = "right"
					pr_language = "heb"
					lang_id = "1"
				end if																			
				USER_ID = trim(rs_products("USER_ID"))	
				sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 14 Order By word_id"				
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
			else
				show_sale = "False"
				reason = "SEKER NOT FOUND"		
			end if		
		set rs_products = nothing		
	End If
	date_year = Year(date())
	date_month = Month(date())
	date_day = Day(date())

	sizeLogoPr = 0
	If isNumeric(trim(quest_id)) Then
		sqlpg="SELECT PRODUCT_LOGO FROM PRODUCTS WHERE PRODUCT_ID="&quest_id&""
		'Response.Write sqlpg
		'Response.End
		set pg=con.getRecordSet(sqlpg)	
		If not pg.EOF Then 
		sizeLogoPr=pg.Fields("PRODUCT_LOGO").ActualSize
		Else
		sizeLogoPr = 0
		End If
		set pg = Nothing
End If

sizeLogoOrg = 0
If trim(OrgID) <> "" And sizeLogoPr = 0 Then
	sqlstring = "SELECT ORGANIZATION_LOGO FROM organizations WHERE ORGANIZATION_ID = " & OrgID
	set logo=con.GetRecordSet(sqlstring)
	if not logo.EOF  then				
		sizeLogoOrg = logo("ORGANIZATION_LOGO").ActualSize
	else
	    sizeLogoOrg = 0
	end if
	set logo = Nothing	
End If

UseCaptcha = 0	
If trim(OrgID) <> "" Then
	sqlstring = "SELECT UseCaptcha FROM organizations WHERE ORGANIZATION_ID = " & OrgID
	set rs_tmp = con.GetRecordSet(sqlstring)
	if not rs_tmp.EOF  then				
		UseCaptcha = trim(rs_tmp("UseCaptcha"))
	else
	    UseCaptcha = 0	
	end if
	set rs_tmp = Nothing	
End If

Function RandomPassword(PasswordLength)

    Dim lNumberOfLowerCases
    Dim lNumberOfUpperCases
    Dim lNumberOfNumbers
    Dim l, j
    
    ReturnedPassword = ""
    If PasswordLength < 3 Then        
        Exit Function
    End If

    'Get the number of digits for each type of characters
    Randomize
    lNumberOfLowerCases = CInt((PasswordLength - 3) * Rnd) + 1
    'lNumberOfUpperCases = CInt((PasswordLength - lNumberOfLowerCases - 2) * Rnd) + 1
    lNumberOfUpperCases = 0
    lNumberOfNumbers = PasswordLength - lNumberOfLowerCases - lNumberOfUpperCases

       
    ReturnedPassword = ""
    For l = 1 To PasswordLength
        Randomize
        j = CInt(2 * Rnd + 1)
        Select Case j
        Case 1 'Lower Case
            If lNumberOfLowerCases > 0 Then
                ReturnedPassword = ReturnedPassword & Chr(CInt(25 * Rnd) + 97)
                lNumberOfLowerCases = lNumberOfLowerCases - 1
            Else
                l = l - 1 'Re-do the loop
            End If
        Case 2 'Upper Case
            If lNumberOfUpperCases > 0 Then
                ReturnedPassword = ReturnedPassword & Chr(CInt(25 * Rnd) + 65)
                lNumberOfUpperCases = lNumberOfUpperCases - 1
            Else
                l = l - 1 'Re-do the loop
            End If
        Case 3 'Number
            If lNumberOfNumbers > 0 Then
                ReturnedPassword = ReturnedPassword & CInt(9 * Rnd)
                lNumberOfNumbers = lNumberOfNumbers - 1
            Else
                l = l - 1 'Re-do the loop
            End If
        End Select
        For j = 1 To 100
            'Give the seed some time
        Next
    Next
    
        
    RandomPassword = ReturnedPassword
End Function %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=PRODUCT_NAME%></title>
<meta http-equiv="Content-Type" content="text/html; charset=Windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../netcom/IE4.css" rel="STYLESHEET" type="text/css">
<script language="javascript" type="text/javascript">
<!--
function CheckFields()
{
		<%If trim(addClient) = "1" Or trim(addClient) = "2" Then%>
		if (window.document.getElementById("CONTACT_NAME").value=='')
		{
		   <%
				If trim(lang_id) = "1" Then
					str_alert = "נא למלא שם"
				Else
					str_alert = "Please, insert the full name"
				End If	
			%>
 			window.alert("<%=str_alert%>");
 			window.document.getElementById("CONTACT_NAME").focus();
			return false;
		}	
		
		if (window.document.getElementById("contact_cellular").value=='')
		{
			window.alert("נא למלא טלפון נייד");
			window.document.getElementById("contact_cellular").focus();
			return false;
		}
				
		if (window.document.getElementById("contact_email").value=='')
		{
			window.alert("נא למלא כתובת דואר אלקטרוני");
			window.document.getElementById("contact_email").focus();
			return false;
		}
		else if (!checkEmail(document.getElementById("contact_email").value))
		{
			<%
				If trim(lang_id) = "1" Then
					str_alert = "כתובת דואר אלקטרוני לא חוקית"
				Else
					str_alert = "The email address is not valid!"
				End If	
			%>
				window.alert("<%=str_alert%>");
				document.getElementById("contact_email").focus();
				return false;
		} 
		<%End If%>	
		<%If trim(addClient) = "2" Then%>
		if (window.document.getElementById("company_name").value=='')
		{
		   <%
				If trim(lang_id) = "1" Then
					str_alert = "נא למלא שם חברה"
				Else
					str_alert = "Please, insert the company name"
				End If	
			%>
 			window.alert("<%=str_alert%>");
 			window.document.getElementById("company_name").focus();
			return false;
		}	
		<%End If%>
		<%If trim(quest_id) = "8365" Then%>
		if(! checkAnswers(5))
			return false;
		<%End If%>
		<%If trim(quest_id) = "8366" Then%>
		if(! checkAnswers(4))
			return false;
		<%End If%>		
		if(click_fun() == false)
		{
			return false;
		}
		<%If trim(UseCaptcha) = "1" Then%>
		if(window.document.getElementById("txt_captcha").value == "")
		{
			window.alert('עליך להקליד את התווים שאתה רואה בתמונה  למטה');
			return false;
		}
		<%End If%>
		window.document.form1.submit();
		return false;
}		

function checkAnswers(numOfAnswers)
{
   arr = document.form1.getElementsByTagName("INPUT");
   //window.alert(arr.length);  
   count = 0;
   for(i=0;i<arr.length;i++)
   {
     id = arr[i].id;
     if(arr[i].type == "checkbox")
     {
		arr_radio = document.getElementsByName(id);		
		for(j=0;j<arr_radio.length;j++)
		{		
			if(arr_radio[j].checked == true)
			{
				count = count + 1;
			}	
		}		
	 }		 
   }
   if((count > numOfAnswers) || (count == 0))
   {		 	
		window.alert(".יש לבחור עד " + numOfAnswers + " מועמדים");	
		return false;
   }
     		
   return true;
}	
		
function checkEmail(addr)
{
	addr = addr.replace(/^\s*|\s*$/g,"");
	if (addr == '') {
	return false;
	}
	var invalidChars = '\/\'\\ ";:?!()[]\{\}^|';
	for (i=0; i<invalidChars.length; i++) {
	if (addr.indexOf(invalidChars.charAt(i),0) > -1) {
		return false;
	}
	}
	for (i=0; i<addr.length; i++) {
	if (addr.charCodeAt(i)>127) {     
		return false;
	}
	}

	var atPos = addr.indexOf('@',0);
	if (atPos == -1) {
	return false;
	}
	if (atPos == 0) {
	return false;
	}
	if (addr.indexOf('@', atPos + 1) > - 1) {
	return false;
	}
	if (addr.indexOf('.', atPos) == -1) {
	return false;
	}
	if (addr.indexOf('@.',0) != -1) {
	return false;
	}
	if (addr.indexOf('.@',0) != -1){
	return false;
	}
	if (addr.indexOf('..',0) != -1) {
	return false;
	}
	var suffix = addr.substring(addr.lastIndexOf('.')+1);
	if (suffix.length != 2 && suffix != 'com' && suffix != 'net' && suffix != 'org' && suffix != 'edu' && suffix != 'int' && suffix != 'mil' && suffix != 'gov' & suffix != 'arpa' && suffix != 'biz' && suffix != 'aero' && suffix != 'name' && suffix != 'coop' && suffix != 'info' && suffix != 'pro' && suffix != 'museum') {
	return false;
	}
return true;
}
	
function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("<%=strLocal%>Netcom/calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=157pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}

function KeyPress(e) 
{
	if( typeof( e ) == "undefined" && typeof( window.event ) != "undefined" )
		e = window.event;
	
	if (e.keyCode == 13)
	{		
		if(e.srcElement.tagName != 'TEXTAREA')
		{
			e.keyCode = 9;		
		}	
	}	
	return true;
}

function ClearFields()
{
	var answer = window.confirm('?האם ברצונך לנקות את כל שדות הטופס');
	if(answer == true)
	{
		var arr_input = window.document.form1.getElementsByTagName('INPUT');
		for(i=0;i<arr_input.length;i++)
		{
			var elem = arr_input[i];
			///Clear all input text fields on the form
			if(elem.type == "text")
			{
				elem.value = "";
			} 
			///Clear all checkboxes and radiobuttons on the form
			if(elem.type == "checkbox" || elem.type == "radio")
			{
				elem.checked = false;
			} 
		}
		///Clear all textareas on the form
		var arr_textarea = window.document.form1.getElementsByTagName('textarea');
		for(i=0;i<arr_textarea.length;i++)
		{
			arr_textarea[i].value = "";
		}
		///Clear all comboboxes on the form
		var arr_select = window.document.form1.getElementsByTagName('select');
		for(i=0;i<arr_select.length;i++)
		{
			arr_select[i].value = "";
		}
	}
	return false;
}
//-->
</script>
</head>
<body onKeyPress="return KeyPress(event);">
<table border="0" width="650" align=center cellspacing="0" cellpadding="0">
<%If sizeLogoPr > 0 OR sizeLogoOrg > 0 Then%>
<tr>
      <td width="100%" valign="bottom" align=center>
      <table border="0" width="90%" cellspacing="1" cellpadding="1" align=center>
       <tr><td height="5"></td></tr>
        <tr>
         <td valign="bottom" align="center" width="100%">     
         <%If sizeLogoPr > 0 Then%>
         <img src="<%=strLocal%>/netcom/GetImage.asp?DB=PRODUCT&amp;FIELD=PRODUCT_LOGO&amp;ID=<%=quest_id%>">
         <%ElseIf sizeLogoOrg > 0 Then%>     
          <img src="<%=strLocal%>/netcom/GetImage.asp?DB=ORGANIZATION&amp;FIELD=ORGANIZATION_LOGO&amp;ID=<%=OrgId%>">          
         <%End If%> 
          </td>                
        </tr>
       <tr><td height="5"></td></tr> 
      </table>
      </td>
    </tr>  
<%End If%>      
<%If trim(show_sale) = "True" then%>   
<!-- start SEKER -->
	<tr><td align="center">
		<table WIDTH="90%" ALIGN="center" BORDER="0" CELLPADDING="0" cellspacing="0">			
		<tr>
		<td align="<%=td_align%>" width="100%" valign="top"> 
		<form action="addappeal.asp" id="form1" name="form1" method="post">			
		<input type="hidden" id="pageId" name="pageId" value="<%=pageId%>">
		<input type="hidden" id="fromID" name="fromID" value="<%=fromID%>">
		<input type="hidden" id="referer" name="referer" value="<%=Server.HTMLEncode(referer)%>">		
		<input type="hidden" name="OrgID" value="<%=OrgID%>" ID="OrgID">
		<input type="hidden" name="in_date" value="<%=date_day%>/<%=date_month%>/<%=date_year%>" ID="in_date">
		<input type="hidden" name="Langu" value="<%=Langu%>" ID="Hidden2">		
		<table border="0" cellpadding="5" cellspacing="1" width="100%" align="<%=td_align%>" bgcolor="#FFFFFF">					  
			<tr>
				<td class="Field_Title" style="font-size:16pt" align=center width=100% dir="<%=dir_var%>">
				<INPUT type="hidden" id=quest_id name=quest_id value="<%=quest_id%>">
				<b><%=PRODUCT_NAME%></b>
				</td>
			</tr>							
			<%if PRODUCT_DESCRIPTION <> "" then%>
			<TR>
				<td class="form_makdim" dir="<%=dir_var%>" width=100% align="<%=td_align%>">
				<%=breaks(PRODUCT_DESCRIPTION)%>
				</td>
			</tr>
			<%end if%>
			<%If Len(attachment_file) > 0 Then%>
			<TR>
				<td class="form_makdim" dir="rtl" width=100% align=<%=td_align%>>
				<a class="file_link" href="<%=strLocal%>/download/products/<%=attachment_file%>" target=_blank><%=attachment_title%></a>
				</td>
			</tr>	
			<%End If%>
			<tr><td height=10 nowrap></td></tr>			
			<!-- הוספת לקוח פרטי -->			
			<%If trim(addClient) = "1" Then%>
			<tr>       		 		                  
			<td width=100% dir="<%=dir_obj_var%>"> 
			<table border="0" bgcolor=#FFFFFF style="border: 1px solid #999999" cellpadding="2" cellspacing="0" width="100%" dir="<%=dir_obj_var%>" ID="Table1">   
				<tr><td width="100%" colspan=2 class="form_header" style="padding-right:15px;padding-left:15px;" align="<%=td_align%>">&nbsp;
				<b><!--פרטים אישיים--><%=arrTitles(21)%></b>&nbsp;
				</td></tr>	
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>
				<tr>
					<td width="100%" class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">           
					<input type=text dir="<%=dir_var%>" name="contact_name" value="<%=vFix(contact_name)%>" maxlength=50 style="width:200" class="Field_Title" ID="CONTACT_NAME">           
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap class="Field_Title" style="padding-right:15px;padding-left:15px;">					
					&nbsp;<b><!--שם מלא--><%=arrTitles(14)%></b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
				</tr> 
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%" class="Field_Title" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="contact_phone" value="<%=vFix(contact_phone)%>" maxlength="20" style="width:100" class="Field_Title" ID="contact_phone">                
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">
					&nbsp;<b><!--טלפון--><%=arrTitles(16)%></b>&nbsp;</td>
				</tr>
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   				
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="contact_cellular" value="<%=vFix(contact_cellular)%>" style="width:100" maxlength=20 class="Field_Title" ID="contact_cellular">              
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">
					&nbsp;<b><!--טלפון נייד--><%=arrTitles(17)%></b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
				</tr>			
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   				
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">                    
						<input type=text name="contact_email" dir=ltr value="<%=vFix(contact_email)%>" maxlength=50 style="width:200" class="Field_Title" ID="contact_email">              
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">
					&nbsp;<b>Email</b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
				</tr>
			</table>
			</td>
			</tr>	
			<!--הוספת לקוח ואיש קשר -->
			<%ElseIf trim(addClient) = "2" Then ' הוספת לקוח ארגוני%>		
			<tr>       		 		                  
			<td width=100% dir="<%=dir_obj_var%>"> 
			<table border="0" bgcolor=#FFFFFF style="border: 1px solid #999999" cellpadding="2" cellspacing="0" width="100%" dir="<%=dir_obj_var%>">   
				<tr><td width="100%" colspan=2 class="form_header" style="padding-right:15px;padding-left:15px;" align="<%=td_align%>">&nbsp;
				<b><!--פרטים איש קשר--><%=arrTitles(22)%></b>&nbsp;
				</td></tr>	
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>
				<tr>
					<td width="100%" class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">           
					<input type=text dir="<%=dir_var%>" name="contact_name" value="<%=vFix(contact_name)%>" maxlength=50 style="width:200" class="Field_Title">           
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">					
					&nbsp;<b><!--שם מלא--><%=arrTitles(14)%></b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
				</tr> 
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   
			    <tr>
			        <td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_obj_var%>" style="padding-right:15px;padding-left:15px;"><span dir="rtl"><b><%if FIELD_MUST then%><font color=red>&nbsp;*&nbsp;</font><%end if%><%=breaks(Field_Title)%></b></span>
					  <!--INPUT type=image src="../netcom/images/icon_find.gif" name=messangers_list id="messangers_list" onclick='window.open("../netcom/members/companies/messangers_list.asp?OrgId=<%=OrgId%>","MessangersList","left=300,top=100,width=250,height=350,scrollbars=1");return false;' title="<%=arrTitles(19)%>" tabindex=-1-->&nbsp;			        
             		  <input type=text dir="<%=dir_var%>" name="messangerName" value="<%=vFix(messangerName)%>" style="width:200" maxlength=50 class="Field_Title" ID="Text2">
             		</td>
				<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">					
				<b><!--עיסוק--><%=arrTitles(15)%></b></td>
				</tr> 
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
					<input type=text dir="<%=dir_var%>" name="contact_address" value="<%=vFix(contact_address)%>"  style="width:200" class="Field_Title" ID="contact_address"></td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;"><!--כתובת--><b><%=arrTitles(24)%></b></td>
				</tr>             
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_obj_var%>" style="padding-right:15px;padding-left:15px;">
					<!--INPUT type=image src="../netcom/images/icon_find.gif" onclick='window.open("../netcom/members/companies/cities_list.asp?fieldID=contact_city_Name&OrgId=<%=OrgId%>","CitiesList","left=100,top=50,width=450,height=450,scrollbars=1");return false;' title="<%=arrTitles(19)%>"-->     
					<input type=text dir="<%=dir_var%>" class="Field_Title" name="contact_city_Name" id="contact_city_Name" value="<%=contact_city_Name%>" style="width:200">
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;"><!--עיר--><b><%=arrTitles(25)%></b></td>
				</tr>
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>	
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
					<input type=text dir="ltr" name="contact_zip_code" value="<%=vFix(contact_zip_code)%>" size="10" maxlength=10 class="Field_Title" ID="contact_zip_code">
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">&nbsp;<!--מיקוד--><b><%=arrTitles(28)%></b></td>
				</tr>				    				
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="contact_phone" value="<%=vFix(contact_phone)%>" maxlength="20"  style="width:150" class="Field_Title" ID="Text3">                
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">
					<b><!--טלפון--><%=arrTitles(16)%></b></td>
				</tr>
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="contact_fax" value="<%=vFix(contact_fax)%>" maxlength="20" style="width:150" class="Field_Title" ID="fax">                
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">
					<b><!--פקס--><%=arrTitles(9)%></b></td>
				</tr>						
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   				
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="contact_cellular" value="<%=vFix(contact_cellular)%>" style="width:150" maxlength=20 class="Field_Title" ID="Text4">              
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">
					<b><!--טלפון נייד--><%=arrTitles(17)%></b></td>
				</tr>		
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   				
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">                    
						<input type=text name="contact_email" dir=ltr value="<%=vFix(contact_email)%>" maxlength=50 style="width:200" class="Field_Title" ID="Text5">              
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">
					<b>Email</b></td>
				</tr>
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>				
				<tr>
					<td width="100%" class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_obj_var%>" style="padding-right:15px;padding-left:15px;">
					<!--INPUT type=image src="../netcom/images/icon_find.gif" name=types_list id="types_list" onclick='window.open("../netcom/members/companies/types_list.asp?contactID=<%=contactID%>&OrgId=<%=OrgId%>&contact_types=" + contact_types.value,"TypesList","left=300,top=100,width=250,height=350,scrollbars=1");return false;' title="<%=arrTitles(19)%>"-->&nbsp;
					<!--span class="Field_Title_R" dir="<%=dir_var%>" style="width:200;line-height:16px" name="types_names" id="types_names"><%=types%></span>
					<input type=hidden name=contact_types id=contact_types value="<%=contact_types%>"-->
					<%sqlStr = "select type_Id, type_Name from contact_type WHERE ORGANIZATION_ID=" & OrgId & " order by type_Name"
					set rs_contact_type = con.GetRecordSet(sqlStr)
					If not rs_contact_type.eof Then
					%>
					<select name="contact_types" id="contact_types" class=norm style="width:200" dir="<%=dir_var%>" size=3 multiple>
					<option value="" selected><!--בחר--><%=arrTitles(19)%></option>
					<%
					do while not rs_contact_type.eof
						typeId = trim(rs_contact_type(0))				
						type_Name = trim(rs_contact_type(1))					
					%>
					<option value="<%=typeId%>"><%=type_Name%></option>
					<%
					rs_contact_type.movenext
					loop

					%>
					</select>
					<%
					End If
					%>
					</td>
					<td align="<%=td_align%>" valign=top width=150 nowrap class="Field_Title" style="padding-right:15px;padding-left:15px;"><b><!--סיווג--><%=arrTitles(26)%></b></td>
				</tr> 
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>				 
				<tr>
				<td width="100%" class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
				<textarea class="Field_Title" dir="<%=dir_var%>" style="width:250;" rows=5 name="contact_desc" id="contact_desc"><%=trim(contact_desc_short)%></textarea>      
				</td>
				<td align="<%=td_align%>" valign=top width=150 nowrap class="Field_Title" style="padding-right:15px;padding-left:15px;"><b><!--פרטים נוספים--><%=arrTitles(27)%></b></td>
				</tr> 
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>
				<tr><td width="100%" colspan=2 class="form_header" style="padding-right:15px;padding-left:15px;" align="<%=td_align%>">&nbsp;
				<b><!--פרטי חברה--><%=arrTitles(23)%></b>&nbsp;
				</td></tr>
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>
				<tr>
					<td width="100%" class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">           
					<input type=text dir="<%=dir_var%>" name="company_name" value="<%=vFix(company_name)%>" maxlength=50 style="width:200" class="Field_Title" ID="company_name">           
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">					
					&nbsp;<b><!--שם חברה--><%=arrTitles(4)%></b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
				</tr> 
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
					<input type=text dir="<%=dir_var%>" name="address" value="<%=vFix(address)%>"  style="width:200" class="Field_Title" ID="Text6"></td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;"><!--כתובת--><b><%=arrTitles(24) & " " & Request.Cookies("bizpegasus")("CompaniesOne")%></b></td>
				</tr>             
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_obj_var%>" style="padding-right:15px;padding-left:15px;">
					<!--INPUT type=image src="../netcom/images/icon_find.gif" onclick='window.open("../netcom/members/companies/cities_list.asp?fieldID=cityName&OrgId=<%=OrgId%>","CitiesList","left=100,top=50,width=450,height=450,scrollbars=1");return false;' title="<%=arrTitles(19)%>"-->       
					<input type=text dir="<%=dir_var%>" class="Field_Title" name="cityName" id="cityName" value="<%=cityName%>" style="width:200">
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;"><!--עיר--><b><%=arrTitles(25) & " " & Request.Cookies("bizpegasus")("CompaniesOne")%></b></td>
				</tr>
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>	
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
					<input type=text dir="ltr" name="zip_code" value="<%=vFix(zip_code)%>" size="10" maxlength=10 class="Field_Title" ID="zip_code">
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">&nbsp;<!--מיקוד--><b><%=arrTitles(28)%></b></td>
				</tr>				    				
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="phone1" value="<%=vFix(phone1)%>" maxlength="20"  style="width:150" class="Field_Title" ID="Text8">                
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">
					<b><!--טלפון--><%=arrTitles(16)%> 1</b></td>
				</tr>
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="phone2" value="<%=vFix(phone2)%>" maxlength="20"  style="width:150" class="Field_Title" ID="Text9">                
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">
					<b><!--טלפון--><%=arrTitles(16)%> 2</b></td>
				</tr>				
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="fax1" value="<%=vFix(fax1)%>" maxlength="20" style="width:150" class="Field_Title" ID="fax1">                
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">
					<b><!--פקס--><%=arrTitles(9)%></b></td>
				</tr>						
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   				
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">                    
						<input type=text name="email" dir=ltr value="<%=vFix(email)%>" maxlength=50 style="width:200" class="Field_Title" ID="Text11">              
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">
					<b>Email</b></td>
				</tr>				
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>			   				
				<tr>
					<td width="100%"class="Field_Title" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="url" value="<%=vFix(url)%>" style="width:200" maxlength=100 class="Field_Title" ID="Text12">              
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="Field_Title" style="padding-right:15px;padding-left:15px;">
					<b><!--אתר--><%=arrTitles(10)%></b></td>
				</tr>
				<tr><td height=1 bgcolor=#999999 colspan=2 nowrap></td></tr>				 
				<tr>
				<td width="100%" class="Field_Title" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
				<textarea class="Field_Title" dir="<%=dir_var%>" style="width:250;" rows=5 name="company_desc" id="company_desc"><%=trim(company_desc)%></textarea>      
				</td>
				<td align="<%=td_align%>" valign=top width=150 nowrap class="Field_Title" style="padding-right:15px;padding-left:15px;"><b><!--פרטים נוספים--><%=arrTitles(27)%></b></td>
				</tr>				    					
			</table>
			</td>
			</tr>	
			<%End If%>     								
			<!-- start fields dynamics -->
			<%If IsNumeric(appID) Then%>
	        <tr><td align="right" dir="rtl"><input type="button" value="נקה טופס" class="but_menu" size="10" 
	        style="width:100px;" onclick="return ClearFields();"></td></tr>
			<%End If%>
			<tr>
				<td width="100%" align="<%=td_align%>">
					<table width=100% border=0 cellspacing=1 cellpadding=3 bgcolor=#999999>	
					<!--#INCLUDE FILE="default_fields.asp"-->
					</table>
		        </td>
		    </tr>		
		<!-- end fields dynamics -->
		<!-- Authentication picture -->
		<%If trim(UseCaptcha) = "1" Then%>
		<% Authentication_Str = RandomPassword(6)
		       Session("AuthenticationStr") = Authentication_Str	%>
		<tr><td width="100%" align="<%=td_align%>">
		<table width=100% border="0" cellspacing="0" cellpadding="1">
		<tr><td align="<%=td_align%>" class="Field_Title"><b><%if Langu = "eng" then%>Type the characters you see in the picture below.<%Else%>הקלד את התווים שאתה רואה בתמונה למטה<%End If%></b></td></tr>       
		<tr><td align="<%=td_align%>"><img src="generate-captcha.asp" border="0"/></td></tr>
		<tr><td align="<%=td_align%>"><input type="text" dir="ltr" id="txt_captcha" name="txt_captcha" class="Form" maxlength="7" size="11"></td></tr>
		</table></td></tr>
		<%End If%>
		<%if is_must_fields then %>
		<tr><td height="10" align="<%=td_align%>">&nbsp;<font color=red><!--שדות חובה&nbsp;-&nbsp;*--><%=arrTitles(20)%></font>&nbsp;</td></tr>
		<%end if%>	
		<tr><td align=center><img OnClick="return CheckFields();" style="cursor:pointer;"
		<%if Langu = "eng" then%>SRC="<%=strLocal%>/netcom/images/SEND-ENG.gif"<%else%>SRC="<%=strLocal%>/netcom/images/b-send-a.gif"<%end if%> border="0" name="submit_button" id="submit_button"></td></tr>			
		</table>		
	</form>		
	<!-- end SEKER -->
	<%Else%>
	<tr><td align="center" class="form_makdim">הטופס חסום להזנה</td></tr>
    <%End If%>       
    </td></tr>    
    <tr><td height="5" nowrap></td></tr>
	<tr><td valign="bottom" width=100% align=center>
	<table border="0" align=center cellpadding=0 cellspacing=0 width=100%>
	<!--#include file="../Netcom/members/products/bottom_inc.asp"-->
     </table>
     </td>
     </tr>
     </table>
     </td>
     </tr>
     </table>
<% If (isNumeric(quest_id) = true) And (show_sale = "True") Then%>
<script language="javascript" type="text/javascript">
<!-- 		
		function click_fun()
		{ 
		<%
			If trim(lang_id) = "1" Then
				str_alert = "!!!נא למלא את השדה"
			Else
				str_alert = "Please provide the answer!!!"
			End If	
		%> 		 	
        <% sqlStr = "SELECT Field_Id,Field_Must,Field_Type FROM Form_Field Where product_id=" & quest_id & " Order by Field_Order"
		'Response.Write sqlStr
		'Response.End
		set fields=con.GetRecordSet(sqlStr)
		 do while not fields.EOF 
		  Field_Id = fields("Field_Id")				
		  Field_Must = trim(fields("Field_Must"))
		  Field_Type = trim(fields("Field_Type"))
		  'Response.Write  Field_Type
		  If Field_Must = "True"  Then		  
		 %> 
		
     field =  document.getElementsByName("field<%=Field_Id%>");  
    <%If trim(Field_Type) = "5" Or trim(Field_Type) = "8" Or trim(Field_Type) = "9" Or trim(Field_Type) = "11" Or trim(Field_Type) = "12" Then%>			
		if(field != null)
		{
			answered = false;
			for(j=0;j<field.length;j++)
			{									
				if(field[j].checked == true)
				{
					answered = true;					
				}	
			}
			if(answered == false)
			{
				window.alert("<%=str_alert%>");			
				field[0].focus();			 
				return false;
			}
		}	  	 	   
	  <%Else%>
	    if((field != null) && document.getElementById("field<%=Field_Id%>").value == '') 
	      { document.getElementById("field<%=Field_Id%>").focus(); 
		    window.alert("<%=str_alert%>"); 
		    return false; } 
	    <%
	     End If	 
	   End If
       fields.moveNext()
	   loop
       set fields=nothing
    %>   
       return true;	}
//-->
</script>  
<%End If%>
</body>
</html>
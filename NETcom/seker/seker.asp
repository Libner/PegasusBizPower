<!--#INCLUDE FILE="../reverse.asp"-->
<!--#include file="../connect.asp"-->
<%
PEOPLE_ID = Decode(trim(Request.QueryString(Encode("C"))))
prod_id = Decode(trim(Request.QueryString(Encode("P"))))
OrgID = Decode(trim(Request.QueryString(Encode("O"))))
STRING_ID = Decode(trim(Request.QueryString(Encode("S"))))

show_sale = "True"
reason = ""
PathCalImage = "../"

if trim(prod_id) = "" OR IsNumeric(prod_id) = false then
	show_sale = "False"
	reason = "SEKER NOT FOUND"		
else
	sqlStr = "Select Product_Name, PRODUCT_TYPE, PRODUCT_DESCRIPTION, Langu, QUESTIONS_ID,USER_ID, DATE_START, DATE_END from products where product_id=" & prod_id
	'Response.Write sqlStr
	'Response.End
	set rs_products = con.GetRecordSet(sqlStr)
	if not rs_products.eof then			
			PRODUCT_TYPE = rs_products("PRODUCT_TYPE")														
			quest_id = rs_products("QUESTIONS_ID") '//sheloon number of the seker														
			sqlstr = "Select PRODUCT_NAME,PRODUCT_DESCRIPTION,Langu, FILE_ATTACHMENT, ATTACHMENT_TITLE From Products WHERE Product_ID=" & quest_id
			set rsd = con.getRecordSet(sqlstr)
			If not rsd.eof Then
				PRODUCT_DESCRIPTION = rsd(1)
				PRODUCT_NAME = trim(rsd(0))
				Langu = trim(rsd("Langu"))
				attachment_file = trim(rsd("FILE_ATTACHMENT"))
				attachment_title = trim(rsd("ATTACHMENT_TITLE"))
				if Langu = "eng" then
					dir_align = "ltr"
					td_align = "left"
					pr_language = "eng"
				else
					dir_align = "rtl"
					td_align = "right"
					pr_language = "heb"
				end if
			End If
			set rsd = Nothing	
																				
			USER_ID = trim(rs_products("USER_ID"))
			
			DATE_START = trim(rs_products("DATE_START"))
			DATE_END = trim(rs_products("DATE_END"))
							
			'Response.Write datediff("d",DATE_START,date())
							
			if datediff("d",DATE_START,date()) < 0 or datediff("d",DATE_END,date()) >= 0 then
				if datediff("d",DATE_START,date()) < 0 then
						reason = "SEKER NOT START"
						show_sale = "False"	
				elseif datediff("d",DATE_END,date()) >= 0 then
						reason = "SEKER END"																																													
						show_sale = "False"
				end if
			end if
							
			''Response.Write reason
								
			
			'//START OF check client authorization				
				if trim(PEOPLE_ID) = "" or IsNumeric(PEOPLE_ID) = "False" then
					show_sale = "False"
					reason = "CLIENT NOT AUTHORIZ"	
					'Response.Write reason
					'Response.End								
				else		
					sqlStr = "Select PEOPLE_NAME,PEOPLE_COMPANY From PRODUCT_CLIENT Where PRODUCT_ID = "& prod_id &" And PEOPLE_ID=" & PEOPLE_ID	
					'Response.Write sqlStr
					'Response.End
					set rs_CONTACT = con.GetRecordSet(sqlStr)
					if not rs_CONTACT.eof then
																												
							client_name = rs_CONTACT("PEOPLE_NAME")
							if trim(client_name) <> "" then
								client_name = "(" & client_name & ")"
							end if
														
							sqlStr = "Select APPEAL_ID from APPEALS where PEOPLE_ID="& PEOPLE_ID &" and PRODUCT_ID=" & prod_id	
							set rs_CONTACT = con.GetRecordSet(sqlStr)
							if not rs_CONTACT.eof then
								show_sale = "False" '//לא לאפשר הצבעה חוזרת ללקוח שכבר הצביע
								reason="ANSWERED"
							end if
							set rs_CONTACT = nothing	
																									
						
						else
					show_sale = "False"
					reason = "CLIENT NOT AUTHORIZ"																
				end if
				set rs_CONTACT = nothing
			end if
			'//END OF check client authorization															
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
		If not pg.eof Then
			sizeLogoPr=pg.Fields("PRODUCT_LOGO").ActualSize
		End If
		set pg = Nothing
End If

sizeLogoOrg = 0
If trim(OrgID) <> "" And sizeLogoPr = 0 Then
	sqlstring = "SELECT ORGANIZATION_LOGO FROM organizations WHERE ORGANIZATION_ID = " & OrgID
	'Response.Write sqlstring
	'Response.End
	set logo=con.GetRecordSet(sqlstring)
	if not logo.EOF  then				
		sizeLogoOrg = logo("ORGANIZATION_LOGO").ActualSize
	else
	    sizeLogoOrg = 0	
	end if
	set logo = Nothing	
End If	


%>

<html>
<head>
<title><%=Product_Name%></title>
<meta charset="windows-1255">
<link href="../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="javascript">
<!--
function CheckFields()
{
		<%If trim(addClient) = "1" Or trim(addClient) = "2" Then%>
		if (window.document.all("CONTACT_NAME").value=='')
		{
		   <%
				If trim(lang_id) = "1" Then
					str_alert = "נא למלא שם"
				Else
					str_alert = "Please, insert the full name"
				End If	
			%>
 			window.alert("<%=str_alert%>");
 			window.document.all("CONTACT_NAME").focus();
			return false;
		}	
		if (window.document.all("contact_email").value=='')
		{
		}
		else if (!checkEmail(document.all("contact_email").value))
		{
			<%
				If trim(lang_id) = "1" Then
					str_alert = "כתובת דואר אלקטרוני לא חוקית"
				Else
					str_alert = "The email address is not valid!"
				End If	
			%>
				window.alert("<%=str_alert%>");
				document.all("contact_email").focus();
				return false;
		} 
		<%End If%>	
		<%If trim(addClient) = "2" Then%>
		if (window.document.all("company_name").value=='')
		{
		   <%
				If trim(lang_id) = "1" Then
					str_alert = "נא למלא שם חברה"
				Else
					str_alert = "Please, insert the company name"
				End If	
			%>
 			window.alert("<%=str_alert%>");
 			window.document.all("company_name").focus();
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
		return click_fun();
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
//-->
</script>
</head>

<body>
<table border="0" width="650" align=center cellspacing="0" cellpadding="0">
<%If sizeLogoPr > 0 OR sizeLogoOrg > 0 Then%>
<tr>
      <td width="100%" valign="bottom" align=center>
      <table border="0" width="90%" cellspacing="1" cellpadding="1" align=center ID="Table2">
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
		<table WIDTH="90%" ALIGN="center" BORDER="0" CELLPADDING="0" cellspacing="0" ID="Table3">			
		<tr>
		<td align="<%=td_align%>" width="100%" valign="top"> 
		<table border="0" cellpadding="5" cellspacing="1" width="100%" align="<%=td_align%>" bgcolor="#FFFFFF" ID="Table4">			
		  <form action="addappeal.asp" id="form1" name="form1" method="post" onSubmit="return CheckFields()">			
			<tr>
				<td class="Field_Title" style="font-size:16pt" align=center width=100% dir="<%=dir_align%>">
				<INPUT type="hidden" id=quest_id name=quest_id value="<%=quest_id%>">
				<b><%=PRODUCT_NAME%></b>
				</td>
			</tr>							
			<%if PRODUCT_DESCRIPTION <> "" then%>
			<TR>
				<td class="form_makdim" dir="<%=dir_align%>" width=100% align="<%=td_align%>">
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
			<input type="hidden" name="O" value="<%=OrgID%>" ID="O">
			<input type="hidden" name="C" value="<%=PEOPLE_ID%>" ID="C">
			<input type="hidden" name="P" value="<%=prod_id%>" ID="P">
			<input type="hidden" name=in_date value="<%=date_day%>/<%=date_month%>/<%=date_year%>" ID="in_date">
			<input type="hidden" name="Langu" value="<%=Langu%>" ID="Langu">			
			<!-- start fields dynamics -->				
			<tr>
				<td width="100%" align="<%=td_align%>">
					<table width=100% border=0 cellspacing=1 cellpadding=3 bgcolor=#999999>	
					<!--#INCLUDE FILE="default_fields.asp"-->
					</table>
		        </td>
		    </tr>		
<!-- end fields dynamics -->
	<%If trim(Langu) = "heb" Then%>
	<tr><td colspan="4" height="10" align="<%=td_align%>"><font color=red>שדות חובה&nbsp;-&nbsp;*&nbsp;</font></td></tr>
	<%Else%>
	<tr><td colspan="4" height="10" align="<%=td_align%>"><font color=red>&nbsp;*&nbsp;-&nbsp;Required fields</font></td></tr>
	<%End If%>	
	<tr><td colspan="4" align=center><a HREF="<%=strLocal%>/default.asp" target="_self"><input type="image" <%if Langu = "eng" then%>SRC="<%=strLocal%>/netcom/images/SEND-ENG.gif"<%else%>SRC="<%=strLocal%>/netcom/images/b-send-a.gif"<%end if%> border="0" name="submit_button" id="submit_button" hspace="60"></td></tr>			
	</table></td></tr>
		</table></td></tr>
		<%Else%>
		<tr><td align="center">
			<table WIDTH="90%" BGCOLOR="#B2B2B2" ALIGN="center" border="0" CELLPADDING="0" cellspacing="0" id="tb_1" name="tb_1">			
			<tr>
				<td class="title_form" style="font-size:13pt" align=center width=100% dir="<%=dir_align%>" colspan=4><INPUT type="hidden" id=P name=P value="<%=prod_id%>">
				<%=PRODUCT_NAME%>
				</td>
			</tr>
			<%if reason = "SEKER NOT START" then%>
			<tr>
				<td class="form" align=center dir=rtl width=100% bgcolor=#F1BF3A colspan=4>
					<span style="color:#FFFFFF;font-size:12pt"><b>ההרשמה למבצע זה טרם החלה</b></span>
				</td>
			</tr>
			<%elseif reason = "SEKER END" then%>
			<tr>
				<td class="form" align=center width=100% bgcolor=#F1BF3A colspan=4>
					<span style="color:#FFFFFF;font-size:12pt"><b><%if Langu = "eng" then%>This survey is now closed for further answering.<%else%>ההרשמה למבצע זה נסגרה<%end if%></b></span>
				</td>
			</tr>			
			<%elseif reason = "ANSWERED" then %>
			<tr>
				<td class="form" align=center width=100% bgcolor=#F1BF3A colspan=4>
					<%if Langu = "eng" then%>
					<span style="color:#FFFFFF;font-size:12pt"><b>You have already provided an answer to this survey.</b><br>Please note that it is not possible to reply to the same survey twice. </span>					
					<%else%>
					<span style="color:#FFFFFF;font-size:12pt" dir="rtl"><b>לתשומת לבך, טופס  זה כבר מולא בעבר על ידך.<br>לא ניתן להשתתף פעמיים.</span>					
					<%end if%>
				</td>
			</tr>			
			<%elseif reason = "SEKER NOT FOUND" then%>	
				<tr>
					<td class="form" align=center width=100% bgcolor=#F1BF3A colspan=4>
						<span style="color:#FFFFFF;font-size:12pt"><b>מבצע חסום להרשמה</b></span>
					</td>
				</tr>
			<%elseif reason = "CLIENT NOT AUTHORIZ" then%>	
				<tr>
					<td class="form" align=center width=100% bgcolor=#F1BF3A colspan=4>						
						<span style="color:#FFFFFF;font-size:12pt"><b><%if Langu = "eng" then%>Sorry, you are not authorized to participate in this survey.<%else%>אינך מורשה להרשם למבצע זה<%end if%></b></span>
					</td>
				</tr>				
			<%end if%>						
		</table>		
		</form>			
<!-- end SEKER -->
  <%End If%>       
   
    <tr><td height="5" nowrap></td></tr>
	<tr><td valign="bottom" align="center">
	<table height="40" border="0" cellpadding=0 cellspacing=0 width=90%>
	<!--#include file="../members/products/bottom_inc.asp"-->
</table>
</td></tr></table>
<% If (isNumeric(quest_id) = true) And (show_sale = "True") Then%>
<SCRIPT LANGUAGE=javascript>
<!-- 
		<%
			If trim(Langu) = "heb" Then
				str_alert = "!!!נא למלא את השדה"
			Else
				str_alert = "Please provide the answer!!!"
			End If	
		%> 	
		document.all("submit_button").onclick = click_fun;
		function click_fun()
		{ 
        <% sqlStr = "SELECT Field_Id,Field_Must,Field_Type FROM FORM_FIELD Where product_id=" & quest_id & " Order by Field_Order"
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
				if(field(j).checked == true)
				{
					answered = true;					
				}	
			}		
			if(answered == false)
			{
				window.alert("<%=str_alert%>");			
				field(0).focus();			 
				return false;
			}
		}	  	 	   
	  <%Else%>
	    if((field != null) && document.all("field<%=Field_Id%>").value == '') 
	      { document.all("field<%=Field_Id%>").focus(); 
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
</SCRIPT>  
<%End If%>
</body>
</html>

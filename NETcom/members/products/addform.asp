<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
	If request("prodId") <> nil Then
		prodId=request("prodId")
	Else
		prodId=wprodID
	End If	
	SubjId = request("SubjId")
	id=request("id")
	place=Request.QueryString("place")
	par_id = Request.QueryString("par_id")
%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
<!--
 function CheckDel(txt) {

  <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את השדה"
     Else
		str_confirm = "Are you sure want to delete the field?"
     End If   
  %>
  return (window.confirm("<%=str_confirm%>"))    
}

function popupcal(elTarget){
		if (showModalDialog){
			var sRtn;
			sRtn = showModalDialog("../../calendar.asp",elTarget.value,"center=yes; dialogWidth=110pt; dialogHeight=134pt; status=0; help=0;");
			if (sRtn!="")
			   elTarget.value = sRtn;
		 }else
		   alert("Internet Explorer 4.0 or later is required.");
		return false;
		window.document.focus;   
}

function moveTo(fromPlace,fieldId,fromOrder)
{
	<%
		sql  = "SELECT count(Field_Id) From Form_Field Where Product_ID = "&prodID
		set  rsc = con.getRecordSet(sql)		        
		if not  rsc.eof Then         
		count_ = rsc(0) 'NUMBER Of Product_Family	        
		Else
		count_ = 0
		End If
		set rsc = nothing
	%>
	
	toPlace=document.all("toPlace-"+fromPlace).value;
	if (toPlace == '')
		alert('חובה לציין מקום להעברה') 
	else
		if (eval(toPlace)<1)
			alert('מספר שורה לא חוקי') 
		else
		if (eval(toPlace)><%=count_%>)
			alert('מספר שורה יותר גדול מכמות שדות') 
		else
			document.location.href="addform.asp?fromPlace="+fromPlace+"&toPlace="+toPlace+"&prodId=<%=prodId%>&id=" + fieldId + "&order=" + fromOrder; 
	return false
}

function getNumbers(){
	var ch=event.keyCode;
	event.returnValue =(ch >= 48 && ch <= 57);
}
var oPopup = window.createPopup();
function CheckDelField(FieldID) 
{
    var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");		
    xmlHTTP.open("POST", 'checkDeleteField.asp', false);
	xmlHTTP.send('<?xml version="1.0" encoding="UTF-8"?><request><FieldID>' + FieldID + '</FieldID></request>');	 		
    result = new String(xmlHTTP.ResponseText);
    if(result == "1")
    {
		window.alert("שים לב, קיימות תשובות במערכת לשאלה זו\n\n<%=Space(8)%>לפי כך לא ניתן למחוק את השאלה הנ''ל ");
		return false;
    }
    <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את השדה"
     Else
		str_confirm = "Are you sure want to delete the field?"
     End If   
    %>
   else return (window.confirm("<%=str_confirm%>"))        
   xmlHTTP.close();
}

//-->
</script>  

</head>
<%
	
	If trim(prodId) <> "" Then
	Set product=con.GetRecordSet("SELECT PRODUCT_NUMBER,Product_Name,Langu,ADD_CLIENT FROM Products where Product_Id="&prodId&" and ORGANIZATION_ID=" & trim(OrgID) )
	if not product.Eof then
		productName=product("Product_Name")
		productNumber=product("PRODUCT_NUMBER")
		addClient = trim(product("ADD_CLIENT"))
		if product("Langu") = "eng" then
			dir_align = "ltr"
			td_align = "left"
			pr_language = "eng"
		else
			dir_align = "rtl"
			td_align = "right"
			pr_language = "heb"
		end if
		found_product = true
	else
		found_product = false	
    end if 'not product.Eof
    Set product= nothing
    End If
	
PathCalImage = "../../"

'order_field in form  *******8
if Request.QueryString("down")<>nil then
	newPlace=CInt(place)+1
	con.ExecuteQuery("UPDATE Form_Field SET Field_Order=-10 WHERE Field_Order=" & newPlace  & " AND PRODUCT_ID = " & prodId & " and ORGANIZATION_ID=" & trim(OrgID) )
	con.ExecuteQuery("UPDATE Form_Field SET Field_Order=" & newPlace & " WHERE Field_Order=" & place & " AND PRODUCT_ID = " & prodId  & " and ORGANIZATION_ID=" & trim(OrgID) )
	con.ExecuteQuery("UPDATE Form_Field SET Field_Order=" & place & " WHERE Field_Order=-10 " & " AND PRODUCT_ID = " & prodId & " and ORGANIZATION_ID=" & trim(OrgID) )
	Response.Redirect "addform.asp?prodId="&prodId
end if

if Request.QueryString("up")<>nil then
	newPlace=CInt(place)-1
	con.ExecuteQuery("UPDATE Form_Field SET Field_Order=-10 WHERE Field_Order=" & newPlace & " AND PRODUCT_ID = " & prodId & " and ORGANIZATION_ID=" & trim(OrgID) )
	con.ExecuteQuery("UPDATE Form_Field SET Field_Order=" & newPlace & " WHERE Field_Order=" & place  & " AND PRODUCT_ID = " & prodId & " and ORGANIZATION_ID=" & trim(OrgID) )
	con.ExecuteQuery("UPDATE Form_Field SET Field_Order=" & place & " WHERE Field_Order=-10 " & " AND PRODUCT_ID = " & prodId & " and ORGANIZATION_ID=" & trim(OrgID) )
	Response.Redirect "addform.asp?prodId="&prodId
end if

    if Request.QueryString("fromPlace")<>nil and Request("toPlace")<>nil then
		fromOrder=Request.QueryString("order")
		toPlace=Request("toPlace")
		sqlstr="Select top " & toPlace & " Field_Id,Field_Order FROM Form_Field WHERE Product_ID="&prodID& " ORDER BY Field_Order"
		'Response.Write sqlstr
		'Response.End
		set ord=con.getRecordSet(sqlstr)
		ord.move cint(toPlace-1)
		if not ord.eof then
			toOrder=ord(1)
			if cInt(toOrder) > cInt(fromOrder) then
				fromValue=fromOrder
				toValue=toOrder
				operand="-"
				cond="<="
			else
				fromValue=toOrder
				toValue=fromOrder
				operand="+"
				cond="<"
			end if
			newOrder=toOrder
			'Response.Write "Update Form_Field set Field_Order=Field_Order" & operand & "1 WHERE Product_ID=" & prodID & " and Field_Order >= " & fromValue & " and Field_Order " & cond & toValue & "<bR>"
			con.ExecuteQuery("Update Form_Field set Field_Order=Field_Order" & operand & "1 WHERE Product_ID=" & prodID & " and Field_Order >= " & fromValue & " and Field_Order " & cond & toValue)
			'Response.Write "Update Form_Field set Field_Order="&newOrder&" WHERE Product_ID=" & prodID & " and Field_Id="&id & "<br>"
			con.ExecuteQuery("Update Form_Field set Field_Order="&newOrder&" WHERE Product_ID=" & prodID & " and Field_Id="&id)
			'Response.End
		end if	
		set ord=Nothing
	    Response.Redirect "addform.asp?prodId=" & prodID
    end if
'end order field in form *********

if Request("delField")<>nil then
	con.ExecuteQuery("DELETE FROM Form_field WHERE Field_Id="&Id)
	con.ExecuteQuery("DELETE FROM Form_Select WHERE Field_Id="&Id)
	con.ExecuteQuery("DELETE FROM Form_Value WHERE Field_Id="&Id)
end if

If Request.QueryString("editKey") <> nil Then    
	sqlstr = "UPDATE Form_field SET Field_Key = Field_Key - 1 WHERE  Field_Id=" & Id & " AND PRODUCT_ID = " & prodId
	con.executeQuery(sqlstr)
	Response.Redirect "addform.asp?prodID=" & prodID
End If
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 21 Order By word_id"				
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
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 1%>
<%numOfLink = 1%>
<%topLevel2 = 14 'current bar ID in top submenu - added 03/10/2019%>
<%If trim(wizard_id) = "" then%>
<!--#include file="../../top_in.asp"-->
<%End if%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<tr>   	
	<td class="page_title" width=100% colspan=2 dir="<%=dir_obj_var%>">&nbsp;<font color="#6E6DA6"><%=productName%>&nbsp;</td>
</tr>
<tr>
<td align=center valign=top width=100%>
<table cellpadding=0 cellspacing=0 width=100% border=0>
<!--------------------------------הוספת לקוח פרטי -------------------------------------------------->			
<%If trim(addClient) = "1" Then%>
	<tr>       		 		                  
	<td width=100% dir="<%=dir_var%>" colspan=6> 
	<table border="0" bgcolor=#f0f0f0 style="border: 1px solid white" cellpadding="2" cellspacing="0" width="100%" dir="<%=dir_var%>">   
	<tr><td width="100%" colspan=2 class="title_form" style="padding-right:15px;padding-left:15px;" align="<%=td_align%>">&nbsp;
				<b><span id=word21 name=word21><!--פרטים אישיים--><%=arrTitles(11)%></span></b>&nbsp;
				</td></tr>	
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">           
					<input type=text dir="<%=dir_var%>" name="contact_name" value="<%=vFix(contact_name)%>" maxlength=50 style="width:200" class="Form" ID="CONTACT_NAME">           
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap class="form" style="padding-right:15px;padding-left:15px;">					
					&nbsp;<b><span id="word14" name=word14><!--שם מלא--><%=arrTitles(12)%></span></b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
				</tr> 
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="contact_phone" value="<%=vFix(contact_phone)%>" maxlength="20" style="width:100" class="Form" ID="contact_phone">                
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">
					&nbsp;<b><span id="word16" name=word16><!--טלפון--><%=arrTitles(14)%></span></b>&nbsp;</td>
				</tr>
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   				
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="contact_cellular" value="<%=vFix(contact_cellular)%>" style="width:100" maxlength=20 class="Form" ID="contact_cellular">              
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">
					&nbsp;<b><span id="word17" name=word17><!--טלפון נייד--><%=arrTitles(15)%></span></b>&nbsp;</td>
				</tr>			
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   				
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">                    
						<input type=text name="contact_email" dir=ltr value="<%=vFix(contact_email)%>" maxlength=50 style="width:200" class="Form" ID="contact_email">              
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">
					&nbsp;<b>Email</b>&nbsp;</td>
				</tr>
			</table>
			</td>
			</tr>	
			<!--------------------------------הוספת לקוח ואיש קשר -------------------------------------------------->
			<%ElseIf trim(addClient) = "2" Then ' הוספת לקוח ארגוני%>		
			<tr>       		 		                  
			<td width=100% dir="<%=dir_var%>" colspan=6> 
			<table border="0" bgcolor=#f0f0f0 style="border: 1px solid white" cellpadding="2" cellspacing="0" width="100%" dir="<%=dir_var%>" ID="Table2">   
				<tr><td width="100%" colspan=2 class="title_form" style="padding-right:15px;padding-left:15px;" align="<%=td_align%>">&nbsp;
				<b><!--פרטים איש קשר--><%=arrTitles(17) & " " & Request.Cookies("bizpegasus")("ContactsOne")%></b>&nbsp;
				</td></tr>	
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">           
					<input type=text dir="<%=dir_var%>" name="contact_name" value="<%=vFix(contact_name)%>" maxlength=50 style="width:200" class="Form" ID="Text1">           
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">					
					&nbsp;<b><!--שם מלא--><%=arrTitles(12)%></b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
				</tr> 
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   
			    <tr>
			        <td width="100%" class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;"><span dir="rtl"><b><%if FIELD_MUST then%><font color=red>&nbsp;*&nbsp;</font><%end if%><%=breaks(Field_Title)%></b></span>
					 <INPUT type=image src="../../images/icon_find.gif" onclick='window.open("../companies/messangers_list.asp","MessangersList","left=300,top=100,width=250,height=350,scrollbars=1");return false;' title="<%=arrTitles(22)%>" tabindex=-1>&nbsp;			        
             		  <input type=text dir="<%=dir_var%>" name="messangerName" value="<%=vFix(messangerName)%>" style="width:150" maxlength=50 class="Form" ID="Text2">
             		</td>
				<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">					
				<b><!--עיסוק--><%=arrTitles(13)%></b></td>
				</tr> 
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
					<input type=text dir="<%=dir_var%>" name="contact_address" value="<%=vFix(contact_address)%>"  style="width:200" class="Form" ID="contact_address"></td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;"><!--כתובת--><b><%=arrTitles(18)%></b></td>
				</tr>             
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
					<INPUT type=image src="../../images/icon_find.gif" onclick='window.open("../companies/cities_list.asp?fieldID=contact_city_Name","CitiesList","left=100,top=50,width=450,height=450,scrollbars=1");return false;' title="<%=arrTitles(22)%>">       
					<input type=text dir="<%=dir_var%>" class="Form" name="contact_city_Name" id="contact_city_Name" value="<%=contact_city_Name%>" style="width:150">
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;"><!--עיר--><b><%=arrTitles(19)%></b></td>
				</tr>  
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>	
				<tr>
					<td width="100%"class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
					<input type=text dir="ltr" name="contact_zip_code" value="<%=vFix(contact_zip_code)%>" size="10" maxlength=10 class="Form" ID="contact_zip_code">
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">&nbsp;<!--מיקוד--><b><%=arrTitles(25)%></b></td>
				</tr>				  				
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="contact_phone" value="<%=vFix(contact_phone)%>" maxlength="20"  style="width:150" class="Form">                
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">
					<b><!--טלפון--><%=arrTitles(14)%></b></td>
				</tr>
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="contact_fax" value="<%=vFix(contact_fax)%>" maxlength="20" style="width:150" class="Form" ID="fax">                
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">
					<b><!--פקס--><%=arrTitles(16)%></b></td>
				</tr>						
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   				
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="contact_cellular" value="<%=vFix(contact_cellular)%>" style="width:150" maxlength=20 class="Form" ID="Text4">              
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">
					<b><!--טלפון נייד--><%=arrTitles(15)%></b></td>
				</tr>		
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   				
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">                    
						<input type=text name="contact_email" dir=ltr value="<%=vFix(contact_email)%>" maxlength=50 style="width:200" class="Form" ID="Text5">              
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">
					<b>Email</b></td>
				</tr>
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>				
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
					<INPUT type=image src="../../images/icon_find.gif" name=types_list id="types_list" onclick='window.open("../companies/types_list.asp?contactID=<%=contactID%>&contact_types=" + contact_types.value,"TypesList","left=300,top=100,width=250,height=350,scrollbars=1");return false;' title="<%=arrTitles(22)%>">&nbsp;
					<span class="Form_R" dir="<%=dir_var%>" style="width:200;line-height:16px" name="types_names" id="types_names"><%=types%></span>
					<input type=hidden name=contact_types id=contact_types value="<%=contact_types%>">
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap class="form" style="padding-right:15px;padding-left:15px;"><b><!--סיווג--><%=arrTitles(20)%></b></td>
				</tr> 
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>				 
				<tr>
				<td width="100%" class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
				<textarea class="Form" dir="<%=dir_var%>" style="width:250;" rows=5 name="contact_desc" id="contact_desc"><%=trim(contact_desc_short)%></textarea>      
				</td>
				<td align="<%=td_align%>" valign=top width=150 nowrap class="form" style="padding-right:15px;padding-left:15px;"><b><!--פרטים נוספים--><%=arrTitles(21)%></b></td>
				</tr> 
				<tr><td width="100%" colspan=2 class="title_form" style="padding-right:15px;padding-left:15px;" align="<%=td_align%>">&nbsp;
				<b><!--פרטי חברה--><%=arrTitles(17) & " " & Request.Cookies("bizpegasus")("CompaniesOne")%></b>&nbsp;
				</td></tr>
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">           
					<input type=text dir="<%=dir_var%>" name="company_name" value="<%=vFix(company_name)%>" maxlength=50 style="width:200" class="Form" ID="company_name">           
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">					
					&nbsp;<b><!--שם חברה--><%=arrTitles(23) & " " & Request.Cookies("bizpegasus")("CompaniesOne")%></b>&nbsp;<font style="color:#FF0000;font-size:12px">*</font></td>
				</tr> 
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
					<input type=text dir="<%=dir_var%>" name="address" value="<%=vFix(address)%>"  style="width:200" class="Form" ID="Text6"></td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;"><!--כתובת--><b><%=arrTitles(18) & " " & Request.Cookies("bizpegasus")("CompaniesOne")%></b></td>
				</tr>             
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
					<INPUT type=image src="../../images/icon_find.gif" onclick='window.open("../companies/cities_list.asp?fieldID=cityName","CitiesList","left=100,top=50,width=450,height=450,scrollbars=1");return false;' title="<%=arrTitles(22)%>">
					<input type=text dir="<%=dir_var%>" class="Form" name="cityName" id="cityName" value="<%=cityName%>" style="width:150">
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;"><!--עיר--><b><%=arrTitles(19) & " " & Request.Cookies("bizpegasus")("CompaniesOne")%></b></td>
				</tr>   
				<tr>
					<td width="100%"class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
					<input type=text dir="ltr" name="zip_code" value="<%=vFix(zip_code)%>" size="10" maxlength=10 class="Form" ID="zip_code">
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">&nbsp;<!--מיקוד--><b><%=arrTitles(25)%></b></td>
				</tr>				 				
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="phone1" value="<%=vFix(phone1)%>" maxlength="20"  style="width:150" class="Form" ID="Text8">                
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">
					<b><!--טלפון--><%=arrTitles(14)%> 1</b></td>
				</tr>
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="phone2" value="<%=vFix(phone2)%>" maxlength="20"  style="width:150" class="Form" ID="Text9">                
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">
					<b><!--טלפון--><%=arrTitles(14)%> 2</b></td>
				</tr>				
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="fax1" value="<%=vFix(fax1)%>" maxlength="20" style="width:150" class="Form" ID="fax1">                
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">
					<b><!--פקס--><%=arrTitles(16)%></b></td>
				</tr>						
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   				
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">                    
						<input type=text name="email" dir=ltr value="<%=vFix(email)%>" maxlength=50 style="width:200" class="Form" ID="Text11">              
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">
					<b>Email</b></td>
				</tr>				
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>			   				
				<tr>
					<td width="100%" class="form" align="<%=td_align%>" valign="top" dir=ltr style="padding-right:15px;padding-left:15px;">             
						<input type=text dir="ltr" name="url" value="<%=vFix(url)%>" style="width:200" maxlength=20 class="Form" ID="Text12">              
					</td>
					<td align="<%=td_align%>" valign=middle width=150 nowrap  class="form" style="padding-right:15px;padding-left:15px;">
					<b><!--אתר--><%=arrTitles(24)%></b></td>
				</tr>
				<tr><td height=1 bgcolor=white colspan=2 nowrap></td></tr>				 
				<tr>
				<td width="100%" class="form" align="<%=td_align%>" valign="top" dir="<%=dir_var%>" style="padding-right:15px;padding-left:15px;">
				<textarea class="Form" dir="<%=dir_var%>" style="width:250;" rows=5 name="company_desc" id="company_desc"><%=trim(company_desc)%></textarea>      
				</td>
				<td align="<%=td_align%>" valign=top width=150 nowrap class="form" style="padding-right:15px;padding-left:15px;"><b><!--פרטים נוספים--><%=arrTitles(21)%></b></td>
				</tr>				    					
			</table>
			</td>
			</tr>				
			<%End If%>  
<%If found_product Then%>
<form name=form1 id=form1 target=_self>
<tr>
<td width=100% valign=top>
<table width="100%" align=center border="0" cellpadding="1" cellspacing="1"  bgcolor="#ffffff" ID="Table1">			   		
<tr>
	<td width="50" nowrap class="title_sort" align=center>&nbsp;<span id=word3 name=word3><!--מחיקה--><%=arrTitles(3)%></span>&nbsp;</td>
	<td width="50" nowrap class="title_sort" align=center>&nbsp;<span id="word4" name=word4><!--עריכה--><%=arrTitles(4)%></span>&nbsp;</td>	
	<td width="50" nowrap class="title_sort" align=center>&nbsp;<span id="word5" name=word5><!--מפתח--><%=arrTitles(5)%></span>&nbsp;</td>		
	<td width="100%" class="title_sort" align="center">&nbsp;&nbsp;&nbsp;&nbsp;</td>
    <td width="90" nowrap colspan=2 class="title_sort" align="right">&nbsp;</td>
</tr>
<!--#INCLUDE FILE="getfields.asp"-->
</form>
</table></td></tr></table>
</td>
<td width=110 nowrap align=right valign=top class="td_menu">
<table cellpadding=1 cellspacing=0 width=100%>
<tr><td align="right" colspan=2 height="17" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='addfield.asp?prodId=<%=prodId%>'><span id="word6" name=word6><!--הוספת שדה--><%=arrTitles(6)%></span></a></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='addquest.asp?prodId=<%=prodId%>'><span id="word7" name=word7><!--הוספת נושא--><%=arrTitles(7)%></span></a></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='#' onclick="window.open('check_form.asp?prodId=<%=prodId%>','Preview','left=20,top=20,tollbar=0,menubar=0,scrollbars=1,resizable=0,width=660,height=500');"><span id="word8" name=word8><!--תצוגה מקדימה--><%=arrTitles(8)%></span></a></td></tr>
<tr><td align="right" colspan=2 height="10" nowrap></td></tr>
</table></td></tr>
<%End If%>

<%If trim(wizard_id) <> "" Then%>
	<tr><td dir="rtl" align="center" colspan=4 height=10></td></tr>
	<tr>
		<td dir="rtl" align="center" colspan=4>
		<input type=button id="Button1" value="<< חזור" onclick="document.location.href='../wizard/wizard_<%=wizard_id%>_<%=wizard_page_id%>.asp'" class="but_menu" style="width:80" NAME="Button1">
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<input type=button id="Button2" value="המשך >>" onclick="document.location.href='../wizard/wizard_<%=wizard_id%>_<%=wizard_page_id+1%>.asp?wprodId=<%=prodId%>'" class="but_menu" style="width:80">
		</td>
	</tr>	
<%End If%>	
<tr><td height=10 nowrap></td></tr>
</table>
</body>
<%set con=nothing%>
</html>

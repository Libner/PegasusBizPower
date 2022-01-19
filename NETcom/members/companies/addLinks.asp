<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<title>קישור איש קשר</title>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<%	
 	LinkedContactid = trim(Request.QueryString("LinkedContactid"))

	  set listContact=con.GetRecordSet("select CONTACT_NAME from CONTACTS where contact_id=" & LinkedContactid )
  if not listContact.EOF then 
          contacter = trim(listContact("CONTACT_NAME"))
  end if
     set  listContact = Nothing  

	if trim(Request.Form("cont_name")) <> "" then
		cont_name = trim(Request.Form("cont_name"))
	end if
	
		if trim(Request.QueryString("cont_name")) <> "" then
			cont_name = trim(Request.QueryString("cont_name"))
		end if
		is_search = true
		if trim(cont_name)<>"" then
		where_contact_name = " and UPPER(CONTACT_NAME) LIKE '"& UCase(sFix(cont_name)) &"%'"
	else
		where_contact_name = ""
	end if
	'response.Write "cont_name="& cont_name
	 
    
	if trim(Request.Form("phone")) <> "" then
		phone = trim(Request.Form("phone"))
		end if
		if trim(Request.QueryString("phone")) <> "" then
			phone = trim(Request.QueryString("phone"))
		end if
		is_search = true
		if trim(phone)<>"" then
		where_pelephone = " and (UPPER(phone) LIKE '%"& UCase(sFix(phone)) &"%' or cellular like '%"& UCase(sFix(phone)) &"%'  or  phone_additional LIKE  '%"& UCase(sFix(phone)) &"%') "
		
		''or cellulary....
	else
		where_pelephone = ""
	end if
	
	if Request("Page")<>"" then
	Page=request("Page")
	else
		Page=1
	end if
all= Request.QueryString("all")
	%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 13 Order By word_id"				
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
<SCRIPT LANGUAGE=javascript>
<!--


function submitenter(myfield,e)
{
	var keycode;
	if (window.event) keycode = window.event.keyCode;
	else if (e) keycode = e.which;
	else return true;

	if (keycode == 13)
	{
		myfield.form.submit();
		return false;
	}
	else
	return true;
}

//-->
</SCRIPT>
</HEAD>
<BODY style="margin:0px" onload="window.focus()">
<table width="100%" align=center cellpadding=0 cellspacing=0 border=0 ID="Table1">
<tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table2">
<tr>
	<td align="<%=align_var%>" class="page_title" nowrap width="70" style="border-right: none"><A class="but_menu" href='addLinks.asp?all=1&LinkedContactid=<%=LinkedContactid%>' target=_self id=word1 name=word1><!--הצג הכול--><%=arrTitles(1)%></a></td>
	<td align="<%=align_var%>" class="page_title" nowrap width="70" style="border-right: none"><A class="but_menu" href="#" onclick="form_search.submit();"style="width:100%"><span id="word7" name=word7><!--חפש--><%=arrTitles(7)%></span></a></td>
	<td align="<%=align_var%>" class="page_title" width="100%" style="border-left: none">&nbsp;<%=trim(Request.Cookies("bizpegasus")("ContactsMulti"))%>&nbsp;  ל-  <%=contacter%></td>
</tr>
</table>
</td></tr>
<TR>
<td>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align=center ID="Table3">
<tr>
<td width="100%" bgcolor="#E4E4E4" align="middle" valign=top><!-- start code --> 
		<table align="<%=align_var%>" border="0" cellpadding="1" cellspacing="1" bgcolor=white width="100%" dir="<%=dir_var%>" ID="Table4">
		<tbody bgcolor="#E4E4E4">
		<FORM action="addLinks.asp?contact_id=<%=contact_id%>&LinkedContactid=<%=LinkedContactid%>" method=post id=form_search name=form_search>
		<tr>
		  <td align=center class="title_sort" width=29 nowrap>&nbsp;</td>     
   
		<td >&nbsp;</td>
		 <td colspan=3 align=center>&nbsp;<INPUT type=text class="search" dir="<%=dir_obj_var%>" id="phone" name=phone value="<%=trim(vFix(phone))%>" onKeyPress="return submitenter(this,event)" style="width:150"></td>
         <td align=center><INPUT type=text class="search" dir="<%=dir_obj_var%>" id=cont_name name=cont_name value="<%=trim(vFix(cont_name))%>" onKeyPress="return submitenter(this,event)" style="width:99%"></td>
		</tr>
	
		</tbody>
	  <tr> 
      <td align=right class="title_sort"   nowrap>קישור לאיש קשר</td>
     
       <td width="150" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort">&nbsp;אימייל&nbsp;</td>
         <td width="72"  nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;טלפון נוסף</td>
  	
	    <td width="72"  nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;טלפון נייד</td>
  	    <td width="75"  nowrap class="title_sort" align="<%=align_var%>" dir="<%=dir_obj_var%>">&nbsp;טלפון ישיר</td>
        <td align="<%=align_var%>" dir="<%=dir_obj_var%>" width=175 nowrap class="title_sort"><%=trim(Request.Cookies("bizpegasus")("ContactsOne"))%></td>          
 
    </tr>  
<%' response.Write all &":"& where_pelephone &":"& where_contact_name
if all="1" then
   sqlstr = "Select Contact_ID, CONTACT_NAME,phone,cellular,email,Company_ID,phone_additional "&_		
  " From CONTACTS Where 1=1 "

 elseif where_pelephone =""  and  where_contact_name ="" then
   sqlstr = "Select Contact_ID, CONTACT_NAME,phone,cellular,email,Company_ID,phone_additional "&_		
  " From CONTACTS Where 1<>1 "
  else
  sqlstr = "Select Contact_ID, CONTACT_NAME,phone,cellular,email,Company_ID,phone_additional "&_		
  " From CONTACTS Where 1=1 "&   where_pelephone & where_contact_name & " Order By  CONTACT_NAME"	
  
 'Response.End
  end if
 
  set pr=con.GetRecordSet(sqlstr)  
  
  if not pr.eof then 
		recnom = 1		
		pr.PageSize = 10			
		pr.AbsolutePage=Page		
' Response.Write "sqlstr="& sqlstr
 
		NumberOfPages = pr.PageCount
		page_size = pr.PageSize
		prArray = pr.GetRows(page_size)	
		recCount=pr.RecordCount 	
		pr.Close
		set pr=Nothing    
		i=1
		j=0
		do while (j<=ubound(prArray,2))			
			Contact_ID = trim(prArray(0,j))
			CONTACT_NAME = trim(prArray(1, j))		
			phoneStr = trim(prArray(2, j)) 	
					cellular = trim(prArray(3, j)) 				
			email = trim(prArray(4, j)) 
		CompanyID = trim(prArray(5, j)) 
		phoneAdd = trim(prArray(6, j)) 
		
		
		
					
		%>	
		
		<tr>			
	<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top  bgcolor="#E6E6E6"><a style="text-decoration:none;" href="javascript:void(0)" onclick="javascript:window.open('newcontact.asp?companyID=<%=companyID%>&contactId=<%=Contact_ID%>&LinkedContactid=<%=LinkedContactid%>','','top=50,left=50,resizable=0,width=800,height=600');"><span style="font-size:18pt;text-decoration:none;">&#8734;</span></a></td>
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top class="form_data" bgcolor="#E6E6E6"><a  class=link_categ  href="javascript:void(0)" onclick="javascript:window.open('newcontact.asp?companyID=<%=companyID%>&contactId=<%=Contact_ID%>&LinkedContactid=<%=LinkedContactid%>','','top=50,left=50,resizable=0,width=800,height=600');">&nbsp;<%=email%>&nbsp;</a></td>			
	
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top class="form_data" bgcolor="#E6E6E6"><a class=link_categ  href="javascript:void(0)" onclick="javascript:window.open('newcontact.asp?companyID=<%=companyID%>&contactId=<%=Contact_ID%>&LinkedContactid=<%=LinkedContactid%>','','top=50,left=50,resizable=0,width=800,height=600');">&nbsp;<%=cellular%>&nbsp;</a></td>			
	
	
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top class="form_data" bgcolor="#E6E6E6"><a class=link_categ  href="javascript:void(0)" onclick="javascript:window.open('newcontact.asp?companyID=<%=companyID%>&contactId=<%=Contact_ID%>&LinkedContactid=<%=LinkedContactid%>','','top=50,left=50,resizable=0,width=800,height=600');">&nbsp;<%=phoneAdd%>&nbsp;</a></td>			
	
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top class="form_data" bgcolor="#E6E6E6"><a class=link_categ  href="javascript:void(0)" onclick="javascript:window.open('newcontact.asp?companyID=<%=companyID%>&contactId=<%=Contact_ID%>&LinkedContactid=<%=LinkedContactid%>','','top=50,left=50,resizable=0,width=800,height=600');">&nbsp;<%=phoneStr%>&nbsp;</a></td>			
			<TD align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top class="form_data" bgcolor="#E6E6E6"><a class=link_categ  href="javascript:void(0)" onclick="javascript:window.open('newcontact.asp?companyID=<%=companyID%>&contactId=<%=Contact_ID%>&LinkedContactid=<%=LinkedContactid%>','','top=50,left=50,resizable=0,width=800,height=600');">&nbsp;<%=CONTACT_NAME%>&nbsp;</a></td>
		</tr>
	<%	'pr.movenext
		recnom = recnom + 1
		j = j + 1
		loop%>	
	
	    <%if NumberOfPages > 1 then%>
		<tr>
		<td width="100%" align=middle colspan="6" nowrap dir=ltr class="card">
			<table border="0" cellspacing="0" cellpadding="2" ID="Table5">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
	           If Request.QueryString("numOfRow") <> nil Then
	               numOfRow = Request.QueryString("numOfRow")
	           Else numOfRow = 1
	           End If
	         %>
	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center"><A class=pageCounter title="לדפים הקודמים" href="addLinks.asp?all=<%=all%>&amp;cont_name=<%=vFix(cont_name)%>&amp;phone=<%=vFix(phone)%>&amp;page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td class="form" valign=top><font color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle"><p style="PADDING-RIGHT: 2px; PADDING-LEFT: 2px; FONT-WEIGHT: bold; FONT-SIZE: 9pt; BACKGROUND: #74b3d5; CURSOR: default; COLOR: #ffffff" 
                 ><%=i+10*(numOfRow-1)%></p></td>
	                  <%else%>
	                     <td align="middle"><A class=pageCounter href="addLinks.asp?all=<%=all%>&LinkedContactid=<%=LinkedContactid%>&amp;cont_name=<%=vFix(cont_name)%>&amp;phone=<%=vFix(phone)%>&amp;page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>	            
					<td class="form" valign=top><font color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center"><A class=pageCounter title="לדפים הבאים" href="addLinks.asp?all=<%=all%>&LinkedContactid=<%=LinkedContactid%>&amp;cont_name=<%=vFix(cont_name)%>&amp;phone=<%=vFix(phone)%>&amp;page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>&contact_id=<%=contact_id%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>					
<%End If%>	
	<tr><td class="card" align=center colspan=6 style="color:#6E6DA6;font-weight:600" dir="ltr"><span id=word4 name=word4><!--נמצאו--><%=arrTitles(4)%></span>&nbsp;<%=recCount%>&nbsp;<span id="word5" name=word5><!--רשומות--><%=arrTitles(5)%></span></td></tr>					
<%	else 'not pr.eof
%>
	<tr><td align=center colspan=6 bgcolor="#E4E4E4" style="font-size:12px;font-weight:600;color:red;"><span id=word6 name=word6><!--לא נמצאו רשומות--><%=arrTitles(6)%></span></td></tr>	
<%	end if 'not pr.eof %>
<!-- end code -->
</td></tr></table>
</td></tr></table>
<%set con =nothing%>
	</FORM>	
		
</BODY>
</HTML>



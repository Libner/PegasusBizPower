<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">	  
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<%	
 	contact_id = trim(Request.QueryString("contact_id"))
	
	if trim(Request.Form("comp_name")) <> "" Or trim(Request.QueryString("comp_name")) <> "" then
		comp_name = trim(Request.Form("comp_name"))
		if trim(Request.QueryString("comp_name")) <> "" then
			comp_name = trim(Request.QueryString("comp_name"))
		end if
		is_search = true
		
		where_comp_name = " and UPPER(companies.COMPANY_NAME) LIKE '%"& UCase(sFix(comp_name)) &"%'"
	else
		where_comp_name = ""
	end if
	
	if trim(Request.Form("city_name")) <> "" Or trim(Request.QueryString("city_name")) <> "" then
		city_name = trim(Request.Form("city_name"))
		if trim(Request.QueryString("city_name")) <> "" then
			city_name = trim(Request.QueryString("city_name"))
		end if
		is_search = true
		
		where_city_name = " and UPPER(companies.CITY_NAME) LIKE '"& UCase(sFix(city_name)) &"%'"
	else
		where_city_name = ""
	end if
	
	if Request("Page")<>"" then
	Page=request("Page")
	else
		Page=1
	end if
	if Request.QueryString("all") <> nil then
		all = 1
		show_all = true		
	else
		all = 0
		show_all = false
	end if	

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
	  set rstitle = Nothing	 %>
<SCRIPT LANGUAGE="javascript" type="text/javascript">
<!--
	function choose(companyId, CompanyName)
	{		
		if(window.opener.document.getElementById("companyId") != null)
		{
			window.opener.document.getElementById("companyId").value = companyId;
		}	
			
		if(window.opener.document.getElementById("CompanyName") != null)	
		{
			window.opener.document.getElementById("CompanyName").value = unescape(CompanyName);
		}	
			
		window.opener.document.getElementById("project_id").value = "";
		window.opener.document.getElementById("projectName").value = "";
		
		window.opener.document.getElementById("contactID").value = "";
		window.opener.document.getElementById("contactName").value = "";		
		
		window.opener.document.getElementById("contacter_body").style.display = "inline";
		window.opener.document.getElementById("project_body").style.display = "inline";	 
		self.close();
		return false;
	}

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
<table width="100%" align="center" cellpadding=0 cellspacing=0 border=0>
<tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
	<td align="<%=align_var%>" class="page_title" nowrap width="70" style="border-right: none"><A class="but_menu" href='companies_list.asp?all=1' target="_self" style="width: 70px;" ><!--הצג הכול--><%=arrTitles(1)%></a></td>
	<td align="<%=align_var%>" class="page_title" nowrap width="70" style="border-right: none"><A class="but_menu" href="#" onclick="form_search.submit();"style="width:100%"><!--חפש--><%=arrTitles(7)%></a></td>
	<td align="<%=align_var%>" class="page_title" width="100%" style="border-left: none">&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesMulti"))%>&nbsp;</td>
</tr>
</table>
</td></tr>
<TR>
<td>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="100%" bgcolor="#E4E4E4" align="middle" valign="top"><!-- start code --> 
		<FORM action="companies_list.asp?contact_id=<%=contact_id%>" method=post id=form_search name=form_search>
		<table align="<%=align_var%>" border="0" cellpadding="1" cellspacing="1" bgcolor=white width="100%" dir="<%=dir_var%>">
		<tbody bgcolor="#E4E4E4">
		<tr>
		 <td>&nbsp;</td>
		 <td>&nbsp;<INPUT type=text class="search" dir="<%=dir_obj_var%>" id="city_name" name=city_name value="<%=trim(vFix(city_name))%>" onKeyPress="return submitenter(this,event)" style="width:99%"></td>
         <td>&nbsp;<INPUT type=text class="search" dir="<%=dir_obj_var%>" id=comp_name name=comp_name value="<%=trim(vFix(comp_name))%>" onKeyPress="return submitenter(this,event)" style="width:99%"></td>
		</tr>
		</tbody>
		<tr>						
			<td class="title_sort" align="center" width="5%">&nbsp;בחר&nbsp;</td>
			<td class="title_sort" align="center" width="20%">&nbsp;<!--עיר--><%=arrTitles(2)%>&nbsp;</td>			
			<td class="title_sort" align="center" nowrap width="75%">&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;</td>			
		</tr>
<%sqlstr = "SELECT COMPANY_ID, COMPANY_NAME, city_Name, address, zip_code, phone1, phone2, fax1, email, url, private_flag, " & _
	 " company_desc FROM dbo.companies WHERE ORGANIZATION_ID = " & trim(OrgID) &_
	 where_city_name & where_comp_name & " ORDER BY private_flag DESC, COMPANY_NAME"	
	'Response.Write sqlstr
	'Response.End
	set pr=con.GetRecordSet(sqlstr)    
	if not pr.eof then 
		recnom = 1		
		pr.PageSize = 20			
		pr.AbsolutePage=Page		
		NumberOfPages = pr.PageCount
		page_size = pr.PageSize
		prArray = pr.GetRows(page_size)	
		recCount=pr.RecordCount 	
		pr.Close
		set pr=Nothing    
		i=1
		j=0
		do while (j<=ubound(prArray,2))			
			COMPANY_ID = trim(prArray(0,j))
			COMPANY_NAME = trim(prArray(1, j))					
			cityName = trim(prArray(2, j))	
			address = trim(prArray(3, j))
			zip_code = trim(prArray(4, j))			
			phone1 = trim(prArray(5, j)) 				
			phone2 = trim(prArray(6, j)) 			
			fax1 = trim(prArray(7, j)) 
			email = trim(prArray(8, j)) 
			url = trim(prArray(9, j)) 
			private_flag = trim(prArray(10, j))
			company_desc = trim(prArray(11, j))		%>	
		<tr>			
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data" bgcolor="#E6E6E6"><input type="button" ID="btnSend" NAME="btnSend" value="בחר" onclick="return choose('<%=COMPANY_ID%>', escape('<%=altFix(COMPANY_NAME)%>'));" class="but_site"></td>
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data" bgcolor="#E6E6E6"><A class="link_categ" onclick="return choose('<%=COMPANY_ID%>', escape('<%=altFix(COMPANY_NAME)%>'));" href="javascript:void(0)">&nbsp;<%=cityName%>&nbsp;</a></td>			
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign="top" class="form_data" bgcolor="#E6E6E6"><A class="link_categ" onclick="return choose('<%=COMPANY_ID%>', escape('<%=altFix(COMPANY_NAME)%>'));" href="javascript:void(0)">&nbsp;<%=COMPANY_NAME%>&nbsp;</a></td>
		</tr>
	<%	'pr.movenext
		recnom = recnom + 1
		j = j + 1
		loop%>			
	    <%if NumberOfPages > 1 then%>
		<tr>
		<td width="100%" align="middle" colspan="3" nowrap dir="ltr" class="card">
			<table border="0" cellspacing="0" cellpadding="2">               
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
			 <td valign="center"><A class=pageCounter title="לדפים הקודמים" href="companies_list.asp?all=<%=all%>&amp;comp_name=<%=vFix(comp_name)%>&amp;city_name=<%=vFix(city_name)%>&amp;page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>&contact_id=<%=contact_id%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td class="form" valign="top"><font color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle"><p style="PADDING-RIGHT: 2px; PADDING-LEFT: 2px; FONT-WEIGHT: bold; FONT-SIZE: 9pt; BACKGROUND: #74b3d5; CURSOR: default; COLOR: #ffffff" 
                 ><%=i+10*(numOfRow-1)%></p></td>
	                  <%else%>
	                     <td align="middle"><A class=pageCounter href="companies_list.asp?all=<%=all%>&amp;comp_name=<%=vFix(comp_name)%>&amp;city_name=<%=vFix(city_name)%>&amp;page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>&contact_id=<%=contact_id%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>	            
					<td class="form" valign="top"><font color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center"><A class=pageCounter title="לדפים הבאים" href="companies_list.asp?all=<%=all%>&amp;comp_name=<%=vFix(comp_name)%>&amp;city_name=<%=vFix(city_name)%>&amp;page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>&contact_id=<%=contact_id%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>					
<%End If%>	
	<tr><td class="card" align="center" colspan="3" style="color:#5F435A;font-weight:600" dir="ltr"><!--נמצאו--><%=arrTitles(4)%>&nbsp;<%=recCount%>&nbsp;<!--רשומות--><%=arrTitles(5)%></td></tr>					
<%Else 'not pr.eof%>
	<tr><td align="center" colspan="3" bgcolor="#E4E4E4" style="font-size:12px;font-weight:600;color:red;"><!--לא נמצאו רשומות--><%=arrTitles(6)%></td></tr>	
<%End If 'not pr.eof %>
<!-- end code -->
</form></table>
</td></tr></table>
</td></tr></table>
</BODY>
</HTML>
<%set con =nothing%>
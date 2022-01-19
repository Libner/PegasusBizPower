<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<%	
 	contact_id = trim(Request.QueryString("contact_id"))
	
	if trim(Request.Form("comp_name")) <> "" Or trim(Request.QueryString("comp_name")) <> "" then
		comp_name = trim(Request.Form("comp_name"))
		if trim(Request.QueryString("comp_name")) <> "" then
			comp_name = trim(Request.QueryString("comp_name"))
		end if
		is_search = true
		
		where_comp_name = " and UPPER(companies.COMPANY_NAME) LIKE '"& UCase(sFix(comp_name)) &"%'"
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
function go_back(recId)
{		
	private_flag = document.form2.all("private_flag"+recId).value;
	if(window.opener.document.all("companyId") != null)
	{
		window.opener.document.all("companyId").value = document.form2.all("COMPANY_ID"+recId).value;		
	}	
		
	if(window.opener.document.all("COMPANY_NAME") != null)	
	{
		window.opener.document.all("COMPANY_NAME").value = document.form2.all("COMPANY_NAME"+recId).value;
		if(private_flag == "1")
		{
			window.opener.document.all("COMPANY_NAME").readOnly = true;
			window.opener.document.all("COMPANY_NAME").className = "Form_R";
		}
		else
		{
			window.opener.document.all("COMPANY_NAME").readOnly = false;
			window.opener.document.all("COMPANY_NAME").className = "Form";
		}	
	}	
		
	if(window.opener.document.all("cityName") != null)
	{		
		window.opener.document.all("cityName").value = document.form2.all("cityName"+recId).value;		
		if(private_flag == "1")
		{
			window.opener.document.all("cityName").readOnly = true;
			window.opener.document.all("cityName").className = "Form_R";
		}
		else
		{
			window.opener.document.all("cityName").readOnly = false;
			window.opener.document.all("cityName").className = "Form";
		}	
	}	
		
	if(window.opener.document.all("address") != null)
	{
		window.opener.document.all("address").value = document.form2.all("address"+recId).value;	
		if(private_flag == "1")
		{		
			window.opener.document.all("address").readOnly = true;
			window.opener.document.all("address").className = "Form_R";
		}
		else
		{
			window.opener.document.all("address").readOnly = false;
			window.opener.document.all("address").className = "Form";
		}	
	}			
 
  
 	if(window.opener.document.all("zip_code") != null)
	{
		window.opener.document.all("zip_code").value = document.form2.all("zip_code"+recId).value;	
		if(private_flag == "1")
		{		
			window.opener.document.all("zip_code").readOnly = true;
			window.opener.document.all("zip_code").className = "Form_R";
		}
		else
		{
			window.opener.document.all("zip_code").readOnly = false;
			window.opener.document.all("zip_code").className = "Form";
		}	
	}

	if(window.opener.document.all("phone1") != null)
	{
		window.opener.document.all("phone1").value = document.form2.all("phone1"+recId).value;			
		if(private_flag == "1")
		{				
			window.opener.document.all("phone1").readOnly = true;
			window.opener.document.all("phone1").className = "Form_R";
		}
		else
		{
			window.opener.document.all("phone1").readOnly = false;
			window.opener.document.all("phone1").className = "Form";
		}	
	}			

		
	if(window.opener.document.all("phone2") != null)
	{
		window.opener.document.all("phone2").value = document.form2.all("phone2"+recId).value;				
		if(private_flag == "1")
		{						
			window.opener.document.all("phone2").readOnly = true;
			window.opener.document.all("phone2").className = "Form_R";
		}
		else
		{
			window.opener.document.all("phone2").readOnly = false;
			window.opener.document.all("phone2").className = "Form";
		}	
	}
		
	if(window.opener.document.all("fax1") != null)
	{
		window.opener.document.all("fax1").value = document.form2.all("fax1"+recId).value;
		if(private_flag == "1")
		{											
			window.opener.document.all("fax1").readOnly = true;
			window.opener.document.all("fax1").className = "Form_R";
		}
		else
		{
			window.opener.document.all("fax1").readOnly = false;
			window.opener.document.all("fax1").className = "Form";
		
		}	
	}		
	
	if(window.opener.document.all("company_email") != null)	
	{
		window.opener.document.all("company_email").value= window.document.all("email" + recId).value;
		if(private_flag == "1")
		{													
			window.opener.document.all("company_email").readOnly = true;
			window.opener.document.all("company_email").className = "Form_R";
		}
		else
		{
			window.opener.document.all("company_email").readOnly = false;
			window.opener.document.all("company_email").className = "Form";
		}			
	}	
	
		
	if(window.opener.document.all("url") != null)	
	{
		window.opener.document.all("url").value= window.document.all("url" + recId).value;
		if(private_flag == "1")
		{															
			window.opener.document.all("url").readOnly = true;
			window.opener.document.all("url").className = "Form_R";
		}
		else
		{
			window.opener.document.all("url").readOnly = false;
			window.opener.document.all("url").className = "Form";		
		}	
	}		

	
	if(window.opener.document.all("company_desc") != null)	
	{
		window.opener.document.all("company_desc").value= window.document.all("company_desc" + recId).value;
		if(private_flag == "1")
		{															
			window.opener.document.all("company_desc").style.height = "85px";
			window.opener.document.all("company_desc").style.lineHeight = "16px";
			window.opener.document.all("company_desc").readOnly = true;
			window.opener.document.all("company_desc").className = "Form_R";
		}
		else
		{
			window.opener.document.all("company_desc").readOnly = false;
			window.opener.document.all("company_desc").className = "Form";		
		}	
	}
	
	if(window.opener.document.all("cities_list") != null)	
	{
	    if(private_flag == "1")
		{															
			window.opener.document.all("cities_list").disabled = true;
			window.opener.document.all("cities_list").style.display = "none";
			window.opener.document.all("cities_list").style.visible = false;
		}
		else
		{
			window.opener.document.all("cities_list").disabled = false;
			window.opener.document.all("cities_list").style.display = "inline";
			window.opener.document.all("cities_list").style.visible = true;		
		}	
	}
	
	window.opener.focus();
	window.close();
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
<table width="100%" align=center cellpadding=0 cellspacing=0 border=0>
<tr><td>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr>
	<td align="<%=align_var%>" class="page_title" nowrap width="70" style="border-right: none"><A class="but_menu" href='companies_list.asp?all=1' target=_self id=word1 name=word1><!--הצג הכול--><%=arrTitles(1)%></a></td>
	<td align="<%=align_var%>" class="page_title" nowrap width="70" style="border-right: none"><A class="but_menu" href="#" onclick="form_search.submit();"style="width:100%"><span id="word7" name=word7><!--חפש--><%=arrTitles(7)%></span></a></td>
	<td align="<%=align_var%>" class="page_title" width="100%" style="border-left: none">&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesMulti"))%>&nbsp;</td>
</tr>
</table>
</td></tr>
<TR>
<td>
<table border="0" width="100%" cellspacing="0" cellpadding="0" align=center>
<tr>
<td width="100%" bgcolor="#E4E4E4" align="middle" valign=top><!-- start code --> 
		<table align="<%=align_var%>" border="0" cellpadding="1" cellspacing="1" bgcolor=white width="100%" dir="<%=dir_var%>">
		<tbody bgcolor="#E4E4E4">
		<FORM action="companies_list.asp?contact_id=<%=contact_id%>" method=post id=form_search name=form_search>
		<tr>
		 <td>&nbsp;<INPUT type=text class="search" dir="<%=dir_obj_var%>" id="city_name" name=city_name value="<%=trim(vFix(city_name))%>" onKeyPress="return submitenter(this,event)" style="width:100%"></td>
         <td>&nbsp;<INPUT type=text class="search" dir="<%=dir_obj_var%>" id=comp_name name=comp_name value="<%=trim(vFix(comp_name))%>" onKeyPress="return submitenter(this,event)" style="width:100%"></td>
		</tr>
		</FORM>
		</tbody>
		<tr>						
			<td class=title_sort align=center width="25%">&nbsp;<span id=word2 name=word2><!--עיר--><%=arrTitles(2)%></span>&nbsp;</td>			
			<td class=title_sort align=center nowrap width="75%">&nbsp;<%=trim(Request.Cookies("bizpegasus")("CompaniesOne"))%>&nbsp;</td>			
		</tr>
		<FORM action="" method=post id=form2 name=form2 onsubmit="return false;">	
<% 
  
  sqlstr = "Select COMPANY_ID, COMPANY_NAME, city_Name,address, zip_code, "&_		
  " phone1,phone2,fax1,email,url,private_flag "&_
  " From companies Where ORGANIZATION_ID = " & trim(OrgID) &_
  where_city_name & where_comp_name & " Order By private_flag DESC, COMPANY_NAME"	
  'Response.Write sqlstr
  'Response.End
  set pr=con.GetRecordSet(sqlstr)  
  
  if not pr.eof then 
		recnom = 1		
		pr.PageSize = 10			
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
			If isNumeric(trim(COMPANY_ID)) Then			
				company_desc = ""
				sqlStr = "SELECT company_desc FROM companies WHERE company_id="& COMPANY_ID 
				'Response.Write "test"
				'Response.End
				set rs_d=con.GetRecordSet(sqlStr)
				If not rs_d.EOF then				
  					company_desc = trim(rs_d(0))
  				Else
  					company_desc = ""	
				End if 
				set rs_d = Nothing		
			End If    	
					
		%>	
		<INPUT type="hidden" id="COMPANY_ID<%=recnom%>" name="COMPANY_ID<%=recnom%>" value="<%=trim(vFix(COMPANY_ID))%>">		
		<INPUT type="hidden" id="COMPANY_NAME<%=recnom%>" name="COMPANY_NAME<%=recnom%>" value="<%=vFix(COMPANY_NAME)%>">		
		<INPUT type="hidden" id="types<%=recnom%>" name="types<%=recnom%>" value="<%=vFix(types)%>">		
		<INPUT type="hidden" id="cityName<%=recnom%>" name="cityName<%=recnom%>" value="<%=vFix(cityName)%>">
		<INPUT type="hidden" id="address<%=recnom%>" name="address<%=recnom%>" value="<%=vFix(address)%>">				
		<INPUT type="hidden" id="prefix_phone1<%=recnom%>" name="prefix_phone1<%=recnom%>" value="<%=vFix(prefix_phone1)%>">
		<INPUT type="hidden" id="phone1<%=recnom%>" name="phone1<%=recnom%>" value="<%=vFix(phone1)%>">
		<INPUT type="hidden" id="prefix_phone2<%=recnom%>" name="prefix_phone2<%=recnom%>" value="<%=vFix(prefix_phone2)%>">
		<INPUT type="hidden" id="phone2<%=recnom%>" name="phone2<%=recnom%>" value="<%=vFix(phone2)%>">
		<INPUT type="hidden" id="prefix_fax1<%=recnom%>" name="prefix_fax1<%=recnom%>" value="<%=vFix(prefix_fax1)%>">
		<INPUT type="hidden" id="fax1<%=recnom%>" name="fax1<%=recnom%>" value="<%=vFix(fax1)%>">
		<INPUT type="hidden" id="email<%=recnom%>" name="email<%=recnom%>" value="<%=vFix(email)%>">
		<INPUT type="hidden" id="url<%=recnom%>" name="url<%=recnom%>" value="<%=vFix(url)%>">
		<INPUT type="hidden" id="zip_code<%=recnom%>" name="zip_code<%=recnom%>" value="<%=vFix(zip_code)%>">
		<INPUT type="hidden" id="company_desc<%=recnom%>" name="company_desc<%=recnom%>" value="<%=vFix(company_desc)%>">
		<INPUT type="hidden" id="private_flag" name="private_flag<%=recnom%>" value="<%=vFix(private_flag)%>">		
		<tr>			
			<td align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top class="form_data" bgcolor="#E6E6E6"><A class=link_categ onclick="return go_back(<%=recnom%>);" href="#0">&nbsp;<%=cityName%>&nbsp;</a></td>			
			<TD align="<%=align_var%>" dir="<%=dir_obj_var%>" valign=top class="form_data" bgcolor="#E6E6E6"><A class=link_categ onclick="return go_back(<% =recnom%>);" href="#0">&nbsp;<%=COMPANY_NAME%>&nbsp;</a></td>
		</tr>
	<%	'pr.movenext
		recnom = recnom + 1
		j = j + 1
		loop%>	
		</FORM>	
		
	    <%if NumberOfPages > 1 then%>
		<tr>
		<td width="100%" align=middle colspan="6" nowrap dir=ltr class="card">
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
	             <td class="form" valign=top><font color="#001c5e">[</font></td>
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
					<td class="form" valign=top><font color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center"><A class=pageCounter title="לדפים הבאים" href="companies_list.asp?all=<%=all%>&amp;comp_name=<%=vFix(comp_name)%>&amp;city_name=<%=vFix(city_name)%>&amp;page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>&contact_id=<%=contact_id%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>					
<%End If%>	
	<tr><td class="card" align=center colspan=3 style="color:#6E6DA6;font-weight:600" dir="ltr"><span id=word4 name=word4><!--נמצאו--><%=arrTitles(4)%></span>&nbsp;<%=recCount%>&nbsp;<span id="word5" name=word5><!--רשומות--><%=arrTitles(5)%></span></td></tr>					
<%	else 'not pr.eof
%>
	<tr><td align=center colspan=3 bgcolor="#E4E4E4" style="font-size:12px;font-weight:600;color:red;"><span id=word6 name=word6><!--לא נמצאו רשומות--><%=arrTitles(6)%></span></td></tr>	
<%	end if 'not pr.eof %>
<!-- end code -->
</td></tr></table>
</td></tr></table>
<%set con =nothing%>
</BODY>
</HTML>

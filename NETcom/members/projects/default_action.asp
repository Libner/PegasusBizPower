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
    var oPopup_Status = window.createPopup();
	function StatusDropDown(obj)
	{
	    oPopup_Status.document.body.innerHTML = Status_Popup.innerHTML;	        	 	   	      			
	    oPopup_Status.show(0-60+obj.offsetWidth, 17, 60, 89, obj); 	   
	    return true;   
	}
//-->	
</script> 
</head> 
<%
status = trim(Request.QueryString("status"))
search = false

if trim(Request("search_company")) <> "" then		
	search_company = trim(Request("search_company"))
	If trim(search_company) <> "כללי" Then
		where_company = " And company_Name LIKE '%"& sFix(search_company) &"%'"		
	Else
		where_company = " And projects.company_id = 0"	
	End If
	status = "all"
	search = true
else
	where_company = ""	
	search_company = ""	
end if

if trim(Request("search_project")) <> "" then		
	search_project = trim(Request("search_project"))	
	where_project = " And project_Name LIKE '%"& sFix(search_project) &"%'"			
	status = "all"	
	search = true
else
	where_project = ""	
	search_project = ""			
end if

if trim(lang_id) = "1" then
  arr_status = Array("","עתידי","בביצוע","סגור")
else
  arr_status = Array("","new","active","close")
end if

If trim(status) = "" Then
	where_status = " AND projects.STATUS IN ('1','2')"	
	If lang_id  = "1" Then
	status_name_ = "עתידיים ובביצוע"
	Else
	status_name_ = "new or active"
	End If
ElseIf trim(status) <> "all" Then
	where_status = " AND projects.STATUS = '" & status & "'"
	If lang_id  = "1" Then
		Select Case status
			Case "1" : status_name_  = "עתידיים" 
			Case "2" : status_name_  = "בביצוע" 
			Case "3" : status_name_  = "סגורים" 
		End Select
	Else
		Select Case status
			Case "1" : status_name_  = "new" 
			Case "2" : status_name_  = "active" 
			Case "3" : status_name_  = "close" 
		End Select
	End If	
	search = true
Else
	where_status = ""
	status_name_ = ""		
End If

 sort = Request.QueryString("sort")	
 if trim(sort)="" then  sort = 0  end if
 
 if trim(Request.QueryString("page"))<>"" then
    page=Request.QueryString("page")
 else
    Page=1
 end if  
 
 if trim(Request.QueryString("numOfRow"))<>"" then
    numOfRow=Request.QueryString("numOfRow")
 else
    numOfRow = 1
 end if  

urlSort="default_action.asp?status="&status&"&search_company="& search_company & "&search_project=" & search_project

dim sortby(16)	
sortby(0) = " projects.status"
sortby(1) = " start_date"
sortby(2) = " start_date DESC"
'sortby(3) = " CASE WHEN projects.company_id = 0 THEN 1 ELSE 0 END, rtrim(ltrim(company_name))"
'sortby(4) = " CASE WHEN projects.company_id = 0 THEN 1 ELSE 0 END, rtrim(ltrim(company_name)) DESC"
sortby(5) = " rtrim(ltrim(project_name))"
sortby(6) = " rtrim(ltrim(project_name)) DESC"
sortby(7) = " rtrim(ltrim(project_code))"
sortby(8) = " rtrim(ltrim(project_code)) DESC"
sortby(9) = " end_date"
sortby(10) = " end_date DESC"

if Request.QueryString("delPROJECT_ID")<>nil then			
	con.ExecuteQuery "delete from PROJECTS where PROJECT_ID=" & Request.QueryString("delPROJECT_ID")
	con.ExecuteQuery "Delete From pricing_to_jobs where pricing_id=" & Request.QueryString("delPROJECT_ID")	
	con.ExecuteQuery "Delete From pricing_to_projects where pricing_id=" & Request.QueryString("delPROJECT_ID")
end if
%>
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 5 Order By word_id"				
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
<%numOftab = 0%>
<%numOfLink = 2%>
<%topLevel2 = 12 'current bar ID in top submenu - added 03/10/2019%>

<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
   <td bgcolor="#e6e6e6" align="left"  valign="middle" nowrap>
	 <table  width="100%" border="0"  cellpadding="0" cellspacing="0">
	  <tr><td class="page_title">&nbsp;</td></tr>		   	  		       	
   </table></td></tr>         
<tr><td width=100%>
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
   <tr>    
    <td width="100%" valign="top" align="center" bgcolor="#FFFFFF">
	<table width="100%" align="center" border="0" cellpadding="0" cellspacing="1" dir="<%=dir_var%>">
	<tr style="line-height:17px">	
	<td nowrap align="center" class="title_sort">&nbsp;<span id=word3 name=word3><!--פעיל--><%=arrTitles(3)%></span>&nbsp;</td>
	<td nowrap align="center" valign="top" class="title_sort">&nbsp;</td>
	<td width=100 nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="7" OR trim(sort)="8" then%>_act<%end if%>"><%if trim(sort)="7" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word14 title="<%=arrTitles(14)%>"><%elseif trim(sort)="8" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word15 title="<%=arrTitles(15)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=7" name=word15 title="<%=arrTitles(15)%>"><%end if%><span id="word4" name=word4 dir="<%=dir_obj_var%>"><!--קוד--><%=arrTitles(4)%></span><img src="../../images/arrow_<%if trim(sort)="7" then%>bot<%elseif trim(sort)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>	
	<td width=70 nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="9" OR trim(sort)="10" then%>_act<%end if%>"><%if trim(sort)="9" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word14 title="<%=arrTitles(14)%>"><%elseif trim(sort)="10" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word15 title="<%=arrTitles(15)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=9" name=word15 title="<%=arrTitles(15)%>"><%end if%><span id="word5" name=word5 dir="<%=dir_obj_var%>"><!--סגירה--><%=arrTitles(5)%></span><img src="../../images/arrow_<%if trim(sort)="9" then%>bot<%elseif trim(sort)="10" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	<td width=70 nowrap align="center" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="1" OR trim(sort)="2" then%>_act<%end if%>"><%if trim(sort)="1" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word14 title="<%=arrTitles(14)%>"><%elseif trim(sort)="2" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word15 title="<%=arrTitles(15)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=1" name=word15 title="<%=arrTitles(15)%>"><%end if%><span id="word6" name=word6 dir="<%=dir_obj_var%>"><!--פתיחה--><%=arrTitles(6)%></span><img src="../../images/arrow_<%if trim(sort)="1" then%>bot<%elseif trim(sort)="2" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	<!--td width="100%" align="<%=align_var%>" class="title_sort">לקוח&nbsp;&nbsp;&nbsp;</td-->
	<td width="100%" nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>" class="title_sort<%if trim(sort)="5" OR trim(sort)="6" then%>_act<%end if%>"><%if trim(sort)="5" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" name=word14 title="<%=arrTitles(14)%>"><%elseif trim(sort)="6" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" name=word15 title="<%=arrTitles(15)%>"><%else%><a class="title_sort" href="<%=urlSort%>&sort=5" name=word15 title="<%=arrTitles(15)%>"><%end if%><%=trim(Request.Cookies("bizpegasus")("ActivitiesOne"))%><img src="../../images/arrow_<%if trim(sort)="5" then%>bot<%elseif trim(sort)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	<td width=50 nowrap align="<%=align_var%>" dir="<%=dir_obj_var%>"  class="title_sort" name="status_td" id="status_td">&nbsp;<span id=word7 name=word7 dir="<%=dir_obj_var%>" ><!--סט'--><%=arrTitles(7)%></span>&nbsp;<IMG name=word16 style="cursor:hand;" src="../../images/icon_find.gif" BORDER=0 title="<%=arrTitles(16)%>" align=absmiddle onmousedown="StatusDropDown(status_td)"></td>
	</tr>
<%
    If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
	    PageSize = RowsInList
    Else	
     	PageSize = 10
    End If	   
	sqlStr = "SELECT projects.project_id, projects.project_code, projects.project_name, projects.company_id,"&_
	" projects.active, projects.status, projects.start_date, projects.end_date  FROM  projects " &_
	" where projects.ORGANIZATION_ID = "& OrgID 
	sqlStr = sqlStr & where_status & where_company & where_project & " and company_id=0 "&" order by "& sortby(sort)
	'Response.Write sqlStr
	'Response.End
	set rs_PROJECTS = con.GetRecordSet(sqlStr)
	if not rs_PROJECTS.eof then
			rs_PROJECTS.PageSize=PageSize
			rs_PROJECTS.AbsolutePage=Page
			recCount=rs_PROJECTS.RecordCount 
			NumberOfPages=rs_PROJECTS.PageCount
			i=1		       
  do while (not rs_PROJECTS.EOF and i<=rs_PROJECTS.PageSize)
	PROJECT_ID = rs_PROJECTS("PROJECT_ID")
	If recCount = 1 And search = true Then ' נמצא רק פרויקט אחד
		Response.Redirect "project.asp?projectID=" & project_id
    End If
	company_id = rs_PROJECTS("company_id")
	If trim(company_id) = "0" Then
		company_name = "כללי"
	Else
		company_name = rs_PROJECTS("company_name")	
	End If
	project_code = rs_PROJECTS("project_code")
	project_name = rs_PROJECTS("project_name")
	ACTIVE = rs_PROJECTS("ACTIVE")	
	status = trim(rs_PROJECTS("status"))
	Select Case status
		Case "1" : urlLink = "project.asp?projectID=" & project_id
		Case "2" : urlLink = "project.asp?projectID=" & project_id
		Case "3" : urlLink = "project.asp?projectID=" & project_id
	End Select
	
	start_date = trim(rs_PROJECTS("start_date"))
	If IsDate(start_date) Then
		start_date	= Day(start_date) & "/" & Month(start_date) & "/" & Right(Year(start_date),2)
	End If
	end_date = trim(rs_PROJECTS("end_date"))
	If IsDate(end_date) Then
		end_date = Day(end_date) & "/" & Month(end_date) & "/" & Right(Year(end_date),2)
	End If
%>
	<tr>		
	    <td align="center" valign="top" class="card"><A href="vsbPress_PROJECT.asp?idsite=<%=PROJECT_ID%>"><%if active = "0" then%><img src="../../images/lamp_off.gif" name=word17 title="<%=arrTitles(17)%>" border="0" vspace=0><%else%><img src="../../images/lamp_on.gif" name=word18 title="<%=arrTitles(18)%>" border="0" vspace=0><%end if%></A></td>	    
		<td align="center" class="card">
		<%If trim(status) = "1" Then%>
	    <%If lang_id = "1" Then%>
			<%str = "הפעלה"%>
	    <%Else%>
	        <%str = "activate"%>
	    <%End if%>
	    <span class="status_num<%=status+1%> hand" onclick="window.open('activate.asp?PROJECT_ID=<%=PROJECT_ID%>','','top=200,left=100,width=400,height=120')"><%=str%></span>
	    <%ElseIf trim(status) = "2" Then%>
	    <%If lang_id = "1" Then%>
			<%str = "סגירה"%>
	    <%Else%>
	        <%str = "close"%>
	    <%End if%>	    
	    <span class="status_num<%=status+1%> hand" onclick="window.open('close.asp?PROJECT_ID=<%=PROJECT_ID%>','','top=200,left=100,width=400,height=120')"><%=str%></span>
	    <%Else%>
	    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	    <%End If%>
	    </td>
		<td align="<%=align_var%>" dir="<%=dir_obj_var%>"  valign="top" class="card"><A class="link_categ" href="<%=urlLink%>" target=_self><%=project_code%>&nbsp;</A></td>
		<td align="<%=align_var%>" dir="<%=dir_obj_var%>"  valign="top" class="card"><A class="link_categ" href="<%=urlLink%>" target=_self><%=end_date%></A></td>
		<td align="<%=align_var%>" dir="<%=dir_obj_var%>"  valign="top" class="card"><A class="link_categ" href="<%=urlLink%>" target=_self><%=start_date%></A></td>
		<!--td align="<%=align_var%>"  valign="top" class="card"><A class="link_categ" href="<%=urlLink%>" target=_self><%=company_name%>&nbsp;</A></td-->		
		<td align="<%=align_var%>" dir="<%=dir_obj_var%>"  valign="top" class="card"><A class="link_categ" href="<%=urlLink%>" target=_self><%=project_name%>&nbsp;</A></td>
		<td align="center" class="card"><A class="status_num<%=status%>" href="<%=urlLink%>" target=_self><%=arr_status(status)%></A></td>
	</tr>
<%
rs_PROJECTS.movenext
i=i+1
loop
if NumberOfPages > 1 then
urlSort = urlSort & "&sort=" & sort
	  %>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2" ID="Table1">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>
	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word19 title="<%=arrTitles(19)%>" href="<%=urlSort%>&page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="<%=align_var%>"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter href="<%=urlSort%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="<%=align_var%>"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="<%=align_var%>"><A class=pageCounter name=word20 title="<%=arrTitles(20)%>" href="<%=urlSort%>&page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
				<%end if%>   
				</tr> 				  	             
			</table>
		</td>
	</tr>			
	<%rs_PROJECTS.close 
	set rs_PROJECTS=Nothing
	End if
	%>
	<tr>
	   <td colspan="10" height=18 bgcolor="#e6e6e6" align=center class="card" dir="ltr" style="color:#6F6DA6;font-weight:600"><span id=word8 name=word8><!--נמצאו--><%=arrTitles(8)%></span> <%=recCount%>&nbsp;<%=trim(Request.Cookies("bizpegasus")("ActivitiesMulti"))%>&nbsp;</td>
	</tr>
	<% 
	Else %>
	   <tr>
	   <td colspan="10" align=center  class="title_sort1" dir="<%=dir_var%>"><span id="word9" name=word9><!--לא נמצאו--><%=arrTitles(9)%></span>&nbsp;&nbsp;</td>
	   </tr>
<% End If%>
</table>
</td>
<td width=100 nowrap valign=top class="td_menu">
<table cellpadding=1 cellspacing=0 width=100%>
<tr><td align="<%=align_var%>" colspan=2 class="title_search"><span id="word10" name=word10><!--חיפוש--><%=arrTitles(10)%></span></td></tr>
<FORM action="default_action.asp?sort=<%=sort%>" method=POST id="form_search_project" name="form_search_project" target="_self">   
<tr><td colspan=2 align="<%=align_var%>" class="right_bar"><%=trim(Request.Cookies("bizpegasus")("ActivitiesOne"))%></td></tr>
<tr><td align="<%=align_var%>" width="15" nowrap><input type="image" onclick="form_search_project.submit();" src="../../images/search.gif"></td>
<td align="<%=align_var%>" width="85" nowrap><input type="text" class="search" dir="<%=dir_obj_var%>" style="width:80;" value="<%=vFix(search_project)%>" name="search_project" ID="search_project"></td></tr>
</FORM>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:94;" nohref  onclick="javascript:window.open('editProject_action.asp?companyID=0','addProject','top=120,left=120,resizable=0,width=480,height=450,scrollbars=1');"><span id="word11" name=word11><!--הוסף--><%=arrTitles(11)%></span> <%=trim(Request.Cookies("bizpegasus")("ActivitiesOne"))	%></a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
</table>
</td></tr></table>
</td></tr></table>
<DIV ID="Status_Popup" STYLE="display:none;">
<div dir="<%=dir_obj_var%>" style="position:absolute; top:0; left:0; width:60; height:89; border-top:1px solid black;" >	
<%For i=1 To Ubound(arr_status)%>
 <DIV onmouseover="this.style.background='#9CCDF6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; padding:3px; cursor:hand; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black""
	ONCLICK="parent.location.href='default_action.asp?status=<%=i%>'">
    <%=trim(arr_status(i))%>
    </DIV>
<%Next%>       
   <DIV onmouseover="this.style.background='#9CCDF6'" onmouseout="this.style.background='#d3d3d3'" 
	STYLE="font-family:Arial; font-size:12px; color:#000000; height:20px; background:#d3d3d3; padding:3px; cursor:hand; border:1px solid black; padding:3px; cursor:hand; border-top:0px solid black"
	ONCLICK="parent.location.href='default_action.asp?status=all'">
    <span id=word13 name=word13><!--הכל--><%=arrTitles(13)%></span>
    </DIV>
</div>
</DIV>
</body>
</html>
<%
set con = nothing
%>
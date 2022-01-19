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
function CheckDel(str) {
  return (confirm("? האם ברצונך למחוק את הפרויקט העתידי"))    
}
<!--End-->
</script> 
</head> 
<%
where_status = " AND projects.STATUS = 1"

if trim(Request("search_company")) <> "" then		
	search_company = trim(Request("search_company"))
	If trim(search_company) <> "כללי" Then
		where_company = " And company_Name LIKE '%"& sFix(search_company) &"%'"		
	Else
		where_company = " And projects.company_id = 0"	
	End If
else
	where_company = ""	
	search_company = ""		
end if

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

urlSort="prices.asp?search_company="& search_company

dim sortby(16)	
sortby(0) = " CASE WHEN projects.company_id = 0 THEN 1 ELSE 0 END, rtrim(ltrim(company_name)), projects.status"
sortby(1) = " start_date"
sortby(2) = " start_date DESC"
sortby(3) = " CASE WHEN projects.company_id = 0 THEN 1 ELSE 0 END, rtrim(ltrim(company_name))"
sortby(4) = " CASE WHEN projects.company_id = 0 THEN 1 ELSE 0 END, rtrim(ltrim(company_name)) DESC"
sortby(5) = " rtrim(ltrim(project_name))"
sortby(6) = " rtrim(ltrim(project_name)) DESC"
sortby(7) = " rtrim(ltrim(project_code))"
sortby(8) = " rtrim(ltrim(project_code)) DESC"

if Request.QueryString("delpricing_id")<>nil then			
	con.ExecuteQuery "Delete From projects where project_id = " & Request.QueryString("delpricing_id")
	con.ExecuteQuery "Delete From pricing_to_jobs where project_id=" & Request.QueryString("delpricing_id")	
	con.ExecuteQuery "Delete From pricing_to_projects where project_id=" & Request.QueryString("delpricing_id")
end if
%>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 0%>
<%numOfLink = 6%>
<!--#include file="../../top_in.asp"-->
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
   <td bgcolor="#e6e6e6" align="left" valign="middle" nowrap>
	 <table width="100%" border="0" cellpadding="0" cellspacing="0">
	  <tr><td class="page_title">רשימת פרויקטים עתידיים&nbsp;</td></tr>		   
	  		       	
   </table></td></tr>         
<tr><td width=100%>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" ID="Table1">
   <tr>    
    <td width="100%" valign="top" align="center">
	<table width="100%" align="center" border="0" cellpadding="1" cellspacing="1" ID="Table3">
	<tr>
	<td nowrap align="center" valign="bottom" class="title_sort">מחיקה</td>
	<td nowrap align="center" valign="bottom" class="title_sort">עדכון</td>
	<td nowrap align="center" valign="top" class="title_sort">הפעלה</td>
	<td width=85 nowrap align="right" class="title_sort<%if trim(sort)="7" OR trim(sort)="8" then%>_act<%end if%>"><%if trim(sort)="7" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="8" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=7" title="למיון בסדר עולה"><%end if%>קוד פרויקט<img src="../../images/arrow_<%if trim(sort)="7" then%>bot<%elseif trim(sort)="8" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	<td width=130 nowrap align="right" class="title_sort<%if trim(sort)="5" OR trim(sort)="6" then%>_act<%end if%>"><%if trim(sort)="5" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="6" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=5" title="למיון בסדר עולה"><%end if%>שם פרויקט<img src="../../images/arrow_<%if trim(sort)="5" then%>bot<%elseif trim(sort)="6" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	<td width="100%" align="right" class="title_sort<%if trim(sort)="3" OR trim(sort)="4" then%>_act<%end if%>"><%if trim(sort)="3" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort+1%>" title="למיון בסדר יורד"><%elseif trim(sort)="4" then%><a class="title_sort" href="<%=urlSort%>&sort=<%=sort-1%>" title="למיון בסדר עולה"><%else%><a class="title_sort" href="<%=urlSort%>&sort=3" title="למיון בסדר עולה"><%end if%>לקוח<img src="../../images/arrow_<%if trim(sort)="3" then%>bot<%elseif trim(sort)="4" then%>top<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2"></a></td>
	<td width=50 nowrap align="right" valign="bottom" class="title_sort" name="status_td" id="status_td">סטטוס&nbsp;</td>
	</tr>
<%
PageSize = trim(Request.Cookies("bizpegasus")("RowsInList"))
If trim(PageSize) = "" Or IsNumeric(PageSize) = false Then
    PageSize = 10
End If	
sqlStr = "SELECT projects.project_id, projects.project_code, projects.project_name, projects.company_id,"&_
" projects.active, projects.status, projects.start_date, COMPANIES.COMPANY_NAME FROM  projects LEFT OUTER JOIN " &_
" COMPANIES ON projects.company_id = COMPANIES.COMPANY_ID where projects.ORGANIZATION_ID = "& OrgID 
sqlStr = sqlStr & where_status & where_company & " order by "& sortby(sort)
'Response.Write sqlStr
set rs_PROJECTS = con.GetRecordSet(sqlStr)
if not rs_PROJECTS.eof then
		rs_PROJECTS.PageSize=PageSize
		rs_PROJECTS.AbsolutePage=Page
		recCount=rs_PROJECTS.RecordCount 
		NumberOfPages=rs_PROJECTS.PageCount
		i=1		    		  
  while (not rs_PROJECTS.EOF and i<=rs_PROJECTS.PageSize)
	PROJECT_ID = rs_PROJECTS("PROJECT_ID")
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
		Case "1" : status_name  = "עתידי"
		Case "2" : status_name  = "בביצוע"
		Case "3" : status_name  = "סגור"
	End Select
%>
	<tr>
		<td align="center" valign="top" class="card"><a href="prices.asp?delpricing_id=<%=PROJECT_ID%>&pricing=1" ONCLICK="return CheckDel()"><img src="../../images/delete_icon.gif" border="0" alt="מחיקה"></a></td>
		<td align="center" valign="top" class="card"><a href="addProject.asp?project_id=<%=PROJECT_ID%>&pricing=1"><img src="../../images/edit_icon.gif" border="0" alt="עדכון"></a></td>
	    <td align="center" valign="top" class="card"><INPUT type=image src="../../images/ok_icon.gif" title="הפעלת פרויקט" border="0" onclick="window.open('activate.asp?PROJECT_ID=<%=PROJECT_ID%>','','top=100,left=100,width=400,height=120')"></td>
		<td align="right" valign="top" class="card"><A class="link_categ" href="addProject.asp?project_id=<%=PROJECT_ID%>&pricing=1" target=_self><%=project_code%>&nbsp;</A></td>
		<td align="right" valign="top" class="card"><A class="link_categ" href="addProject.asp?project_id=<%=PROJECT_ID%>&pricing=1" target=_self><%=project_name%>&nbsp;</A></td>
		<td align="right" valign="top" class="card"><A class="link_categ" href="addProject.asp?project_id=<%=PROJECT_ID%>&pricing=1" target=_self><%=company_name%>&nbsp;</A></td>
		<td align="right" valign="top" class="card"><A class="status_num<%=status%>" href="addProject.asp?project_id=<%=PROJECT_ID%>&pricing=1" target=_self><%=status_name%>&nbsp;</A></td>
	</tr>
<%	
rs_PROJECTS.movenext
i=i+1
Wend
if NumberOfPages > 1 then
urlSort = urlSort & "&sort=" & sort
	  %>
	  <tr>
		<td width="100%" align=middle colspan="10" nowrap  bgcolor="#e6e6e6">
			<table border="0" cellspacing="0" cellpadding="2" ID="Table4">               
	        <% If NumberOfPages > 10 Then 
	              num = 10 : numOfRows = cInt(NumberOfPages / 10)
	           else num = NumberOfPages : numOfRows = 1    	                      
	           End If
            %>
	         
	         <tr>
	         <%if numOfRow <> 1 then%> 
			 <td valign="center" bgcolor="#e6e6e6" align="right"><A class=pageCounter title="לדפים הקודמים" href="<%=urlSort%>&page=<%=10*(numOfRow-1)-9%>&amp;numOfRow=<%=numOfRow-1%>" >&lt;&lt;</a></td>			                
			  <%end if%>
	             <td><font size="2" color="#001c5e">[</font></td>
	               <%for i=1 to num
	                  If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	                  if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		                 <td align="middle" bgcolor="#e6e6e6" align="right"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	                  <%else%>
	                     <td align="middle"  bgcolor="#e6e6e6" align="right"><A class=pageCounter href="<%=urlSort%>&page=<%=i+10*(numOfRow-1)%>&amp;numOfRow=<%=numOfRow%>" ><%=i+10*(numOfRow-1)%></a></td>
	                  <%end if
	                  end if
	               next%>
	            
					<td bgcolor="#e6e6e6" align="right"><font size="2" color="#001c5e">]</font></td>
				<%if NumberOfPages > cint(num * numOfRow) then%>  
					<td valign="center" bgcolor="#e6e6e6" align="right"><A class=pageCounter title="לדפים הבאים" href="<%=urlSort%>&page=<%=10*(numOfRow) + 1%>&amp;numOfRow=<%=numOfRow+1%>" >&gt;&gt;</a></td>
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
	   <td colspan="10" height=18 bgcolor="#e6e6e6" align=center class="card" dir="ltr" style="color:#6F6DA6;font-weight:600">נמצאו <%=recCount%> פרויקטים עתידיים</td>
	</tr>
	<% 
	Else %>
	   <tr>
	   <td colspan="10" align=center  class="title_sort1" dir="rtl">לא נמצאו פרויקטים עתידיים</td>
	   </tr>
<% End If%>
</table>
</td>
<td width=100 nowrap valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100% ID="Table2">
<tr><td align="right" colspan=2 class="title_search">חיפוש</td></tr>
<FORM action="prices.asp?sort=<%=sort%>" method=POST id=form_search name=form_search target="_self">   
<tr><td colspan=2 align="right" class="right_bar">שם לקוח</td></tr>
<tr>
<td align=right><input type="image" onclick="form_search.submit();" src="../../images/search.gif" ID="Image1" NAME="Image1"></td>
<td align=right><input type="text" class="search" dir="rtl" style="width:75;" value="<%=vFix(search_company)%>" name="search_company" ID="Text1"></td></tr>
</FORM>
<tr><td align="right" colspan=2 height="10" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:96;"  href='addProject.asp?pricing=1'>הוספת פרויקט</a></td></tr>
<tr><td align="right" colspan=2 height="10" nowrap></td></tr>
</table>
</td></tr></table>
</td></tr></table>
</body>
</html>
<%
set con = nothing
%>
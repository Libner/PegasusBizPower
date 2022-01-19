<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	query_groupId = trim(Request.QueryString("groupId"))  
	delGroup = trim(Request.QueryString("delGroup"))
     
  if trim(delGroup) = "True" then
		groupId = cInt(trim(Request("groupId")))
		sqlStr = "DELETE FROM sms_peoples WHERE (ORGANIZATION_ID="&OrgID&") AND (group_id="&groupId&");" &_
		"DELETE FROM sms_groups WHERE (group_id="&groupId&") AND (ORGANIZATION_ID = "&OrgID&")"
		con.ExecuteQuery(sqlStr)
		
		Response.Redirect "groups.asp"
		
 end if

if trim(Request.QueryString("page"))<>"" then
    page = cInt(Request.QueryString("page"))
else
    Page=1
end if     
 
if trim(Request.QueryString("numOfRow"))<>"" then
    numOfRow = cInt(Request.QueryString("numOfRow"))
else
    numOfRow = 1
end if

if trim(Request.QueryString("sort"))<>"" then
    sort = cInt(Request.QueryString("sort"))
else
    sort = 6
end if
			
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
function goForms(groupId)
{
	mywindCust = window.open("groupEdit.asp?groupId=" + groupId ,'groupEdit',"alwaysRaised,left=180,top=200,height=200,width=450,status=no,toolbar=no,menubar=no,location=no");
	return false;
}

function chkDelete(groupId) {
   <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את הקבוצה"
     Else
		str_confirm = "Are you sure want to delete the group ?"
     End If   
   %>		
	if (confirm("<%=str_confirm%>"))		
	{
		window.document.location.href = "groups.asp?groupId=" + groupId + "&delGroup=True"
		return false;
	}
	return false;	
}
//-->
</SCRIPT>

</head>
<body>
<!-- #include file="../../logo_top.asp" -->
<%numOftab = 46%>
<%numOfLink = 0%>
<%topLevel2 = 48 'current bar ID in top submenu - added 03/10/2019%>
<!--#include file="../../top_in.asp"-->
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF dir="<%=dir_var%>">
<tr>
<td width="100%" class="page_title" dir="<%=dir_obj_var%>" colspan=2>&nbsp;&nbsp;</td></tr>         
<tr><td width="100%" valign="top">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td width="100%" valign="top" align="<%=align_var%>">    
<%'//start of groups%>	
<table border=0 width=100% cellpadding=0 cellspacing=1>
<%urlSort = "groups.asp?1=1"
	If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
		PageSize = RowsInList
	Else	
    	PageSize = 10
	End If		
	
	sqlstr = "Exec dbo.sms_get_groups " & Page & "," & PageSize & "," & OrgID & ",'" & sort & "'"	
	set rs_groups = con.GetRecordSet(sqlStr)														
%>
<tr>
	<td width=50 class="title_sort" align=center nowrap>מחיקה</td>
	<td width=50 class="title_sort" align=center nowrap>פרטים</td>
	<td width=90 align=center nowrap class="title_sort<%if sort=5 or sort=6 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=6 then%>5<%elseif sort=5 then%>6<%else%>6<%end if%>" target="_self">תאריך יצירה<img src="../../images/arrow_<%if trim(sort)="5" then%>down<%elseif trim(sort)="6" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>
	<td width=90 align=center nowrap class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=4 then%>3<%elseif sort=3 then%>4<%else%>4<%end if%>" target="_self">כמות<img src="../../images/arrow_<%if trim(sort)="3" then%>down<%elseif trim(sort)="4" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>							
	<td width=100% align="<%=align_var%>" class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self">שם קבוצה<img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>
</tr>					
<%If not rs_groups.eof Then	   		
		group_count = cLng(rs_groups("CountRecords"))
		NumberOfPages = Fix((group_count / PageSize) + 0.99)
		i=1
		j=0
		do while not rs_groups.eof
			groupId = trim(rs_groups("group_id"))
			groupName = trim(rs_groups("group_name"))
			group_date = trim(rs_groups("group_date"))
			If IsDate(group_date) And trim(group_date) <> "" Then
				group_date = DateValue(Day(group_date) & "/" & Month(group_date) & "/" & Year(group_date))
			End if
			count_people = trim(rs_groups("count_people"))%>		
	<tr>
		<td class="card" align=center><a href="" ONCLICK="return chkDelete(<%=groupId%>);"><IMG class=img_class SRC=../../images/delete_icon.gif BORDER=0 ALT="מחיקה"></a></td>							
		<td class="card" align=center><a href="" ONCLICK="return goForms(<%=groupId%>);"><IMG SRC="../../images/edit_icon.gif" BORDER=0 alt="עדכון קבוצה"></a></td>
		<td class="card" align=center><%=group_date%></td>
		<td class="card" align=center><%=count_people%></td>
		<td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a href="group.asp?groupId=<%=groupId%>" class="link_categ" <%if trim(groupId) = trim(query_groupId) then%>style="background-color:#6E6DA6"<%end if%>>&nbsp;<%=groupName%>&nbsp;</a></td>
	</tr>					
<%
		rs_groups.movenext
		j=j+1
		loop
		set rs_groups = nothing
%>
<% if NumberOfPages > 1 then%>
<tr class="card">
  <td width="100%" align=center nowrap class="card" colspan="12">
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
	<%If numOfRow <> 1 Then%> 
		<td valign="middle" class="card"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow-1)-9%>&numOfRow=<%=numOfRow-1%>" title="לדפים הקודמים"><b><<</b></a></td>			                
		<%end if%>
	        <td><font size="2" color="#060165">[</font></td>
	        <%for i=1 to num
	            If cInt(i+10*(numOfRow-1)) <= NumberOfPages Then
	            if CInt(Page)=CInt(i+10*(numOfRow-1)) then %>
		            <td align="center"><span class="pageCounter"><%=i+10*(numOfRow-1)%></span></td>
	            <%else%>
	                <td align="center"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=i+10*(numOfRow-1)%>&numOfRow=<%=numOfRow%>"><%=i+10*(numOfRow-1)%></a></td>
	            <%end if
	            end if
	        next%>	    
			<td><font size="2" color="#060165">]</font></td>
		<%if NumberOfPages > cint(num * numOfRow) then%>  
			<td valign="middle"><a class="pageCounter" href="<%=urlSort%>&sort=<%=sort%>&page=<%=10*(numOfRow) + 1%>&numOfRow=<%=numOfRow+1%>" title="לדפים הבאים"><b>>></b></a></td>
		<%end if%>   
		</tr> 				  	             
	</table></td></tr>				
	<%End If%>		
	<tr><td align=center height=20px class="card" colspan="12"><font style="color:#6E6DA6;font-weight:600">נמצאו&nbsp;<%=group_count%> קבוצות</font></td></tr>								
</table></td>
<%		
Else
%>
  <tr><td class="title_sort1" align=center colspan=6>לא נמצאו קבוצות</td></tr>
<%
End If
%>					
	</table></td>
<%'//end of groups%>			
</td>	
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=0 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:100;" href='#' ONCLICK="window.open('groupEdit.asp',null,'alwaysRaised,left=180,top=200,height=200,width=450,status=no,toolbar=no,menubar=no,location=no');">הוסף קבוצה</a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="5" nowrap></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table>
</body>
</html>
<%set con=nothing%>
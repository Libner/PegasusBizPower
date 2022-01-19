<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#INCLUDE FILE="../checkWorker.asp"-->
<%
	query_groupId = trim(Request.QueryString("groupId"))  
	delGroup = trim(Request.QueryString("delGroup"))
     
  if trim(delGroup) = "True" then

	sqlStr = "DELETE FROM PEOPLES where ORGANIZATION_ID="&OrgID&" and GROUPID="&query_groupId 
	con.ExecuteQuery(sqlStr)	
   
	slqStr = "DELETE FROM groups where groupId="&query_groupId&" and ORGANIZATION_ID = "&OrgID
	con.ExecuteQuery(slqStr)	
		
 end if

PEOPLE_ID = Request.QueryString("PEOPLE_ID")
delUser = Request.QueryString("delUser")
if trim(delUser) = "True" then	
	sqlStr = "Delete from PEOPLES Where PEOPLE_ID="&PEOPLE_ID&" and ORGANIZATION_ID=" & OrgID  
    con.ExecuteQuery(sqlStr) 
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
<%
	  sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 44 Order By word_id"				
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
	  set rstitle = Nothing%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
function goForms(groupId)
{
mywindCust = window.open("groupEdit.asp?groupId=" + groupId ,'groupEdit',"alwaysRaised,left=180,top=200,height=200,width=450,status=no,toolbar=no,menubar=no,location=no");

return false;
}

function goFormsClient(PEOPLE_ID,groupId,page)
{
	mywindCust = window.open("addpeople.asp?PEOPLE_ID="+ PEOPLE_ID +"&groupId=" + groupId + "&page=" + page,'addclient',"alwaysRaised,left=100,top=100,height=300,width=450,status=no,toolbar=no,menubar=no,location=no");
	return false;
}

function FailureDelete(groupId) {
   <%
     If trim(lang_id) = "1" Then
        str_confirm = "? האם ברצונך למחוק את הקבוצה"
     Else
		str_confirm = "Are you sure want to delete the group ?"
     End If   
   %>		
	if (confirm("<%=str_confirm%>"))		
	{
		window.document.location.href = "default.asp?groupId=" + groupId + "&delGroup=True"
		return false;
	}
	return false;	
}
//-->
</SCRIPT>

</head>
<body>
<!-- #include file="../../logo_top.asp" -->
<%numOftab = 2%>
<%numOfLink = 1%>
<%topLevel2 = 19 'current bar ID in top submenu - added 03/10/2019%>
<!--#include file="../../top_in.asp"-->
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top" bgcolor=#FFFFFF dir="<%=dir_var%>">
<tr>
<td width="100%" class="page_title" dir="<%=dir_obj_var%>" colspan=2>&nbsp;&nbsp;</td></tr>         
<tr><td width="100%" valign="top">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
<tr><td width="100%" valign="top" align="<%=align_var%>">    
<%'//start of groups%>	
<table border=0 width=100% cellpadding=0 cellspacing=1>
<%	urlSort = "default.asp?1=1"
	If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
		PageSize = RowsInList
	Else	
    	PageSize = 10
	End If		
	
	sqlstr = "Exec dbo.get_groups " & Page & "," & PageSize & "," & OrgID & ",'" & sort & "'"	
	set rs_groups = con.GetRecordSet(sqlStr)														
%>
<tr>
	<td width=50 class="title_sort" align=center nowrap><!--מחיקה--><%=arrTitles(3)%></td>
	<td width=50 class="title_sort" align=center nowrap><!--פרטים--><%=arrTitles(4)%></td>
	<td width=90 align=center nowrap class="title_sort<%if sort=5 or sort=6 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=6 then%>5<%elseif sort=5 then%>6<%else%>6<%end if%>" target="_self"><span id="word5" name=word5><!--תאריך יצירה--><%=arrTitles(5)%></span><img src="../../images/arrow_<%if trim(sort)="5" then%>down<%elseif trim(sort)="6" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>
	<td width=90 align=center nowrap class="title_sort<%if sort=3 or sort=4 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=4 then%>3<%elseif sort=3 then%>4<%else%>4<%end if%>" target="_self"><span id="word6" name=word6><!--כמות--><%=arrTitles(6)%></span><img src="../../images/arrow_<%if trim(sort)="3" then%>down<%elseif trim(sort)="4" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>							
	<td width=100% align="<%=align_var%>" class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self"><span id="word7" name=word7><!--שם קבוצה--><%=arrTitles(7)%></span><img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>
	<td width=90 align=center nowrap class="title_sort"><!--סוג קבוצה--><%=arrTitles(15)%></td>
</tr>					
<%
	If not rs_groups.eof Then	   		
		group_count = cLng(rs_groups("CountRecords"))
		NumberOfPages = Fix((group_count / PageSize) + 0.99)
		i=1
		j=0
		do while not rs_groups.eof
			groupId = trim(rs_groups("groupId"))
			groupName = trim(rs_groups("groupName"))
			GROUPDATE = trim(rs_groups("GROUPDATE"))
			GROUPTYPE = trim(rs_groups("GROUPTYPE"))
			If IsDate(GROUPDATE) And trim(GROUPDATE) <> "" Then
				GROUPDATE = DateValue(Day(GROUPDATE) & "/" & Month(GROUPDATE) & "/" & Year(GROUPDATE))
			End if
			count_people = trim(rs_groups("count_people"))
				
%>		
	<tr>
		<td class="card" align=center><a href="" ONCLICK="return FailureDelete(<%=groupId%>);"><IMG class=img_class SRC=../../images/delete_icon.gif BORDER=0 ALT="מחיקה"></a></td>							
		<td class="card" align=center><a href="" ONCLICK="return goForms(<%=groupId%>);"><IMG SRC="../../images/edit_icon.gif" BORDER=0 alt="עדכון קבוצה"></a></td>
		<td class="card" align=center><%=GROUPDATE%></td>
		<td class="card" align=center><%=count_people%></td>
		<td class="card" align="<%=align_var%>" dir="<%=dir_obj_var%>"><a href="<%If trim(GROUPTYPE) = "1" Then%>group.asp<%Else%>group1.asp<%End If%>?groupId=<%=groupId%>" class="link_categ" <%if trim(groupId) = trim(query_groupId) then%>style="background-color:#6E6DA6"<%end if%>>&nbsp;<%=groupName%>&nbsp;</a></td>
		<td class="card" align=center><%If trim(GROUPTYPE) = "1" Then%><img src="../../images/maatafa.gif" alt="דיוור בדואר אלקטרוני" border="0"><%Else%><img src="../../images/send_icon.gif" alt="דיוור בדואר רגיל" border="0"><%End if%></td>
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
	<tr><td align=center height=20px class="card" colspan="12"><font style="color:#6E6DA6;font-weight:600"><span id=word9 name=word9><!--נמצאו--><%=arrTitles(9)%></span> <%=group_count%> <span id=word8 name=word8><!--קבוצות--><%=arrTitles(8)%></span></font></td></tr>								
</table></td>
<%		
Else
%>
  <tr><td class="title_sort1" align=center colspan=6><!--לא נמצאו קבוצות דיוור--><%=arrTitles(10)%></td></tr>
<%
End If
%>					
	</table></td>
<%'//end of groups%>			
</td>	
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=0 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:100;" href='#' ONCLICK="window.open('groupEdit.asp',null,'alwaysRaised,left=180,top=200,height=200,width=450,status=no,toolbar=no,menubar=no,location=no');"><!--הוסף קבוצה--><%=arrTitles(11)%></a></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:100;line-height:120%" href='#' ONCLICK="window.open('groupEdit1.asp',null,'alwaysRaised,left=180,top=200,height=200,width=450,status=no,toolbar=no,menubar=no,location=no');"><!--הוסף קבוצה למשלוח דואר רגיל--><%=arrTitles(12)%></a></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:100;" onclick="javascript:window.open('blocked_mails.asp',null,'alwaysRaised,left=20,top=20,height=485,width=650,status=no,toolbar=no,menubar=no,location=no');return false;" href="#"><!--מיילים חסומים--><%=arrTitles(13)%></a></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:100;" onclick="javascript:window.open('notvalid_mails.asp',null,'alwaysRaised,left=20,top=20,height=485,width=650,status=no,toolbar=no,menubar=no,location=no');return false;" href="#">מיילים לא תקניים</a></td></tr>
<tr><td colspan=2 align="center"><a class="button_edit_1" style="width:100;line-height:120%" href='#' ONCLICK="window.open('excelUploadBlocked.asp',null,'alwaysRaised,left=100,top=200,height=250,width=500,status=no,toolbar=no,menubar=no,location=no');"><!--מחק נמענים (Excel)--><%=arrTitles(14)%></a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="5" nowrap></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</table>
</body>
</html>
<%set con=nothing%>
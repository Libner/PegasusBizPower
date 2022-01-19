<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
pageID=Request.QueryString("pageID")

found_page = false
If isNumeric(pageId) Then
  set pg=con.getRecordSet("SELECT Page_Title FROM Pages WHERE Page_Id="&pageId&" AND ORGANIZATION_ID = " & OrgID)
  If not pg.eof Then
	PageTitle=pg("Page_Title")		 
	found_page = true
  End If 	 
  set pg = Nothing  
End If%>
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 2%>
<%numOfLink = 0%>
<!--#include file="../../top_in.asp"-->
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<%If found_page Then%>
 <tr>
	<td>
	<table width=100% border="0" cellpadding=0 cellspacing=0>	
	<tr>	
	<td width=100% class="page_title" colspan=2>&nbsp;מקורות פרסום&nbsp;<font color="#6F6DA6"><%=PageTitle%></font></td>
</tr> 
 <tr>
    <td width=100% valign=top>
	<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1">	
	<tr>   
	<td class="title_sort" align=center width=50 nowrap>מחיקה</td>
	<td align="center" class="title_sort" width=80 nowrap>&nbsp;תאריך&nbsp;</td>   
	<td align="center" class="title_sort" width=200 nowrap>כניסות</td>
	<td align="right"  class="title_sort" width="100%">מקור הגעה&nbsp;</td>
	</tr>	
<%  sql="SELECT Max(counter) FROM Statistic WHERE PAGE_ID = " & pageID
	set pr=con.getRecordSet(sql)
	if not pr.EOF then	
		count_all=pr(0)
		If isNumeric(count_all) And isNULL(count_all) = false Then
			koef = cInt((160 / count_all))
		Else
			count_all=0   
			koef = 0
		End if
	else
		count_all=0   
		koef = 0
	end if
	set pr = Nothing
	'Response.Write koef
	'Response.End

   sql="SELECT counter, referer, [Date] FROM Statistic WHERE PAGE_ID = " & pageID & " Order BY [Date]"
   set pr=con.getRecordSet(sql)
   if not pr.EOF then
		arr_st = pr.getRows()	
   end if
   Set pr = Nothing
   
   If isArray(arr_st) Then
   For ss=0 To Ubound(arr_st,2)		
		count = trim(arr_st(0,ss))
		If IsNumeric(count) = False Or IsNull(count) = true Then
			count = 0
		End If
		referer = trim(arr_st(1,ss))
		If Len(referer) = 0 Or isNULL(referer) Then
			referer="כללי"
		End If
		rDate = trim(arr_st(2,ss)) 
		If isDate(rDate) Then
			rDate = Day(rDate) & "/" & Month(rDate) & "/" & Year(rDate)
		End If
		%>	
<tr> 
    <td class="card" align="center"><INPUT type=image onclick="return CheckDel('<%=catsID%>','<%=pageID%>')" border=0 hspace=0 vspace=0 src="../../images/delete_icon.gif" title="מחיקת מקור הגעה"></td>
    <td class="card" align="center"><%=rDate%></td>    	
	<td class="card" align="left" valign="middle"><img src="../../images/modul.gif" border=0 hspace=0 vspace=0 height="10" width="<%=cInt(count*koef)%>">&nbsp;<b><%=count%></b></td>
	<td class="card" align="right" valign="middle"><%=referer%></td>
</tr>
<% Next
End If%>
</table></tr></table>
</td></tr>
<%End If%>
</table>
</body>
<%set con=Nothing%>
</html>
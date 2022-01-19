<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%  prodID=Request.QueryString("prodID")
	found_product = false
	If isNumeric(prodID) Then
	set pg=con.getRecordSet("SELECT product_name FROM Products WHERE Product_ID="&prodID&" AND ORGANIZATION_ID = " & OrgID)
	If not pg.eof Then
		PrTitle=pg("product_name")		 
		found_product = true
	End If 	 
	set pg = Nothing  
	End If

	If Request.QueryString("delId")<>nil Then
		con.executeQuery "Delete From Statistic Where Statistic_ID = " & Request.QueryString("delId")
	End If
	
	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 76 Order By word_id"				
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
<script language=javascript>
<!--
function openPreview(pageId,prodId)
{
	result = window.open("../Pages/result.asp?prodId="+prodId+"&pageId="+pageId,"Result","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=780, height=540, left=2, top=2");
	return false;
}

function CheckDel(stID, prodID)
{
  var answer = window.confirm("? האם ברצונך למחוק את המקור עם סטטיסטיקות");
  if(answer == true)
  {
	  document.location.href="products.asp?delId=" + stID + "&prodID=" + prodID; 
  }
  return answer;
}
//-->
</script>
</head>
<body>
<!--#include file="../../logo_top.asp"-->
<%numOfTab = 1%>
<%numOfLink = 1%>
<%topLevel2 = 14 'current bar ID in top submenu - added 03/10/2019%>
<!--#include file="../../top_in.asp"-->
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0">
<%If found_product Then%>
 <tr>
	<td>
	<table width=100% border="0" cellpadding=0 cellspacing=0>	
	<tr>	
	<td width=100% class="page_title" colspan=2 dir="<%=dir_obj_var%>">&nbsp;<!--מקורות פרסום--><%=arrTitles(1)%>&nbsp;<font color="#6F6DA6"><%=PrTitle%></font></td>
</tr> 
 <tr>
    <td width=100% valign=top>
	<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1" dir="<%=dir_var%>">	
	<tr>   
	<td align="center" class="title_sort" width=80 nowrap>&nbsp;<!--תאריך--><%=arrTitles(12)%>&nbsp;</td>   
	<td align="center" class="title_sort" width=80 nowrap><!--הרשמות--><%=arrTitles(11)%></td>
	<td align="center" class="title_sort" width=300 nowrap><!--דף פרסום--><%=arrTitles(10)%></td>
	<td align="center" class="title_sort" width=100% nowrap><!--מקור פרסום--><%=arrTitles(2)%></td>
	</tr>	
	<%sql="SELECT Count(Statistic_ID), CONVERT(CHAR(10), [Date], 103), Page_ID, " &_
	  "(Select Page_Title From Pages Where Page_ID = Statistic.Page_ID), "&_
	  "(Select referer From statistic_from_banner Where category_ID = Statistic.category_ID) , "&_
	  "(Select page_ID From statistic_from_banner Where category_ID = Statistic.category_ID) "&_
	  "FROM Statistic WHERE Product_ID = " & prodID & " Group BY Page_ID, Category_ID, CONVERT(CHAR(10), [Date], 103) Order BY 2"
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
		rDate = trim(arr_st(1,ss)) 		
		pageID = trim(arr_st(2,ss))
		pageTitle = trim(arr_st(3,ss))
		categTitle = trim(arr_st(4,ss))
		categPageId = trim(arr_st(5,ss))		
		If Len(categTitle) = 0 Or isNULL(categTitle) Then		
			categTitle = trim(arrTitles(8))
		End If%>	
<tr> 
    <td class="card" align="center"><%=rDate%></td>    	
	<td class="card" align="center" valign="middle">&nbsp;<b><%=count%></b></td>	
	<td class="card" align="<%=align_var%>" valign="middle"><%If Len(pageID) > 0 And trim(pageID) <> "0" Then%><a href="#" onclick="return openPreview('<%=pageID%>','<%=prodID%>')" class="link_categ"><%=pageTitle%>&nbsp;</a><%End If%></td>
	<td class="card" align="<%=align_var%>"><a href="pages.asp?pageID=<%=categPageId%>" target="_self" class="link_categ"><%=categTitle%></a></td>
</tr>
<% Next
End If%>
</table></tr></table>
</td></tr>
<%End If%>
</table>
</body>
</html><%set con=Nothing%>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%

pageID=Request.QueryString("pageID")
catID=Request.QueryString("catID")
check=Request.QueryString("check")

found_page = false
If isNumeric(pageId) Then
  set pg=con.getRecordSet("SELECT Page_Title FROM Pages WHERE Page_Id="&pageId&" AND ORGANIZATION_ID = " & OrgID)
  If not pg.eof Then
	PageTitle=pg("Page_Title")		 
	found_page = true
  End If 	 
  set pg = Nothing  
End If
  
If trim(catID)<>"" AND trim(check)<>"" then
   sql="UPDATE statistic_from_banner SET isAction="& check &" WHERE category_ID="& catID
   con.getRecordSet(sql)
End If

if Request.QueryString("delId")<>nil then
    con.executeQuery "delete from statistic_from_banner_counter where category_ID=" & Request.QueryString("delId")
    con.executeQuery "delete from statistic_from_banner where category_ID=" & Request.QueryString("delId")
end if

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
<title>BizPower</title>
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--
function CheckDel(catID,pageID)
{
  var answer = window.confirm("? האם ברצונך למחוק את המקור עם סטטיסטיקות");
  if(answer == true)
  {
	  document.location.href="pages.asp?delId=" + catID + "&pageID=" + pageID; 
  }
  return answer;
}
//-->
</script>  
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 2%>
<%numOfLink = 0%>
<%topLevel2 = 18 'current bar ID in top submenu - added 03/10/2019%>
<!--#include file="../../top_in.asp"-->
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
<%If found_page Then%>
 <tr>
	<td>
	<table width=100% border="0" cellpadding=0 cellspacing=0>	
	<tr>	
	<td width=100% class="page_title" colspan=2 dir="<%=dir_obj_var%>">&nbsp;<!--מקורות פרסום--><%=arrTitles(1)%>&nbsp;<font color="#6F6DA6"><%=PageTitle%></font></td>
</tr> 
 <tr>
    <td width=100% valign=top>
	<table width="100%" align="center" border="0" cellpadding="2" cellspacing="1" dir="<%=dir_var%>">	
	<tr>   
	<td class="title_sort" align=center width=50 nowrap><!--מחיקה--><%=arrTitles(6)%></td>
	<td class="title_sort" align=center width=50 nowrap><!--עדכון--><%=arrTitles(5)%></td>
	<td align="center" class="title_sort" width=50 nowrap>&nbsp;<!--פעיל--><%=arrTitles(4)%>&nbsp;</td>   
	<td align="center" class="title_sort" width=60 nowrap><!--כניסות--><%=arrTitles(3)%></td>
	<td align="center" class="title_sort" width=320 nowrap>URL</td>
	<td align="<%=align_var%>"  class="title_sort" width="100%"><!--מקור פרסום--><%=arrTitles(2)%>&nbsp;</td>
	</tr>	
	<%sql="SELECT SUM(counter) FROM statistic_from_banner_counter WHERE category_ID = 0 AND PAGE_ID = " & pageID
	  set pr=con.getRecordSet(sql)
	  if not pr.EOF then
		catsID=0
		referer=trim(arrTitles(8))
		count=pr(0)
		If IsNumeric(count) = False Or IsNull(count) = true Then
			count = 0
		End If%>	
<tr> 
    <td class="card" colspan=2>&nbsp;</td>    	
	<td nowrap class="card" align="center" valign="middle"><input type="checkbox" name="<%=catID%>" checked class="radio" style="width:22px;height:22px;cursor:hand;" ID="<%=pageID%>" disabled></td>
	<td class="card" align="center" valign="middle">&nbsp;<b><%=count%></b></td>
	<td class="card" align="left" valign="middle"><a href="<%=strLocal & "sale.asp?" & Encode("pageID=" & pageID)%>" target=_blank class="link1"><%=strLocal & "sale.asp?" & Encode("pageID=" & pageID)%></a></td>
	<td class="card" align="<%=align_var%>" valign="middle"><!--ללא מקור פרסום--><%=arrTitles(8)%></td>
</tr>
<%	
	pr.close
	set pr = nothing
	End If
%>
<%  sql="select statistic_from_banner.category_ID,referer,isAction " &_
	",(SELECT SUM(counter) FROM statistic_from_banner_counter " &_
	" WHERE statistic_from_banner_counter.category_ID=statistic_from_banner.category_ID "&_
	" AND PAGE_ID = " & pageID & ") AS statCounter " &_
	" FROM statistic_from_banner  WHERE PAGE_ID = " & pageID & " ORDER BY category_ID"
	set pr=con.getRecordSet(sql)
	do while not pr.EOF
		catsID=pr("category_ID")
		referer=pr("referer")
		isAction=pr("isAction")
		count=pr("statCounter")
		If IsNumeric(count) = False Or IsNull(count) = true Then
			count = 0
		End If	   
%>
	<tr>
     	<td class="card" align="center" valign="middle"><INPUT type=image onclick="return CheckDel('<%=catsID%>','<%=pageID%>')" border=0 hspace=0 vspace=0 src="../../images/delete_icon.gif" title="מחיקת מקור פרסום" ID="Image1" NAME="Image1"></td>
		<td class="card" align="center" valign="middle"><INPUT type=image onclick="window.open('addCat.asp?pageID=<%=pageID%>&catID=<%=catsID%>','add','top=220,left=120,resizable=1,scrollbars=1,width=480,height=150')" src="../../images/edit_icon.gif" border=0 hspace=0 vspace=0 title="עדכון מקור פרסום" ID="Image2" NAME="Image2"></td>
		<td class="card" align="center" valign="middle"><INPUT type="checkbox" name="<%=catID%>" <%if isAction=true then%> checked onclick="window.document.location.href='pages.asp?pageID=<%=pageID%>&catID=<%=catsID%>&check=0'" <%else%> onclick="window.document.location.href='pages.asp?pageID=<%=pageID%>&catID=<%=catsID%>&check=1'" <%end if%> class="radio" style="width:22px;height:22px;cursor:hand;" ID="Checkbox1"></td>
		<td class="card" align="center" valign="middle">&nbsp;<b><%=count%></b></td>
		<td class="card" align="left" valign="middle"><a href="<%=strLocal & "sale.asp?" & Encode("pageID=" & pageID & "&fromID=" & catsID)%>" class="link1" target=_blank><%=strLocal & "sale.asp?" & Encode("pageID=" & pageID & "&fromID=" & catsID)%></a></td>
		<td class="card" align="<%=align_var%>" valign="middle"><%=referer%></td>
	</tr>
<%
	pr.movenext
	loop
	pr.close
%>
</table>
<td width=100 nowrap valign=top class="td_menu" style="border: 1px solid #808080">
<table cellpadding=1 cellspacing=0 width=100%>
<tr><td align="<%=align_var%>" colspan=2>&nbsp;</td></tr>
<tr><td  align="center" nowrap><a class="button_edit_1" style="width:94;" onClick="javascript:window.open('addCat.asp?pageID=<%=pageID%>','add','top=220,left=120,resizable=1,scrollbars=1,width=480,height=150')" nohref><!--הוספת מקור פרסום--><%=arrTitles(7)%>&nbsp;</a></td></tr>
<tr><td align="<%=align_var%>" colspan=2>&nbsp;</td></tr>
</table></td></tr>
</table></td></tr>
<%End If%>
</table>
</body>
<%set con=Nothing%>
</html>
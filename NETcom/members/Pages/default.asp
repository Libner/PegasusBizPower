<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script LANGUAGE="JavaScript">
<!--
function CheckDelProd() {
   <%
     If trim(lang_id) = "1" Then
        str_confirm = "?האם ברצונך למחוק את הדף המעוצב"
     Else
		str_confirm = "Are you sure want to delete the page ?"
     End If   
   %>		
   return window.confirm("<%=str_confirm%>");		
}
function openPreview(pageId)
{
	page = window.open("result.asp?pageId="+pageId,"Page","toolbar=0, menubar=0, resizable=1, scrollbars=1, fullscreen = 0, width=780, height=540, left=2, top=2");	
}
 function openCopy(pageId)
 {
	h = 200;
	w = 500;
	S_Wind = window.open("copyPage.asp?copyPageId=" + pageId, "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
	S_Wind.focus();	
	return false;
 }
//-->
</script>  
</head>

<%
	pageId = Request.QueryString("pageId")
	if request("delPage")="1" and isNumeric(pageId) then			
		con.ExecuteQuery("UPDATE PRODUCTS SET page_id = null where page_id="&pageid&" and ORGANIZATION_ID=" & trim(OrgID) )	
		con.ExecuteQuery("Delete from statistic_from_banner_counter WHERE page_id="&pageId)
		con.ExecuteQuery("Delete from statistic_from_banner WHERE page_id="&pageId)		
		con.ExecuteQuery("Delete from PAGES WHERE page_id="&pageId&" and ORGANIZATION_ID=" & trim(OrgID) )		
		Response.Redirect "default.asp"
	end if
  %>
  <%
  'set count_pages = con.GetRecordSet("Select page_id  from pages where ORGANIZATION_ID=" & trim(OrgID))
  'if  count_pages.eof then 
  'Response.Redirect "Addpage.asp"
  'end if
 dim sortby(16)	 
 sortby(1) = " PAGE_TITLE"
 sortby(2) = " PAGE_TITLE DESC"
 sortby(3) = " PAGE_ID"
 sortby(4) = " PAGE_ID DESC"

 sort = Request("sort")	
 if sort = nil then
	sort = 4
 end if 
 
 urlSort = "default.asp?1=1"
 if Request("Page")<>"" then
	Page=request("Page")
 else
	Page=1
 end if	    

	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 50 Order By word_id"				
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
<body>
<!--#include file="../../logo_top.asp"-->
<%numOftab = 2%>
<%numOfLink = 0%>
<%topLevel2 = 18 'current bar ID in top submenu - added 03/10/2019%>
<!--#include file="../../top_in.asp"-->
<div align="center"><center>
<table border="0" width="100%" bgcolor="#FFFFFF" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
 <tr>
	<td colspan=2>
	<table width=100% border="0" cellpadding=0 cellspacing=0>	
	<tr><td width=100% class="page_title">&nbsp;</td></tr> 
	</table></td>
</tr>	
<tr>
<td width=100% valign=top>
<table cellpadding=1 cellspacing=1 width=100% align=center>
<tr>
	<td width="100" nowrap align=center class="title_sort">&nbsp;<!--מחיקת דף--><%=arrTitles(3)%>&nbsp;</td>	
	<td width="100" nowrap align=center class="title_sort">&nbsp;<!--שכפול דף--><%=arrTitles(4)%>&nbsp;</td>
	<td width="100" nowrap align=center class="title_sort">&nbsp;<!--דוח כניסות--><%=arrTitles(12)%>&nbsp;</td>
	<td width="100" nowrap align=center class="title_sort">&nbsp;<!--מקור פרסום--><%=arrTitles(5)%>&nbsp;</td>
	<td width="100" nowrap align=center class="title_sort">&nbsp;<!--תצוגה מקדימה--><%=arrTitles(6)%>&nbsp;</td>
	<td width="70" nowrap align=center class="title_sort">&nbsp;<!--הופץ--><%=arrTitles(13)%>&nbsp;</td>
	<td width="100%" align="<%=align_var%>" class="title_sort<%if sort=1 or sort=2 then%>_act<%end if%>"><a class="title_sort" HREF="<%=urlSort%>&sort=<%if sort=2 then%>1<%elseif sort=1 then%>2<%else%>2<%end if%>" target="_self">&nbsp;<span id="word7" name=word7><!--שם דף מעוצב--><%=arrTitles(7)%></span><img src="../../images/arrow_<%if trim(sort)="1" then%>down<%elseif trim(sort)="2" then%>up<%else%>grey<%end if%>.gif" width="9" height="6" border="0" hspace="2" vspace=5 style="vertical-align:text-top"></a></td>
</tr>
<%  set rs_pages = con.GetRecordSet("Select page_id, page_title,(Select Count(Page_ID) From Products Where Products.Page_ID = Pages.Page_ID) from pages where ORGANIZATION_ID=" & trim(OrgID) & " order by " & sortby(sort))
    if not rs_pages.eof then 
		If trim(RowsInList) <> "" AND IsNumeric(RowsInList) = true Then
				PageSize = RowsInList
		Else	
     			PageSize = 10
		End If	 
		rs_pages.PageSize = PageSize
		rs_pages.AbsolutePage=Page
		recCount=rs_pages.RecordCount 	
		pages_count=rs_pages.RecordCount	
		NumberOfPages = rs_pages.PageCount
		i=1
		j=0
		do while (not rs_pages.eof and j<rs_pages.PageSize)
    		pageId = trim(rs_pages(0))
    		page_name = trim(rs_pages(1))    	
    		count_send = trim(rs_pages(2))    	    	
%>
<tr>
	<td align="center" class="card" nowrap>&nbsp;
	<%If cInt(count_send) = 0 Then%>
	<a href="default.asp?pageId=<%=pageId%>&delPage=1" ONCLICK="return CheckDelProd()"><IMG SRC="../../images/delete_icon.gif" BORDER=0 alt="<%=arrTitles(3)%>"></a>&nbsp;
	<%Else%>
	<input type=image SRC="../../images/delete_icon.gif" BORDER=0 Onclick="window.alert('<%=Space(14)%>שים לב, לא ניתן למחוק את הדף\n\nמפני שקיימות הפצות של דף זה במערכת\n\n<%=Space(28)%>על מנת למחוק את הדף\n\n<%=Space(6)%>אליך למחוק תחילה את ההפצות הנל');return false;">
	<%End If%>
	</td>		
	<td align="center" class="card" nowrap><INPUT type=image onclick="return openCopy('<%=pageId%>')" SRC="../../images/copy_icon.gif" BORDER=0 alt="<%=arrTitles(4)%>">&nbsp;</td>
	<td align="center" class="card" nowrap><a class=link1 href="reports.asp?pageID=<%=pageID%>" target=_self><img src="../../images/graph.gif" border=0 hspace=0 vspace=0 title="<%=arrTitles(12)%>"></a></td>
	<td align="center" class="card" nowrap><a class=link1 href="../statistic/pages.asp?pageID=<%=pageID%>" target=_self><img src="../../images/plus_sub.gif" border=0 hspace=0 vspace=0 title="<%=arrTitles(5)%>"></a></td>
	<td align="center" class="card" nowrap>&nbsp;<INPUT type=image OnClick="return openPreview('<%=pageId%>')" SRC="../../images/preview_icon.gif" BORDER=0 alt="עריכת מסמך">&nbsp;</td>	
	<td align="center" class="card" nowrap>
	<%If cInt(count_send) > 0 Then%>
	<img src="../../images/vi.gif" border=0 vspace=0 hspace=0>
	<%End if%>
	</td>
	<td align="<%=align_var%>" class="card" dir=rtl>	
	<a class="link_categ" <%If cInt(count_send) = 0 Then%>href="editPage.aspx?pageId=<%=pageId%>"<%Else%> href="#" onclick="window.alert('<%=Space(20)%>,לא ניתן לעדכן את הדף שהופץ בעבר\n\n.באפשרותך לשכפל את הדף ולבצע עריכה מחדש')" 
    <%End If%>><%=page_name%>&nbsp;&nbsp;</a>
	</td>
</tr>
<%		rs_pages.MoveNext
	    j=j+1
	    loop
%>
<% if NumberOfPages > 1 then%>
<tr class="card">
  <td width="100%" align=center nowrap class="card" colspan=11>
	<table border="0" cellspacing="0" cellpadding="2" ID="Table2">               
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
<tr><td align=center height=20px class="card"  colspan=11><font style="color:#6E6DA6;font-weight:600"><span id="word9" name=word9><!--נמצאו--><%=arrTitles(9)%></span> <%=pages_count%> <span id="word8" name=word8><!--דפים--><%=arrTitles(8)%></span></font></td></tr>								
<%		
Else
%>
<tr><td align="center" class="title_sort1" colspan=6><span id="word10" name=word10><!--לא נמצאו דפים מעוצבים--><%=arrTitles(10)%></span></td></tr>
<%
  end if
  set rs_pages=nothing
%>
</table>
</td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100%>
<tr><td align="<%=align_var%>" colspan=2 height="17" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='editPage.aspx'><span id="word11" name=word11><!--בנית דף מעוצב--><%=arrTitles(11)%></span></a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="10" nowrap></td></tr>
</table></td></tr>
</table></td></tr>
<tr><td height=10 nowrap></td></tr>
</center></div></table>
</body>
<%
set con=nothing%>
</html>

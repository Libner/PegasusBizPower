<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
  prodID = trim(Request.QueryString("prodID"))
  If trim(prodID) <> "" Then
	 sqlstr = "Select product_name From products WHERE product_id = " & prodID & " And Organization_ID = " & OrgID
	 set rs_prod = con.getRecordSet(sqlstr)
	 If not rs_prod.eof Then
		 prodName = rs_prod(0)
		 found_prod = true
	 Else
		 found_prod = false	
	 End if
  Else	 
	 found_prod = false
  End if
  
	If Request.QueryString("deleteId") <> nil Then
		delId = trim(Request.QueryString("deleteId"))
		If Len(delId) > 0 Then		    
		    con.executeQuery("Delete From Product_Closings Where Closing_ID = " & delId)			
		End If
		Response.Redirect "closings.asp?prodID="&prodID
	End If    
  

	sqlstr = "Select word_id,word From translate Where lang_id = " & lang_id & " And page_id = 70 Order By word_id"				
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
	
	sqlstr = "select value From buttons Where lang_id = " & lang_id & " Order By button_id"
	set rsbuttons = con.getRecordSet(sqlstr)
	If not rsbuttons.eof Then
	arrButtons = ";" & rsbuttons.getString(,,",",";")		
	arrButtons = Split(arrButtons, ";")
	End If
	set rsbuttons=nothing	 	  
%>
<html>
<head>
<title>Bizpower
- מחולל טפסים
- טפסים וסקרים
- כלים אינטרנטיים חזקים
</title>
<!-- #include file="../../title_meta_inc.asp" -->
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<script language=javascript>
<!--
	function checkDelete()
	{	
      <%
       If trim(lang_id) = "1" Then
           str_confirm = "?האם ברצונך למחוק את המענה"
       Else
		   str_confirm = "Are you sure want to delete the closing status ?"
       End If   
      %>	
	  return window.confirm("<%=str_confirm%>");	
	}
	
	function openClosing(closingID,prodID)
	{
		h = 170;
		w = 400;
		S_Wind = window.open("addclosing.asp?closingID=" + closingID + "&prodID=" + prodID, "S_Wind" ,"scrollbars=1,toolbar=0,top=150,left=50,width="+w+",height="+h+",align=center,resizable=1");
		S_Wind.focus();	
		return false;
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
<%If found_prod Then%>
<table width="100%" cellspacing="0" cellpadding="0" align=center border="0" bgcolor="#ffffff">
<tr><td class="page_title" width=100% colspan=2 dir="<%=dir_obj_var%>">&nbsp;<!--מצבי סגירה--><%=arrTitles(6)%>&nbsp;-&nbsp;<font color="#6F6DA6"><%=prodName%></font></td></tr>
<td align=center width=100% valign=top>
<table width="650" cellspacing="1" cellpadding="2" align="center" border="0" bgcolor="#ffffff" dir="<%=dir_var%>">
<tr><td colspan=3 height=18 nowrap></td></tr>
<tr>
	<td width="80" nowrap align="center" class="card3">&nbsp;<!--מחיקה--><%=arrTitles(2)%>&nbsp;</td>	
	<td width="80" nowrap align="center" class="card3">&nbsp;<!--עדכון--><%=arrTitles(3)%>&nbsp;</td>	
	<td width="100%" align="<%=align_var%>" class="card3">&nbsp;<!--מצב--><%=arrTitles(8)%>&nbsp;</td>	
</tr>
<%	sqlStr = "Select Closing_ID, Closing_Name from Product_Closings WHERE ORGANIZATION_ID= "& OrgID &_
	" And Product_ID = " & prodID & " Order By Closing_Name"
	'Response.Write sqlStr
	set rs_closings = con.GetRecordSet(sqlStr)
	If not rs_closings.eof Then
	while not rs_closings.eof
		Closing_ID = rs_closings("Closing_ID")
		Closing_Name = rs_closings("Closing_Name")
%>
<tr>
	<td align="center" class="card2" nowrap><a href="closings.asp?deleteId=<%=Closing_ID%>&prodID=<%=prodID%>" ONCLICK="return checkDelete()"><IMG SRC="../../images/delete_icon.gif" BORDER=0></a></td>
	<td align="center" class="card2" nowrap><a href="" onClick="return openClosing('<%=Closing_ID%>','<%=prodID%>')" target="_blank"><IMG SRC="../../images/edit_icon.gif" BORDER=0></a></td>	
	<td align="<%=align_var%>" class="card2" dir="<%=dir_obj_var%>">&nbsp;<b><%=Closing_Name%></b>&nbsp;</td>	
</tr>
<%			
		rs_closings.movenext				
	Wend			
	set rs_closings = nothing			
	End If
%>			
</table></td>
<td width=110 nowrap align="<%=align_var%>" valign=top class="td_menu">
<table cellpadding=1 cellspacing=1 width=100% border=0>
<tr><td align="<%=align_var%>" colspan=2 height="18" nowrap></td></tr>
<tr><td nowrap colspan=2 align="center"><a class="button_edit_1" style="width:100;"  href='#' OnClick="return openClosing('','<%=prodID%>')"><!--הוספת מצב סגירה--><%=arrTitles(7)%></a></td></tr>
<tr><td align="<%=align_var%>" colspan=2 height="18" nowrap></td></tr>
</table></td></tr>
<tr><td colspan=2 height=20 nowrap></td></tr>
</table>
<%End If%>
</body>		
</html>		
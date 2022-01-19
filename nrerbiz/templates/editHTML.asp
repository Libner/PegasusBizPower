<%SERVER.ScriptTimeout=3000%>
<!--#INCLUDE file="..\..\netcom/reverse.asp"-->
<!--#include file="..\..\netcom/connect.asp"-->
<!--#include file="../checkWorker.asp"-->
<%
   innerparent=request("innerparent")
   elemId=trim(Request("elemId"))
   pageId=trim(Request("pageId"))
   prodId=trim(Request("prodId"))
   place=CInt(Request("place"))
   newPlace=place+1
   if isNumeric(elemId) = True then
		selm="select * from Template_Elements where Element_Id="&elemId&" "
		set pr=con.getRecordSet(selm)
			pageId=pr("Page_Id")
			pertext=pr("Element_Text")
		pr.close 
   end if
   
if trim(pertext)="" or isNull(perText) then
	pertext1="<table width=100% cellpadding=0 cellspacing=0 border=0>" &chr(13)&chr(10)
	pertext1=pertext1 & "   <tr>" &chr(13)&chr(10)
	pertext1=pertext1 & vbTab & "<td width=100% align=right valign=top>"  &chr(13)&chr(10)
	pertext1=pertext1 & vbTab & "<font style='FONT-SIZE:10pt; font-family:Ariel(Hebrew)'>" &chr(13)&chr(10)
	pertext1=pertext1 & vbTab & "<!--Content of TD-->" &chr(13)&chr(10)
	pertext1=pertext1 & vbTab &chr(13)&chr(10)
	pertext1=pertext1 & vbTab & "</font></td>" &chr(13)&chr(10)	
	pertext1=pertext1 & "   </tr>" &chr(13)&chr(10)
	pertext1=pertext1 & "</table>" &chr(13)&chr(10)

	pertext2="<table width=100% cellpadding=0 cellspacing=0 border=0>" & vbCr
	pertext2=pertext2 & "   <tr>" & vbCrLf
	pertext2=pertext2 & "<!--Left column of table-->" & vbCr	
	pertext2=pertext2 & vbTab & "<td width=&quot;&quot; align=right valign=top>" & vbCrLf
	pertext2=pertext2 & "<!--Content of left TD-->" &vbCrLf & vbCrLf
	pertext2=pertext2 & vbTab & "</td>" & vbCrLf & vbCrLf
	pertext2=pertext2 & "<!--Middle column of table-->" & vbCr
	pertext2=pertext2 & vbTab & "<td width=&quot;&quot; align=right valign=top>" & vbCrLf
	pertext2=pertext2 & "<!--Content of middle TD-->" & vbCrLf & vbCrLf
	pertext2=pertext2 & vbTab & "</td>" & vbCrLf 
	pertext2=pertext2 & "<!--Right column of table-->" & vbCr
	pertext2=pertext2 & vbTab & "<td width=&quot;&quot; align=right valign=top>" & vbCrLf
	pertext2=pertext2 & "<!--Content of right TD-->" & vbCrLf & vbCrLf
	pertext2=pertext2 & vbTab & "</td>" & vbCrLf 
	pertext2=pertext2 & "   </tr>" & vbCr
	pertext2=pertext2 & "</table>"
else
	pertext1=pertext
end if
%>
<html>
<head>
<title>ניהול תבניות דפי מבצע</title>
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
<link href="../../admin.css" rel="STYLESHEET" type="text/css">
<script language="JavaScript">

function FormSubmitFun(){
//	if (CheckFields())
if (confirm("?האם ברצונך לשמור את השינוים"))
{	if (document.f1.numcolumn[1].checked) 
	{	get_mytable();	}
	document.f1.submit();
}	
return false;
}
//-->
</script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
	var html_text = new String();
	var row_number = new Number();
	var column_number = new Number();
	var border_number;
	var tbl_align;
	var tbl_width;
function button1_onclick() {
	
	row_number = document.f1.numrow.value;
	column_number = document.f1.columnnumb.value;
	border_number = document.f1.numborder.value;
	tbl_align = document.f1.tbl_align.value;
	tbl_width = document.f1.numwidth.value;
	html_text = "";
	html_text = "<TABLE BORDER="+ border_number +" width="+ tbl_width +"% align="+ tbl_align +" CELLSPACING=1 CELLPADDING=1>";
	for (i=0;i<row_number;i++)
	{
		html_text = html_text + "<tr>";
		for (j=0;j<column_number;j++)
		{
		html_text = html_text + "<td>";
		html_text = html_text + "<textarea dir=rtl name=txtarea" + i + "_" + j + " rows=3 style='width:95%'></textarea>";
		html_text = html_text + "</td>";
		}
		html_text = html_text + "</tr>"	;
	}
	html_text = html_text + "</TABLE>";
	mytable.innerHTML = html_text;
	if (document.f1.txtarea0_0)
	{	document.f1.txtarea0_0.scrollIntoView();}
	//alert(html_text);
}

function numcolumn_onclick() {
	
	if (document.f1.numcolumn[0].checked){
		document.all("mytable").innerHTML="";
		document.all("textPict").style.display='none';
		document.all("text_html").style.display='inline';
	}
	else
	{
		document.all("textPict").style.display='block';
		document.all("text_html").style.display='none';
	}	
}

function get_mytable() {
	
	widCol=100/column_number;
	html_text = "";
	html_text = "<TABLE BORDER="+ border_number +" width="+ tbl_width +"% align="+ tbl_align +" CELLSPACING=1 CELLPADDING=1>";
	for (i=0;i<row_number;i++)
	{
		html_text = html_text + "<tr>";
		for (j=0;j<column_number;j++)
		{
		html_text = html_text + "<td width="+widCol+"% nowrap align=right valign=top>";
		if (document.f1.Template_Elements['txtarea'+ i + '_' + j].value == '')
		{
			html_text = html_text + "<font style='FONT-SIZE:10pt; font-family:Ariel(Hebrew)'><p>&nbsp;</p></font>";
		}	
		else
		{
			html_text = html_text + "<p>" + document.f1.elements['txtarea'+ i + '_' + j].value + "</p>";
		}	
		html_text = html_text + "</td>";
		}
		html_text = html_text + "</tr>"	;
	}
	html_text = html_text + "</TABLE>";
	document.f1.htmltext2.value = html_text;
	
	//alert(html_text);
}
//-->
</SCRIPT>
</head>
<body topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0>
<table bgcolor="#c0c0c0" border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="5" class="page_title">עדכון טקסט מבצע</td>
  </tr> 
</table>
<table border="0" width="100%"  cellspacing="1" cellpadding="0">
	<tr>
		<td align="center"colspan=2 width=100%>&nbsp;</td>
			</tr>
</table>
<table width="710" cellspacing="0" cellpadding="0" align=center border="0" bgcolor="#ffffff" >
  <tr>
    <td align="center" valign="middle" bgcolor="#3E596E" height=20 class="small_table_header" nowrap>עדכון טקסט</td>
  </tr>
<tr>
  <td bgcolor="#DDDDDD" height="10"><table><tr><td></td></tr></table></td>
</tr>
<form name="f1" ACTION="event.asp?innerparent=<%=innerparent%>&pageId=<%=pageId%>&prodId=<%=prodId%>" METHOD="post">
<tr>
	<td align="right" bgcolor="#DDDDDD">
	<table  height="20">
	<tr>
	<td>
		<table border="1"  height="20" bgcolor="white" cellspacing="1" cellpadding="2">
		<tr>
			<td bgcolor="Green">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<tr>
		</table>
	</td>
	<td class="form"><b>HTM טקסט&nbsp;</b>	
		<input type="radio" name="numcolumn" value="1" checked  LANGUAGE=javascript onclick="return numcolumn_onclick()">
	</td>
	</tr>
	</table>
	</td>
  </tr>
<%

'pertext2=pertext1
'pertext2=Replace(pertext2,"&#13;&#09;","")
'pertext2=Replace(pertext2,chr(13) & chr(10),"")
'pertext2=Replace(pertext2,"------------------------","")
'pertext2=Replace(pertext2,"<tr", "&#13;"+"שורה חדשה"+"&#13;"+"<tr")
'pertext2=Replace(pertext2,"<td","&#13;"+"<td")
'pertext2=Replace(pertext2,"<p>","<p>"+"&#13;"+"כתוב תוכן של התא"+"&#13;"+"----------"+"&#13;")
'pertext2=Replace(pertext2,"</p>","&#13;"+"----------"+"&#13;"+"</p>")
'pertext2=Replace(pertext2,"<tr>","<tr>"+"&#13;")
'pertext2=Replace(pertext2,"<tr>","&#13;"+"<tr>")
%>
<tbody id="text_html" name="text_html">
<tr>
  <td align="center" bgcolor="#DDDDDD">
	 <font face="Arial (Hebrew)"><textarea name="htmltext" rows="15" cols="90"><%=pertext1%></textarea></font>
	 
	 </td>
	 
</tr>
  <tr>
	<td bgcolor="#DDDDDD" height="10" nowrap></td>
  </tr>
</tbody>  
<tr>
	<td align="right" bgcolor="#DDDDDD">
	<table  height="20" border=0>
	<tr>
	<td nowrap colspan=11 align=right class="form"><b>בטבלה HTML טקסט</b>
	<input type="radio" name="numcolumn" value="2" LANGUAGE=javascript onclick="return numcolumn_onclick()">
	</td>  
	</tr>
	<tr id=textpict style="display:none">
	<td>
		<INPUT type="button" value="בנה טבלה" class="but" style="width:80" LANGUAGE=javascript onclick="return button1_onclick()">
	</td>
	<td width=5 nowrap></td>
	<td nowrap class="form">
		<SELECT class="app" style="width:45" id=numborder name=numborder>
		<%for i=0 to 6%>		<OPTION value=<%=i%>><%=i%></OPTION>
		<%next%>		</SELECT>
		<b>גבול</b>
	</td>
	<td width=10>
		&nbsp;
	</td>
	<td nowrap class="form">
		<SELECT class="app" style="width:60" id=numwidth name=numwidth>
		<%for i=10 to 100 step 5%>		<OPTION value=<%=i%> <%if i=100 then%> selected<%end if%>><%=i%> %</OPTION>
		<%next%>		</SELECT>
		<b>רוחב טבלה</b>
	</td>
	<td width=10>
		&nbsp;
	</td>
	<td nowrap class="form">
		<SELECT class="app" style="width:45" id=columnnumb name=columnnumb>
		<%for i=1 to 10%>		<OPTION value=<%=i%>><%=i%></OPTION>
		<%next%>		</SELECT>
		<b>כמות טורים</b>
	</td>
	<td width=10>
		&nbsp;
	</td>
	<td nowrap class="form">
		<SELECT class="app" style="width:45" id=numrow name=numrow>
		<%for i=1 to 30%>		<OPTION value=<%=i%>><%=i%></OPTION>
		<%next%>		</SELECT>
		<b>כמות שורות</b>
	</td>
	<td width=10>
		&nbsp;
	</td>
	<td nowrap align=left class="form">
		<SELECT class="app" style="width:65" id=tbl_align name=tbl_align>
			<OPTION value=right>right</OPTION>			<OPTION value=center>center</OPTION>			<OPTION value=left>left</OPTION>
		</SELECT>
		<b>ישר טבלה</b>
	</td>
	</tr>
	</table>
	</td>
  </tr> 
<tr>
	 <td bgcolor="#DDDDDD"><div id=mytable name=mytable></div>
	 <!--font face="Arial (Hebrew)"><textarea name="htmltext2" rows="15" cols="70"><%'=pertext2%></textarea></font-->
	 </td>
</tr>
</tbody>
<!--//////////////////////////////////////////////-->
	    <input type="hidden" name="elemId" value="<%=elemId%>">
	    <input type="hidden" name="pageId" value="<%=pageId%>">
	    <input type="hidden" name="homePage" value="<%if perHomePage=True then%>&quot;1&quot;<%else%>&quot;0&quot;<%end if%>">
	    <input type="hidden" name="place" value="<%=place%>">
	    <input type="hidden" name="editHTML" value="1">
	    <input type="hidden" name="htmltext2" value="">
</form>
<tr>
	<td width="100%" nowrap bgcolor="#DDDDDD" height=30 valign=middle>
	<table width="100%" border="0" cellpadding="1" cellspacing="0" bgcolor="#FFFFFF">
	  <tr>
	  <td bgcolor="#DDDDDD" width=45% nowrap align="right">
	    <INPUT class="but" style="width:90" type="button" onclick="document.location.href='event.asp?innerparent=<%=innerparent%>&pageId=<%=pageId%>&amp;prodId=<%=prodId%>&isright=<%=isRightMenu%>'" value="ביטול" id=button2 name=button2>
		</td>
		<td width=5% nowrap bgcolor="#DDDDDD"></td>					  
		<td bgcolor="#DDDDDD" width=45% nowrap align="left">
	    <INPUT class="but" style="width:90" type="button" onclick="return FormSubmitFun();" value="אישור" id=button1 name=button1>
		</td>
	  </tr>	
	</table>
	</td>
</tr>	   
</table>
</td></tr></table>
</td></tR></table>
<br>
</body>
</html>
<!--#include file="../../connect.asp"-->
<!--#include file="../../reverse.asp"-->
<!--#include file="../checkWorker.asp"-->
<html>
<head>
<!-- #include file="../../title_meta_inc.asp" -->
<meta charset="windows-1255">
<link href="../../IE4.css" rel="STYLESHEET" type="text/css">
</head>
<script LANGUAGE="JavaScript">
<!--

function CheckFields(){
	var doc=window.document;
	if (doc.forms['frmMain'].elements['title'].value==''){
		 alert("חובה למלא את השדה");
	     return false;
	}else
		 return true;
}

<!--End-->
</script>  
<% 


catID=Request.QueryString("catID")
pageID=Request.QueryString("pageID")

if request.form("referer")<>"" then 'after form filling
	 referer=sFix(request.form("referer"))
	 if trim(catID)="" then 'new record in DataBase
		sql="INSERT INTO statistic_from_banner (Page_ID,referer,isAction) VALUES ("&pageID&",'"&referer&"',0)"
        'Response.Write sql
        con.getRecordSet (sql)
        set selID=con.getRecordSet("SELECT TOP 1 category_ID FROM statistic_from_banner ORDER BY category_ID DESC")
        if not selID.EOF then
           catID=selID("category_ID")
        end if
        selID.close
        set selID=Nothing
	 else	'update existing record
        sql="UPDATE statistic_from_banner SET referer='"& referer &"' WHERE category_ID="& catID
	    'Response.Write(s)
		con.getRecordSet (sql)
     end if 
     %>
     <script language=javascript>
	<!--
		if(window.opener)
			window.opener.location.reload(true);
		self.close();	
	//-->
	</script>

     <%
end if

%>
<body style="margin:0px;background-color:#E5E5E5" onload="window.focus()">
<table border="0" width="100%" cellspacing="0" cellpadding="0" align="center" valign="top">
<tr><td class="page_title">&nbsp;<%if trim(catID)<>"" then%><font style="color:#000000;"><b>עדכון דף נחיתה</b></font><%elseIf trim(catID) = "" Then%><font style="color:#000000;"><b>הוספת דף נחיתה</b><%end if%>&nbsp;</td></tr>         	
<tr><td height="5"></td></tr>
<tr><td>
<table align=center border="0" cellpadding="2" cellspacing="1" width="100%">
<tr>
<%if trim(catID)<>"" then
  sql="select referer FROM statistic_from_banner WHERE category_ID="& catID
  set pr=con.getRecordSet(sql)
  if not pr.EOF then
	 prReferer=pr("referer")
   end if  
end if%>
<FORM name="frmMain" ACTION="addCat.asp?catID=<%=catID%>&pageID=<%=pageID%>" METHOD="post" onSubmit="return CheckFields()">
<tr>
   <td width="70%" class="card" align=right>
   <input type="text" name="referer" style="width:250" value="<%=vFix(prReferer)%>" class="texts" dir=rtl></td>
   <td width="30%" class="card" align="center" nowrap>שם מקור הגעה</td>
</tr>
<tr><td colspan="2" nowrap></td></tr>
<tr>
    <td colspan="2" width="100%">
	<table width="100%" border="0" cellpadding="5" ID="Table1">
    <tr>
    <td align="right" width=45% nowrap>
    <INPUT class="but_menu" type="button" style="width:90px" value="ביטול" onclick="window.close()"></td>
	<td width="10%" nowrap></td>
    <td align="left" width=45%>
    <input type="submit" value="עדכון" class="but_menu" style="width:90">
    </td></tr>
    </table>
   </td>
 </tr>    
<tr><td colspan="2" nowrap></td></tr>
</form>
</table>
</td></tr></table>
</body>
<%set con=Nothing%>
</html>

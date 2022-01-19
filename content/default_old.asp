<html>

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../dynamic_style.css" rel="STYLESHEET" type="text/css">
<script src="../javascript/script.js"></script>
</head>

<script>

function checkFields(objForm)
{
	if(document.form1.username.value == '')
	{
		window.alert("!!נא להכניס שם משתמש");
		document.form1.username.select();
		return false;
	}
	if(document.form1.password.value == '')
	{
		window.alert("!!נא להכניס סיסמה");
		document.form1.password.select();
		return false;
	}
	return true;
}
</script>


<%
catid=Request.QueryString("catid")
subcatid=Request.QueryString("subcatid")
if catid<>nil and subcatid<>nil then
	cid=cint(catid)-1 'Index in category arrays
	sid=cint(subcatid)-1 'Index in subcategory arrays
end if
'catid="2"
%>
<!--#include file="../include/define_categories.asp"-->
<body marginwidth="10" marginheight="0" hspace="10" vspace="0" topmargin="0"
leftmargin="10" rightmargin="10" bgcolor="white">
<FORM method="POST" action="../netCom/members/companies/companies.asp" name=form1 id=form1 onSubmit="return checkFields(this)">									

<!--include file="../PopUpMenus/java_layer.asp"-->

<!--#include file="../include/top.asp"-->
<!--
<table border="0" width="100%" height="25" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" bgcolor="#060165" align="left" height="25">
    <include file="../include/top_bar_links.asp">
    </td>
  </tr>
</table>
-->
<table border="0" width="100%" bgcolor="#6F6DA6" height="302" cellspacing="0"
cellpadding="0" background="../images/list_bgr<%=catid%>.jpg">
  <tr>
    <td width="100%" valign="top" align="right"><div align="right">
    <table border="0"
    width="75%" cellspacing="0" cellpadding="0">
        <tr>
        <td width="100%" valign="top" align="right" height="10"></td>
      </tr>
     <tr>
        <td width="100%"  height="23" align="center" valign="bottom">
        <table border="0" width="90%" cellspacing="0" cellpadding="0">
         <tr><td  valign="bottom" align="right" height="23" class="title_page" dir="rtl">
        	<%'Response.Write cid&" "&sid
        	'Response.End%>
        	<%if catid<>nil and subcatid<>nil then%>
        		<%=subcattitle(cid,sid)%>
        	<%end if%>
        	</td>
			</tr>
		  </table>
		</td>
      </tr>
      <tr>
        <td width="100%" valign="top" align="right" height="16"></td>
      </tr>
      <tr>
        <td width="100%" height="214" valign="top" align="center">
        <table border="0" width="90%"
        cellspacing="0" cellpadding="0">
          <tr>
            <td height="7"><img  border="0" src="../images/pina2.gif" width="7" height="7"></td>
            <td bgcolor="#FFFFFF" width="100%" height="7"></td>
            <td height="7"><img  border="0" src="../images/pina1.gif" width="7" height="7"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF"></td>
            <td bgcolor="#FFFFFF" width="100%" align="center">
            <table border="0" width="96%" cellspacing="0" cellpadding="0">
				<tr>
					<td align="left">
					
					<table border="0" width="100%" cellspacing="0" cellpadding="0">
					<%if catid<>nil and subcatid<>nil then%>
						<tr><td height="5"></td></tr>
						<tr>
							<td align="right" dir="rtl" class="title">
								<%=subtitle(cid,sid)%>
							</td>
						</tr>
						<tr><td height="5"></td></tr>
						<tr>
							<td align="right" dir="rtl">
							<b><%=textdescr(cid,sid)%></b>
							</td>
						</tr>
						<tr><td height="5"></td></tr>
						<tr>
							<td align="right" dir="rtl" class="marked">
המערכת מאפשרת				
							</td>
						</tr>
						<tr>
							<td align="left" width="100%">
							<table width="95%" border="0" cellpadding="0" cellspacing="2">
							<%for i=0 to ubound(listrows,3)
								if listrows(cid,sid,i)<>nil then%>
								<tr>
									<td width="100%" align="right" dir="rtl">
										<%=listrows(cid,sid,i)%>
									</td>
									<td>
										<img src="../images/h_bul.gif" border="0" hspace="2">
									</td>
								</tr>
								<%end if%>
							<%next%>
							</table>
							</td>
						</tr>
						<%end if%>

					</table>
					</td>
				</tr>
       <tr><td height="5"></td></tr>
           </table>
            </td>
            <td bgcolor="#FFFFFF"></td>
          </tr>
          <tr>
            <td height="7"><img  border="0" src="../images/pina3.gif" width="7" height="7"></td>
            <td bgcolor="#FFFFFF" width="100%" height="7"></td>
            <td height="7"><img  border="0" src="../images/pina4.gif" width="7" height="7"></td>
          </tr>

        </table>
        </td>
      </tr>
      <tr><td height="10"></td></tr>
    </table>
    </div></td>
  </tr>
</table>

<table border="0" width="100%" height="15" cellspacing="0" cellpadding="0">
  <tr>
    <td width="100%" height="2"></td>
  </tr>
  <tr>
    <td width="100%" bgcolor="#060165" height="13"></td>
  </tr>
</table>

<table border="0" width="100%" height="50" cellspacing="0" cellpadding="0"
background="images/list_bgr_bot.jpg">
  <tr>
    <td><img  border="0" src="../images/list_pic_bot.jpg" width="235" height="50"></td>
    <td valign="bottom" align="right"><img  border="0" src="../images/shuet.gif" width="435" height="23"
    border="0"></td>
  </tr>
</table>
</form>
</body>
</html>

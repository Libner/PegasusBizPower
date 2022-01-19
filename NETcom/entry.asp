<html>
<!--#include file="connect.asp"-->
<!--#include file="reverse.asp"-->

<%

if Request.Form("username")<>nil and Request.Form("password")<>nil and Request.Form("username")<>"" and Request.Form("password")<>"" then
session("ORGANIZATION_NAME") = ""

%>
<!--#include file="members/checkWorker.asp"-->

<%
  Response.Redirect "members/companies/companies.asp"
  end if
%>
<head>
<title>BizPower</title>
<link href="IE4.css" rel="STYLESHEET" type="text/css">
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<script src="../javascript/script.js"></script>
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
</head>

<body marginwidth="0" marginheight="0" hspace="0" vspace="0" topmargin="0" leftmargin="0" rightmargin="0"  bgcolor="#E5E5E5">
<table border="0" width="100%" height="83" cellspacing="0" cellpadding="0" border=0 ID="Table1">
  <tr>
    <td valign="bottom"><img  border="0"src="images/top_logo.gif" width="223" height="62" border="0"></td>
    <td align="right"><div align="right"><table border="0" width="321" height="87"
    cellspacing="0" cellpadding="0" ID="Table2">
      <tr>
        <td width="100%" valign="bottom" height="25"><table border="0" cellspacing="0"
        cellpadding="0" ID="Table3">
          <tr>
            <td width="27"><a href="" onclick="return false" class="home_link"><img border="0" src="images/top_link4.gif" width="27" height="19" border="0" name="toplink4" id="toplink4"></td><!-- onMouseOver="changeImage(this,'../images/top_link4_a.gif');" onMouseOut="changeImage(this,'../images/top_link4.gif')"></a></td-->
            <td width="7"><img border="0" src="images/top_link_bul.gif" width="7" height="19"></td>
            <td width="47"><a href="contact_us/default.asp" onclick="return false"><img border="0" src="images/top_link3.gif" width="47" height="19" border="0" name="toplink3" id="toplink3"></td><!-- onMouseOver="changeImage(this,'../images/top_link3_a.gif');" onMouseOut="changeImage(this,'../images/top_link3.gif')"></a></td-->
            <td width="7"><img border="0" src="images/top_link_bul.gif" width="7" height="19"></td>
            <td width="73"><a href="" onclick="return false" ><img border="0" src="images/top_link2.gif" width="73" height="19" border="0" name="toplink2" id="toplink2"></td><!--onMouseOver="changeImage(this,'../images/top_link2_a.gif');" onMouseOut="changeImage(this,'../images/top_link2.gif')"></a></td-->
            <td width="7"><img border="0" src="images/top_link_bul.gif" width="7" height="19"></td>
            <td width="45"><a href="" onclick="return false" class="home_link"><img border="0" src="images/top_link1.gif" width="45" height="19" border="0" name="toplink1" id="toplink1"></td><!-- onMouseOver="changeImage(this,'../images/top_link1_a.gif');" onMouseOut="changeImage(this,'../images/top_link1.gif')"></a></td-->
          </tr>
        </table>
        </td>
      </tr>
      <tr>
        <td width="100%" align="right" bgcolor="#FFD012" height="1"></td>
      </tr>
      <tr>
        <td width="100%" valign="bottom" height="57"><img  border="0"src="images/home_slogan.gif" width="261"
        height="37"></td>
      </tr>
    </table>
    </div></td>
  </tr>
</table>
<div align="center"><center>
<table border="0" width="100%" height="18" cellspacing="0" cellpadding="0" bgcolor="#CECECE" ID="Table4">
  <tr>
    <td valign="bottom" width="318" nowrap><img src="images/pict_top_netcom.jpg" width="317" height="9"></td>
    <td align="right" width="100%"></td>
  </tr>
</table>
</center></div><div align="center"><center>

<table border="0" width="100%" bgcolor="#FFFFFF" height="300" cellspacing="0" cellpadding="0" background="images/bgr_list.jpg" ID="Table5">
  <tr>
    <td width="118" align=right nowrap valign="top"><img src="images/pict1_netcom.jpg" width="118" height="301"></td>
    <td valign="top" align=right nowrap width="200"><img src="images/pict2_netcom.jpg" width="200" height="301"></td>
    <td width="100%" valign="top" align="center">
     <table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table6">
     <!-- MAIN -->         
      <FORM method="POST" action="default.asp" name=form1 id=form1 onSubmit="return checkFields(this)">									
       <div align="right">
		<tr><td align=center>
		<table border="0" width="294" height="200" cellspacing="0"  cellpadding="0" ID="Table7">
		  <tr>
		    <td width="100%" valign="bottom" height="113"><img src="images/title_knisa.gif"
		    width="251" height="53" alt="title_knisa.gif (1297 bytes)"></td>
		  </tr>
		  <tr>
		    <td width="100%" height="70"><table border="0" width="294" cellspacing="0" cellpadding="0"
		    height="70" ID="Table8">
		      <tr>
		        <td width="13"><img src="images/login_left.gif" width="13" height="70"></td>
		        <td bgcolor="#E5E5E5" width="176" align="center"><table border="0" width="140"
		        cellspacing="0" cellpadding="0" height="50" ID="Table9">
		          <tr>
		            <td><input type="text"  name="username" id="username" size="12" tabindex="1"
		            style="font-size: 9pt; font-family: Arial; border: 1px solid rgb(128,128,128)"></td>
		            <td align="right"><img src="images/user.gif" width="42" height="7"></td>
		          </tr>
		          <tr>
		            <td><input type="password" name="password" size="12" tabindex="2"
		            style="font-size: 9pt; font-family: Arial; border: 1px solid rgb(128,128,128)" ID="Password1"></td>
		            <td align="right"><img src="images/pass.gif" width="32" height="7"></td>
		          </tr>
		        </table>
		        </td>
		        <td width="105"><img src="images/login_right.gif" width="90" height="70"></td>
		      </tr>
		    </table>
		    </td>
		  </tr>
		  <tr>
		    <td width="100%" height="37"><table border="0" cellspacing="0" cellpadding="0" ID="Table10">
		      <tr>
		        <td><INPUT type="image" src="images/knisa.gif" width="133" height="37" border="0" ID="Image1" NAME="Image1"></td>
		        <td><img src="images/knisa_right.gif" width="92" height="37"></td>
		      </tr>
		    </table>
		    </td>
		  </tr>
		</table>
			</td>
		</tr>	
		</div>	
		</form>
<!-- END MAIN -->        
     </table>
    </td>
    <td width="17" valign="top" align="right"></td>
  </tr>
</table>
</center></div>
</body>
</html>
<%set con=Nothing%>
<!-- #include file="connect.asp" -->
<!-- #include file="reverse.asp" -->
<!-- #include file="members/checkWorker.asp" -->
<html>
<head>
<title>Bizpower</title>
<meta http-equiv=Content-Type content="text/html; charset=windows-1255">
<link href="../dynamic_style.css" rel="STYLESHEET" type="text/css">

<script src="javascript/script.js"></script>
</head>
<body marginwidth="10" marginheight="0" hspace="10" vspace="0" topmargin="0"
leftmargin="10" rightmargin="10" bgcolor="white">
  <table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table1">
    <tr>
      <td width="100%" valign="bottom" height="79">
      <table border="0" width="100%" height="66"  cellspacing="0" cellpadding="0" ID="Table2">
        <tr>
          <td valign="bottom">
          <a href="../default.asp" target=_self>
          <img src="../images/top_logo.gif" width="223" height="59" border="0"></a></td>
          <td valign="bottom" align="right" width="100%"><img src="../images/top_slogan.gif"
          width="282" height="50"></td>
          <td bgcolor="#6F6DA6">
          <table border="0" width="201" height="66" cellspacing="0" cellpadding="0" ID="Table3">
            <tr>
              <td width="100%" height="21"><img src="../images/top_knisa.gif" width="201" height="21"
              alt="top_knisa.gif (730 bytes)"></td>
            </tr>
            <tr>
              <td width="100%" height="45">
              <table border="0" width="201" height="45" cellspacing="1" cellpadding="0" ID="Table4">
                <tr>                 
                  <td align="right" style="color:#FFD011;font-weight:600">&nbsp;<%=user_name%>&nbsp;</td>
                  <td width="67" align=right style="color:#FFFFFF;font-weight:600">&nbsp;משתמש&nbsp;</td>
                  <td width=15 nowrap></td>
                </tr>
                <tr>
                 <td align="right" style="color:#FFD011;font-weight:600">&nbsp;<%=org_name%>&nbsp;</td>
                 <td width="67" align=right style="color:#FFFFFF;font-weight:600">&nbsp;חברה&nbsp;</td>                                 
                 <td width=15 nowrap></td>
                </tr>
              </table>
              </td>
            </tr>
          </table>
          </td>
        </tr>
      </table>
      </td>
    </tr>    
  </table>
<!--#include file="../include/define_links_def.asp"-->
<!--#include file="../PopUpMenus/netcom_layer_def.asp"-->
<table cellpadding=0 cellspacing=0 width=100% bgcolor="#060165" ID="Table18">
<tr><td align=right>
<table border="0" bgcolor="#060165" cellspacing="1" cellpadding="1" align="right" dir=rtl ID="Table19">
<tr>
      <td width="100%" bgcolor="#060165" height="18" valign="top" >
       <table border="0" width="100%"  cellspacing="0" cellpadding="0" valign="top" ID="Table20">
        <tr>
        <%	for i=0 to ubound(catname) %>
			<td  align="center" nowrap valign="top" >			
		  <%if catname(i)<>nil then%>
		<a class="top_link"  href="" onclick="return false;" onMouseOver="popUp('elMenu<%=i+1%>',event);status='';return true" onMouseOut="popDown('elMenu<%=i+1%>');status=''">&nbsp;&nbsp;<%=catname(i)%>&nbsp;&nbsp;</a><img src="../images/h_bul.gif" border="0">
		<%end if %>
		</td>
		<%next%>
	  </tr></table></td></tr>	
 </table></td></tr>
  <tr>
      <td width="100%" height="2" bgcolor="#ffffff"></td>
    </tr>
 </table>    
<table border="0" width="100%" height="2" cellspacing="0" cellpadding="0" ID="Table17">
  <tr>
    <td width="100%" height="2"></td>
  </tr>  
</table>

<table border="0" width="100%" bgcolor="#6F6DA6" height="302" cellspacing="0"
cellpadding="0" background="../images/home_bgr.jpg" ID="Table5">
<tr><td align=right>
<table border="0" width="100%" bgcolor="#6F6DA6" height="302" cellspacing="0"
cellpadding="0" background="../images/list_bgr3.jpg" ID="Table6">
  <tr>
    <td width="100%" valign="top" align="right"><div align="right">
    <table border="0"
    width="76%" cellspacing="0" cellpadding="0" ID="Table7">
        <tr>
        <td width="100%" valign="top" align="right" height="10"></td>
      </tr>
     <tr>
        <td width="100%"  height="23" align="center" valign="bottom">
        <table border="0" width="90%" cellspacing="0" cellpadding="0" ID="Table8">
         <tr><td  valign="bottom" align="right" height="23" class="title_page" dir="rtl">
        		כלכלה ופיננסים
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
        cellspacing="0" cellpadding="0" ID="Table9">
          <tr>
            <td height="7"><img  border="0" src="../images/pina2.gif" width="7" height="7"></td>
            <td bgcolor="#FFFFFF" width="100%" height="7"></td>
            <td height="7"><img  border="0" src="../images/pina1.gif" width="7" height="7"></td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF"></td>
            <td bgcolor="#FFFFFF" width="100%" align="center">
            <table border="0" width="96%" cellspacing="0" cellpadding="0" ID="Table10">
				<tr>
					<td align="left">
					
					<table border="0" width="100%" cellspacing="0" cellpadding="0" ID="Table11">
					
						<tr><td height="5"></td></tr>
						<tr>
							<td align="right" dir="rtl">
							<span  class="subtitle">כדי להכין הצעת מחיר / תמחור לפרויקט עתידי,  
							</span>
							
								אליך לבצע את השלבים הבאים:
							</td>
						</tr>
						<tr><td height="5"></td></tr>
						<tr>
							<td align="left" width="100%">
							<table width="95%" border="0" cellpadding="0" cellspacing="2" ID="Table12">
							
		<tr><td width="100%" align="right" dir="rtl"><b>
		א.	הזן את פרטי העובדים שלך
		</b></td></tr>
		<tr><td width="100%" align="right" dir="rtl" style="padding-right:20px;">
		1.	<a href="members/workers/jobs.asp" style="font-size:10pt;text-decoration:underline" class="home_link" target=_self>הגדר את התפקידים ועלויות</a> 
		</td></tr>
		<tr><td width="100%" align="right" dir="rtl" style="padding-right:20px">
		2.	<a href="members/workers/default.asp" style="font-size:10pt;text-decoration:underline" class="home_link" target=_self>הכנס את פרטי העובדים</a>
		</td></tr>
		<tr><td height="5"></td></tr>
		<tr><td width="100%" align="right" dir="rtl"><b>
		ב.	הזן את רשימת הארגונים שלך
		</b></td></tr>
		<tr><td width="100%" align="right" dir="rtl" style="padding-right:20px">
		1.	<a href="members/companies/company_types.asp" style="font-size:10pt;text-decoration:underline" class="home_link" target=_self>הגדר  קבוצות ע"פ מגזרים</a>
		</td></tr>
		<tr><td width="100%" align="right" dir="rtl" style="padding-right:20px">
		2.	<a href="members/companies/companies.asp" style="font-size:10pt;text-decoration:underline" class="home_link" target=_self>הזן פרטי הארגונים</a>
		</td></tr>
		<tr><td width="100%" align="right" dir="rtl" style="padding-right:20px">
		3.	<a href="members/companies/contacts.asp" style="font-size:10pt;text-decoration:underline" class="home_link" target=_self>הוסף אנשי קשר לארגון</a>
		</td></tr>
		<tr><td height="5"></td></tr>
		<tr><td width="100%" align="right" dir="rtl"><b>
		ג.	הזן את הפרויקטים / הפרויקטים  הפעילים / עתידיים בחברה שלך.
		</b></td></tr>
		<tr><td width="100%" align="right" dir="rtl" style="padding-right:20px">
		1.	<a href="members/projects/default.asp" style="font-size:10pt;text-decoration:underline" class="home_link" target=_self>הזן את הפרויקט תוך כדי הגדרתה בתור עתידית או בביצוע</a>
		</td></tr>
		
		<tr><td height="5"></td></tr>
		<tr><td width="100%" align="right" dir="rtl"><b>
		ד.	תכין את טבלת ההמחרה ההשוואתית
		</b></td></tr>
		<tr><td width="100%" align="right" dir="rtl" style="padding-right:20px">
		1.	<a href="members/projects/prices.asp" style="font-size:10pt;text-decoration:underline" class="home_link" target=_self>הזן את פרטי ההמחרה של הפרויקט העתידי שלך והשווה אותו לפרויקטים <br>&nbsp;&nbsp;&nbsp;&nbsp;קודמים דומים הקיימים במערכת</a>
		</td></tr>
		<tr><td width="100%" align="right" dir="rtl" style="padding-right:20px">
		2.	<a href="members/projects/addPricing.asp" style="font-size:10pt;text-decoration:underline" class="home_link" target=_self>הערך  את מספר שעות העבודה  בכל תפקיד בהשוואה לפרויקטים קודמים</a>
		</td></tr>
		<tr><td width="100%" align="right" dir="rtl" style="padding-right:20px">
		3.	<a href="members/projects/addPricing.asp" style="font-size:10pt;text-decoration:underline" class="home_link" target=_self>קבע את  אחוז הרווח הרצוי וקבל תוצאה לצורך הגשת הצעת המחיר</a>
		</td></tr>
		
		<tr><td height="10"></td></tr>
		</table></td></tr>
		
	</table></td></tr>						
    <tr><td height="5"></td></tr>
    </table></td>
    <td bgcolor="#FFFFFF"></td>
    </tr>
    <tr>
        <td height="7"><img  border="0" src="../images/pina3.gif" width="7" height="7"></td>
        <td bgcolor="#FFFFFF" width="100%" height="7"></td>
        <td height="7"><img  border="0" src="../images/pina4.gif" width="7" height="7"></td>
    </tr>
    </table></td></tr>
    <tr><td height="10"></td></tr>
    </table>
</td></tr></table>
</td></tr></table>
<table border="0" width="100%" height="15" cellspacing="0" cellpadding="0" ID="Table13">
  <tr>
    <td width="100%" height="2"></td>
  </tr>
  <tr>
    <td width="100%" bgcolor="#060165" height="13"></td>
  </tr>
</table>

<table border="0" width="100%" height="50" cellspacing="0" cellpadding="0"
background="images/list_bgr_bot.jpg" ID="Table14">
  <tr>
    <td><img  border="0" src="../images/list_pic_bot.jpg" width="235" height="50"></td>
    <td valign="bottom" align="right"><img  border="0" src="../images/shuet.gif" width="435" height="23"
    border="0"></td>
  </tr>
</table>
</body>
</html>

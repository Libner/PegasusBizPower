<!--#include file="..\netcom/reverse.asp"-->
<!--#include file="..\netcom/connect.asp"-->
<!--#include file="checkWorker.asp"-->

<%
session("siteid") = 2 '//for tavniot tochen
%>
<HEAD>
<meta charset="windows-1255">
<TITLE>Bizpower Administration</TITLE>
<link rel="stylesheet" type="text/css" href="../admin.css">
</HEAD>
<BODY>
<center>

<table width="60%" cellspacing="0" cellpadding="0" >
<%'//start%>
           <tr><td width="100%" height="30"></td></tr>          
           <tr><td height="5" align="center" class="page_title">Bizpower ניהול אתר</td></tr>
           <tr><td width="100%" height="1" bgcolor="#C0C0C0"></td></tr>          
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>           
           <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="choose.asp">Hebrew site administration</a></strong></td>
           </tr>                                                                                                                                                                                                        
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="2" bgcolor="#000000"></td></tr>
          <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>           
          <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="organizations/default.asp">Organizations</a></strong></td>
           </tr>                                                                                                                                                                                                        
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="2" bgcolor="#000000"></td></tr>
           
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>
           <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="templates/default.asp">Sales pages</a></strong></td>
           </tr>
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="2" bgcolor="#000000"></td></tr>
           
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>
           <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="pages_e/admin.asp?maincat=2">Dynamic pages</a></strong></td>
           </tr>
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="2" bgcolor="#000000"></td></tr>
                  
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>
           <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="news/admin.asp?catID=1">News and events</a></strong></td>
           </tr>
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="2" bgcolor="#000000"></td></tr>
                                 
           <tr><td width="100%" height=50 nowrap></td></tr>
           <tr>
			<td width="100%">
				<table border="0" width=100% height="50" cellspacing="0" cellpadding="0">
				<tr>
						<td align="right" width="48%" class="12normalB"><b>Built by</b></td> 
						<td align="left" width="2%" valign="bottom">&nbsp;</td> 
						<td width="50%" align="left"><img src="images/cyber.gif"  style="vertical-align:10%" hspace=5 border=0></td> 
				</tr>
				</table>					
			</td>
		</tr>
<%'//end%>
</table>
</div>
</center>
</BODY>
</HTML>

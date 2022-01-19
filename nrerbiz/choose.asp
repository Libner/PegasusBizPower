<!--#include file="..\netcom/reverse.asp"-->
<!--#include file="..\netcom/connect.asp"-->
<!--#include file="checkWorker.asp"-->
<%
	'//for tavniot tochen
	if request.QueryString("siteid")<>nil then ' change language
		session("siteid") = request.QueryString("siteid")
	else
		if session("siteid")=nil or session("siteid")="" then 'default - hebrew
			session("siteid") = 1 
		end if
	end if
%>
<HEAD>
<meta charset="windows-1255">
<TITLE>Bizpower Administration</TITLE>
<link rel="stylesheet" type="text/css" href="../admin.css">
</HEAD>
<BODY>
<center>
<%if session("siteid") = 1 then 'hebrew menu%>
<table width="60%" cellspacing="0" cellpadding="0" >
<%'//start%>
           <tr><td width="100%" height="30"></td></tr>          
           <tr><td height="5" align="center" class="page_title">Bizpower ניהול אתר</td></tr>
           <tr><td width="100%" height="1" bgcolor="#C0C0C0"></td></tr>          
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>           
           <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="choose.asp?siteid=2">English site administration</a></strong></td>
           </tr>                                                                                                                                                                                                        
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="2" bgcolor="#000000"></td></tr>
          <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>           
          <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="organizations/default.asp">ארגונים</a></strong></td>
           </tr>                                                                                                                                                                                                        
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="2" bgcolor="#000000"></td></tr>
           
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>
           <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="templates/default.asp">תבניות דפי מבצע</a></strong></td>
           </tr>
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="2" bgcolor="#000000"></td></tr>
           
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>
           <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="pages/admin.asp?maincat=1">תכני אתר תדמיתי</a></strong></td>
           </tr>
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="2" bgcolor="#000000"></td></tr>
                  
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>
           <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="news/admin.asp?catID=1">חדשות ואירועים</a></strong></td>
           </tr>
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="2" bgcolor="#000000"></td></tr>
          
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>
           <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="mails/default.asp">הפצות</a></strong></td>
           </tr>
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="2" bgcolor="#000000"></td></tr>
           
          
           <tr><td width="100%" height=50 nowrap></td></tr>
           <tr>
			<td width="100%">
				<table border="0" width=100% height="50" cellspacing="0" cellpadding="0">
				<tr>
						<td width="50%" align="right"><img src="images/cyber.gif"  style="vertical-align:10%" hspace=5 border=0></td> 
						<td align="left" width="2%" valign="bottom">&nbsp;</td> 
						<td align="left" width="48%" class="12normalB"><b>נבנה ע"י</b></td> 
				</tr>
				</table>					
			</td>
		</tr>		
<%'//end%>
</table>
<%else 'English menu%>
<table width="60%" cellspacing="0" cellpadding="0" ID="Table1">
<%'//start%>
           <tr><td width="100%" height="30"></td></tr>          
           <tr><td height="5" align="center" class="page_title">Bizpower ניהול אתר</td></tr>
           <tr><td width="100%" height="1" bgcolor="#C0C0C0"></td></tr>          
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>           
           <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="choose.asp?siteid=1">Hebrew site administration</a></strong></td>
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
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="pages/admin.asp?maincat=2">Dynamic pages</a></strong></td>
           </tr>
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="2" bgcolor="#000000"></td></tr>
                  
           <tr><td width="100%" height="1" bgcolor="#808080"></td></tr>
           <tr><td width="100%" height="1" bgcolor="#FFFFFF"></td></tr>
           <tr height="22">
              <td bgcolor="#C0C0C0" align="center" valign="middle" nowrap><strong><a class="admin" href="news/admin.asp?catID=2">News and events</a></strong></td>
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
<%end if%>
</div>
</center>
</BODY>
</HTML>

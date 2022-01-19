<!--#include file="../include/define_links.asp"-->
<!--#include file="../PopUpMenus/netcom_layer.asp"-->
<table cellpadding=0 cellspacing=0 width=100% bgcolor="#060165">
<tr><td align=right>
<table border="0" bgcolor="#060165" cellspacing="1" cellpadding="1" align="right" dir=rtl ID="Table3">
<tr>
      <td width="100%" bgcolor="#060165" height="18" valign="top" >
       <table border="0" width="100%"  cellspacing="0" cellpadding="0" valign="top" ID="Table1">
        <tr>
        <%for i=0 to ubound(catname)%>
			<td  align="center" nowrap valign="top" >			
		  <%if catname(i)<>nil then%>		 
		<a class="top_link" href="#" onclick="return false;" onMouseOver="popUp('elMenu<%=i+1%>',event);status='';return true" onMouseOut="popDown('elMenu<%=i+1%>');status=''">&nbsp;&nbsp;<%=catname(i)%>&nbsp;&nbsp;</a><img src="../../images/h_bul.gif" border="0">
		<%end if%>
		</td>
		<%next%>
	  </tr></table></td></tr>	
 </table></td></tr>
  <tr>
      <td width="100%" height="2" bgcolor="#ffffff"></td>
    </tr>
 </table>    
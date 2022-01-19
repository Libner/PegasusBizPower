<table border="0" width="100%" cellspacing="0" cellpadding="0">
    <tr>
      <td width="100%" valign="bottom" height="79">
      <table border="0" width="100%" height="66"  cellspacing="0" cellpadding="0" ID="Table2">
        <tr>
          <td valign="bottom">
          <a href="<%=strLocal%>/netCom/" target=_self>
          <img src="../../../images/top_logo.gif" width="223" height="59" border="0"></a></td>
          <td valign="bottom" align="center" width="100%" height=60>
          <%If perSize > 0 Then%>
          <img src="<%=strLocal%>/netcom/GetImage.asp?DB=ORGANIZATION&amp;FIELD=ORGANIZATION_LOGO&amp;ID=<%=OrgId%>">
          <%End If%>
          </td>
          <td valign="bottom" align="right" ><img src="../../../images/topslog.gif"></td>
          <td bgcolor="#6F6DA6" style="background-image:URL('../../../images/exit.gif');background-position:'top left';background-repeat:'no-repeat';">
          <table border="0" width="201" height="66" cellspacing="0" cellpadding="0">
            <tr>
              <!--td><a href="<%=strLocal%>" target=_self><img src="../../../images/exit.gif"></a></td-->
              <td width="100%" height="21"><a href="<%=strLocal%>" target=_self>             
              <img src="../../../images/top_knisa.gif" width="201" height="21" border=0></a></td>
            </tr>
            <tr>            
              <td width="100%" height="45">
              <table border="0" width="201" height="45" cellspacing="1" cellpadding="0" ID="Table4">
                <tr>                                   
                  <td align="right" style="color:#FFD011;font-weight:600">&nbsp;<%=user_name%>&nbsp;</td>
                  <td width="60" align=right style="color:#FFFFFF;font-weight:600">&nbsp;משתמש&nbsp;</td>
                  <td width=15 nowrap></td>
                </tr>
                <tr>
                 <td align="right" style="color:#FFD011;font-weight:600">&nbsp;<%=org_name%>&nbsp;</td>
                 <td width="60" align=right style="color:#FFFFFF;font-weight:600">&nbsp;חברה&nbsp;</td>                                 
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
    <tr>
      <td width="100%" bgcolor="#060165" height="18" valign="top" >
       <table border="0" width="100%" height="15" cellspacing="0" cellpadding="0" ID="Table3">
			<tr>
				<td width="100%" height="2"></td>
			</tr>
			<tr>
			<td width="100%" bgcolor="#060165" height="13" dir=rtl>
			<%''/strat of news%>
			<marquee direction=right  scrolldelay=120>
		
				<%
				sqlStr="SELECT new_ID,New_Title,New_Date FROM News"
				sqlStr=sqlStr& " WHERE New_Site_Visible=1 " 
				sqlStr=sqlStr& " ORDER BY New_Date DESC,New_ID DESC" 
				set newsList=con.getRecordSet(sqlStr)
				do while not newsList.EOF 
				newID=newsList("New_ID")
				newTitle=newsList("New_Title")
				newDate=newsList("New_Date")
				%>
				&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
				<a class="homeNews" href="<%=strLocal%>netcom/news.asp?ID=<%=newID%>">>> <%=newTitle%></a>
				
				<%
				newsList.MoveNext
				loop
				newsList.close
				set newsList=Nothing
				%>											
			</marquee>		
		<%'//end of news%>
		</td></tr></table>
      </td>
    </tr>   
  </table>
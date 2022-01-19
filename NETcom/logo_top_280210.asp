<%	Dim img_, bgr_pos, user_t, comp_t, show_logo, align_logo
	If trim(lang_id) = "2" Then
		  img_ = "_eng" : bgr_pos = "top right" : user_t = "User" : comp_t = "Company"
	Else
		  img_ = "" : bgr_pos = "top left" : user_t = "משתמש" : comp_t = "חברה"
	End If
   
	show_logo = true : align_logo = "center"
	If Not IsNull(Request.Cookies("bizpegasus")("UseBizLogo")) Then
		If trim(Request.Cookies("bizpegasus")("UseBizLogo")) = "0" Then 
			show_logo = false : align_logo = "left"
		End If	
	End If
%>    
<style>
TABLE.top
{
    background-repeat:no-repeat;
    background-position:top left;
}
</style>         
<table border="0" width="100%" cellspacing="0" cellpadding="0" dir=ltr>
    <tr>
      <td width="100%" valign="bottom" height="79">
      <table border="0" width="100%" height="66"  cellspacing="0" cellpadding="0">
        <tr>
          <td valign="bottom"><%If show_logo Then%><a href="<%=Application("VirDir")%>/netCom/" target=_self>
          <img src="<%=Application("VirDir")%>/images/top_logo.gif" width="223" height="59" border="0"></a><%End If%>
          </td>
          <td valign="bottom" align="<%=align_logo%>" width="100%" height=60>
          <%If perSize > 0 Then%>
          <img src="<%=Application("VirDir")%>/netcom/GetImage.asp?DB=ORGANIZATION&amp;FIELD=ORGANIZATION_LOGO&amp;ID=<%=OrgId%>">
          <%End If%>
          </td>
          <td valign="bottom" align="<%=align_var%>"><img src="<%=Application("VirDir")%>/images/topslog.gif"></td>
          <td width="201" valign=bottom>
          <table bgcolor="#6F6DA6" border="0" width="201" cellspacing="0" cellpadding="0" style="background-image:URL('<%=Application("VirDir")%>/images/exit<%=img_%>.gif');" class="top">
            <tr>
              <td width="201" nowrap height="21"><a href="<%=Application("VirDir")%>/exit.asp" target=_self>             
              <img src="<%=Application("VirDir")%>/images/top_knisa<%=img_%>.gif" border=0></a>             
              </td>
            </tr>
            <tr>            
              <td width="100%" height="45">
              <table border="0" width="201" height="45" cellspacing="0" cellpadding="0" dir="<%=dir_var%>">
                <tr>                                   
                  <td width=123 nowrap align="<%=align_var%>" dir="<%=dir_var%>" style="color:#FFD011;font-weight:600;vertical-align:middle;padding-<%=align_var%>:10px"><%=trim(user_name)%></td>
                  <td width=74 nowrap align="<%=align_var%>" dir="<%=dir_var%>" style="color:#FFFFFF;font-weight:600;vertical-align:middle;padding-<%=align_var%>:10px"><%=user_t%></td>
                </tr>
                <tr>
                 <td width=123 nowrap valign=top align="<%=align_var%>" dir="<%=dir_var%>" style="color:#FFD011;font-weight:600;vertical-align:middle;padding-<%=align_var%>:10px"><%=trim(org_name)%></td>
                 <td width=74 nowrap align="<%=align_var%>" dir="<%=dir_var%>" style="color:#FFFFFF;font-weight:600;vertical-align:middle;padding-<%=align_var%>:10px"><%=comp_t%></td>                                 
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
       <table border="0" width="100%" height="18" cellspacing="0" cellpadding="0">
			<tr>
				<td width="100%" height="1" nowrap></td>
			</tr>
			<tr>
			<td bgcolor="#060165" height="18" dir="<%=dir_obj_var%>">
			<%''/strat of news%>
			<marquee direction="<%=align_var%>"  scrolldelay=120>
		
				<%
				sqlStr="SELECT new_ID,New_Title,New_Date FROM News"
				sqlStr=sqlStr& " WHERE New_Site_Visible=1 " 
				If trim(lang_id) = "1" Then
					sqlStr=sqlStr& " AND Category_Id=1 " 
				Else
					sqlStr=sqlStr& " AND Category_Id=2 " 
				End If
				sqlStr=sqlStr& " ORDER BY New_Date DESC,New_ID DESC"				
				set newsList=con.getRecordSet(sqlStr)
				do while not newsList.EOF 
				newID=newsList("New_ID")
				newTitle=newsList("New_Title")
				newDate=newsList("New_Date")
				%>
				&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
				<a class="homeNews" href="<%=Application("VirDir")%>netcom/news.asp?ID=<%=newID%>">>> <%=newTitle%></a>
				
				<%
				newsList.MoveNext
				loop
				newsList.close
				set newsList=Nothing
				%>											
			</marquee>		
		<%'//end of news%>
		</td>
		</tr>
		<tr>
		<td width="100%" height="1" nowrap></td>
		</tr> 
		</table>
      </td>      
    </tr>      
  </table>
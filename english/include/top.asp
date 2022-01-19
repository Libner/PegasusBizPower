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
<%maincat=2%>
<!--#include file="../PopUpMenus/java_layer.asp"-->
  
  <table border="0" width="780" cellspacing="0" cellpadding="0" align=center>
    <tr>
      <td width="100%" valign="bottom">
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <td valign="bottom"><a href="<%=Application("VirDir")%>/english/default.asp"><img src="<%=Application("VirDir")%>/english/images/top_logo.gif" width="223" height="59" border="0"></a></td>
          <td width="100%"></td>
          <td width="320"></td>
          <td valign="bottom" align="left"></td>
          <td><table border="0" width="201" cellspacing="0" cellpadding="0">
            <tr>
              <td width="100%" height="10"></td>
            </tr>
            <tr>
              <td width="100%">
              <table border="0" width="201" cellspacing="0" cellpadding="0" dir="rtl">
				<FORM method="POST" action="<%=Application("VirDir")%>/netCom/default.asp" name=form1 id=form1 onSubmit="return checkFields(this)">									
				<input type=hidden name="wizard_id" id="wizard_id" value="<%=Request("wizard_id")%>">
                <tr>
                  <td align="left" colspan=2>
                  <input type="text" dir="ltr" name="username" id="username" class="form_text_home" >
                  </td>
                  <td width="67" class="lhome" align="left"><b>User name</b></td>
                 </tr>
                 <tr> 
                  <td valign="bottom" width="43">
                  <input type="image" src="<%=Application("VirDir")%>/english/images/top_slah.gif" width="43" height="25" border="0" id=image1 name=image1 tabindex="-1"></td>
                  </td>
                  <td align="left">
                  <input type="password" dir="ltr" name="password" id="password" class="form_text_home">
                  </td>
                  <td width="67" class="lhome" align="left"><b>Password</b></td>
                </tr>
				</form>                
              </table>
              </td></tr>
          </table></td>               
    </tr></table></td></tr>
    <tr>
       <td width="100%" style="border: 1px solid #999999; border-bottom: none" bgcolor="#DDDDDD" height="18" valign="top">
       <table border="0" width="100%"  cellspacing="0" cellpadding="0" valign="top" ID="Table2">
        <tr>                    
        <td  align="left" nowrap valign="top" >
		<!--#include file="../include/top_bar_links.asp"-->
		</td>
			<td width="100%" align="left" class="bar_top" valign="baseline" style="padding-left:35px;">
			<table border="0" width="100%" height="15" cellspacing="0" cellpadding="0" ID="Table6">
			<tr>
				<td width="100%" height="2"></td>
			</tr>
			<tr>
			<td width="100%" bgcolor="#DDDDDD" height="13">
			<%''/strat of news%>
			<marquee direction=left  scrolldelay=120>
		
				<%
				sqlStr="SELECT new_ID,New_Title,New_Date FROM News"
				sqlStr=sqlStr& " WHERE New_Home_Visible=1 and Category_Id=2" 
				sqlStr=sqlStr& " ORDER BY New_Date DESC,New_ID DESC" 
				set newsList=con.EXECUTE(sqlStr)
				do while not newsList.EOF 
				newID=newsList("New_ID")
				newTitle=newsList("New_Title")
				newDate=newsList("New_Date")
				%>
				&nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;
				<a class="homeNews" href="<%=Application("VirDir")%>/english/news/news.asp?ID=<%=newID%>"> >> <%=newTitle%></a>
				
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
      </td>
    </tr>
  </table>
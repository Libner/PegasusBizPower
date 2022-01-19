<%Session.LCID = 2057%>
<%SERVER.ScriptTimeout=3000%>
<%Response.Buffer = False %>
<%facultyId=Request.QueryString("catID")
  Id=request.querystring("Id")
  catID=Request.QueryString("catID")
%>
<!--#include file="../../include/connect.asp"-->
<!--#include file="../../include/reverse.asp"-->
<!--#INCLUDE FILE="../checkAuWorker.asp"-->
<%

if catid<>nil and catid<>"" then
	set rs_category = con.execute("SELECT Category_Name FROM News_Categories WHERE Category_ID=" &catid)
	category_Name = rs_Category("Category_Name")
	rs_category.close
	set rs_category=Nothing
end if

if Request.form("title")<>"" then 'after form filling
    
	 dd=request("pMonth")&"/"&request("pDay")&"/"&request("pYear")
     dateflag=0
     if not (IsDate(dd)) then
		dateflag=1
     end if
     if dateflag=1 then
		dd=Now()
		myAdDay=Day(dd)
		myAdMonth=Month(dd)
		myAdYear=Year(dd)
		dd=myAdMonth & "/" & myAdDay & "/" & myAdYear 
     end if
     new_Time_on = sFix(request.form("new_Time_on"))
     if trim(new_Time_on) = "" then
		new_Time_on = "getdate()"
	 else
		if IsDate(new_Time_on) then
			new_Time_on_DATE =  left(new_Time_on,10)
			new_Time_on_DAY = Day(new_Time_on_DATE)
			new_Time_on_MONTH = MONTH(new_Time_on_DATE)
			new_Time_on_YEAR = YEAR(new_Time_on_DATE)
			new_Time_on_NEW_DATE = new_Time_on_MONTH & "/" & new_Time_on_DAY & "/" & new_Time_on_YEAR
			new_Time_on = Replace(new_Time_on,new_Time_on_DATE,new_Time_on_NEW_DATE)
			new_Time_on = "'"& new_Time_on &"'"	
		else
			new_Time_on = "getdate()"
		end if	
	 end if
	 
	 new_Time_off = sFix(request.form("new_Time_off"))
	 if trim(new_Time_off) = "" then
		new_Time_off = "DATEADD(d, 1, getdate())"
	 else
		if IsDate(new_Time_off) then		
			new_Time_off_DATE =  left(new_Time_off,10)
			new_Time_off_DAY = Day(new_Time_off_DATE)
			new_Time_off_MONTH = MONTH(new_Time_off_DATE)
			new_Time_off_YEAR = YEAR(new_Time_off_DATE)
			new_Time_off_NEW_DATE = new_Time_off_MONTH & "/" & new_Time_off_DAY & "/" & new_Time_off_YEAR
			new_Time_off = Replace(new_Time_off,new_Time_off_DATE,new_Time_off_NEW_DATE)
	 		new_Time_off = "'"& new_Time_off &"'"	
		else
			new_Time_off = "DATEADD(d, 1, getdate())"
		end if	 
	 end if
	 title=sFix(trim(Request.form("title")))
	 desc=sFix(trim(Request.form("desc")))
	 content=sFix(trim(Request.Form("content")))
	
	 if id=nil or id="" then 'new record in DataBase
		con.EXECUTE("UPDATE news SET new_Order=new_Order+1 WHERE Category_ID="& catID &"")
		sqlstring="insert into News (category_ID,New_Order,New_Title,New_Content,New_Desc,New_Date,New_Time_on,New_Time_off,New_Site_visible,New_Home_visible,Page_Width) "
		sqlstring=sqlstring& " values ("& catID &",0,'" & title & "'," & Chr(39) & content & Chr(39) & ",'" & desc & "','" & dd & "',"& new_Time_on &","& new_Time_off &",0,0,'"& sFix(Request.Form("PageWidth")) &"')"
        'Response.Write sqlstring
		con.execute(sqlstring)
		set pr=con.execute("select TOP 1 New_ID from News ORDER BY New_ID DESC")
		Id=CStr(pr("New_ID"))
	 else	'update existing record
        s="update News set New_Title='" & title & "',New_Content="& Chr(39) & content & Chr(39) &",New_Desc='" & desc & "',New_Date='"& dd & "',New_Time_on="& new_Time_on &",New_Time_off="& new_Time_off &",Page_Width='"& sFix(Request.Form("PageWidth")) &"' where New_ID=" & Id
	    'Response.Write(s)
	    'Response.End
		con.execute (s)
     end if 
end if

%>
<html>
<head>
<title>Administration Site</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1255">
<link href="../dynamic_style.css" rel="STYLESHEET" type="text/css">
<!-- European format dd-mm-yyyy -->
<script language="JavaScript" src="calendar1.js"></script><!-- Date only with year scrolling -->
<!-- American format mm/dd/yyyy -->
<script language="JavaScript" src="calendar2.js"></script><!-- Date only with year scrolling -->
<script LANGUAGE="JavaScript">
<!--

function CheckDel() {
  return (confirm("?האם ברצונך למחוק את התמונה"))    
}

function CheckFields()
{
	//window.alert(window.document.all("content").value);
	//return false;
	
	var doc=window.document;
	if (doc.forms['frmMain'].elements['title'].value==''){
		 alert("'חובה למלא את השדה 'כותרת");
	     return false;
	}else
		{
		document.frmMain.submit();
		 return true;
		} 
}

function ifFieldEmpty(){
	  if (document.form1.UploadFile1.value=='')
		{
		alert('בחר תמונה');
		return false;
		}
		else if  ( (document.form1.UploadFile1.value.search("jpg") == -1) && (document.form1.UploadFile1.value.search("gif") == -1) && (document.form1.UploadFile1.value.search("jpeg") == -1) && (document.form1.UploadFile1.value.search("JPG") == -1) && (document.form1.UploadFile1.value.search("GIF") == -1) && (document.form1.UploadFile1.value.search("JPEG") == -1) ){
					window.alert("!יתקבלו jpg, jpeg, gif רק תמונות מסוג ");
					return false;
			}		
	  else
		return true;
}

//-->
</script>  
<script language="Javascript1.2">
<!-- // load htmlarea
_editor_url = "../../htmlareaTav/";                     // URL to htmlareaTav files
_editor_lang = "en";
var page_id = '<%=ID%>';
_view_form = false;	
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
document.write('<scr' + 'ipt src="' +_editor_url+ 'editor.js"');
document.write(' language="javascript"></scr' + 'ipt>');  
} else { document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>'); }

   function SetHTMLArea()
   {
		var config = new Object();    // create new config object
		config.bodyStyle = 'font-family: Arial; font-size: 12pt; MARGIN: 0px; PADDING-LEFT: 0px; PADDING-RIGHT: 0px; SCROLLBAR-FACE-COLOR: #DDECFF;SCROLLBAR-HIGHLIGHT-COLOR: #DEEBFE;SCROLLBAR-SHADOW-COLOR: #719BED;SCROLLBAR-3DLIGHT-COLOR: #DDECFF;SCROLLBAR-ARROW-COLOR: #11449D;SCROLLBAR-TRACK-COLOR: #DDECFF;SCROLLBAR-DARKSHADOW-COLOR: #ffffff';
		config.debug = 0;
		editor_generate("content", config);
  }	
// -->
</script>
</head>
<body class="body_admin"  onload="SetHTMLArea()">
<table border="0" width="100%" align="center"  cellspacing="0" cellpadding="2"> 
<tr><td class="a_title_big" width="100%" dir=rtl><font size="4" color="#ffffff"><strong>"<%=category_Name%>" דף ניהול <%if trim(ID) <> "" then%>עדכון<%else%>הוספה<%end if%> אירוע</strong></font></td></tr>
<tr><td height="10"></td></tr>
<FORM name="frmMain" ACTION="Aparaadd.asp?catID=<%=catID%>&Id=<%=Id%>" METHOD="post" onSubmit="return CheckFields()" ID="Form3">  <tr>
<tr>
    <td>
       <table class="table_admin_1" border="0" width="100%" align="center"  cellspacing="1" cellpadding="2">
         <tr>
              <td class="td_admin_5" align="left" ><input type=submit value="עדכון" class="button_admin_1"></td>  
             <td class="td_admin_5" align="left" nowrap><a class="button_admin_1" href="admin.asp?catID=<%=catID%>">חזרה לרשימת אירועים</a></td>
             <%if trim(ID)<>"" then%>
                <td class="td_title_2" align="left" valign="middle" nowrap><a target="_blank" class="button_admin_1" href="../../news/news.asp?Id=<%=Id%>&showElement=True">תצוגה מקדימה באתר</a></td>
	         <%end if%>
             <td class="td_admin_5" align="left" nowrap><a class="button_admin_1" href="../choose.asp">חזרה לדף ניהול ראשי</a></td>	         
             <td class="td_title_2" align="right" valign="middle" width="100%"></td>
         </tr>
       </table>
    </td>
  </tr>
 <tr><td height="5"></td></tr>
<tr><td>
<table align=center border="0" cellpadding="1" cellspacing="1" width="100%">
  <tr>
    <td colspan="2">
       <table align=center border="0" cellpadding="0" cellspacing="0" width="100%">
          <tr>
             <td class="td_title_2" colspan="2" align="left"></td>
          </tr>
       </table>
    </td>
  </tr>
<%Page_Width=PageWidth
if id<>nil and id<>"" then
  set pr=con.execute("select New_Title,Page_Width,New_Date,new_Time_on,new_Time_off,New_Content,New_Desc,New_Picture from News where New_Id=" & Id)
  perContent=(pr("New_Content")) 
  perTitle=pr("New_Title")	 
  Page_Width=pr("Page_Width")
  perDescription=pr("New_Desc") 
  dt=pr("New_Date")
  new_Time_on = pr("new_Time_on")
  new_Time_off = pr("new_Time_off")	
  New_Picture = pr("New_Picture")	  
  else 
  perContent="<p dir=rtl>&nbsp;</p>"
end if
if trim(Page_Width)="" OR isNull(Page_Width) then
  Page_Width=PageWidth
end if%>
<tr>
  <td width="86%" class="td_admin_4" align=right><input dir="rtl" type="text" name="title" style="font-family:Arial;" style="width:400px;" value="<%=vFix(perTitle)%>">
  <td width="14%" class="td_admin_4" align="center">כותרת</td>
</tr>
<tr>
<%if dt<>"" then
	ddt=day(dt)
	mdt=month(dt)
	ydt=year(dt)
	hdt=Hour(dt)
	medt=Minute(dt)
	if hdt=0 and medt=0 then
	  hdt=""
	  medt=""	
	end if	
else
	dd=Now()
	ddt=Day(dd)
	mdt=Month(dd)
	ydt=Year(dd)
	hdt=""
	medt=""	
end if%>
<td class="td_admin_4_right_align" align="center" nowrap>
		<font face="Arial (Hebrew)">		        
			 (יום / חודש / שנה)&nbsp;
			<input type="text" name="pDay" value="<%=ddt%>" size="2" maxlength="2" >/
			<input type="text" name="pMonth" value="<%=mdt%>" size="2" maxlength="2">/
    		<input type="text" name="pYear" value="<%=ydt%>" size="4" maxlength="4">
		</font>
</td>
<td class="td_admin_4"  align="center" nowrap>תאריך</td>
</tr>
<%if Trim(new_Time_on) = "" then
		sqlStr = "select getdate() as new_Time_on, DATEADD(d, 1, getdate()) as new_Time_off"
		set rs_current_Times = con.execute(sqlStr)
		if not rs_current_Times.eof then
			new_Time_on = rs_current_Times("new_Time_on")
			new_Time_on = Left(new_Time_on,len(new_Time_on)-2) & "00"
			
			new_Time_off = rs_current_Times("new_Time_off")	
			new_Time_off = Left(new_Time_off,len(new_Time_off)-2) & "00"			
		end if
		set rs_current_Times = nothing
	end if
%>

<%if 1 then%>
<input type=hidden name="new_Time_on" ID="new_Time_on" value="">
<input type=hidden name="new_Time_off" ID="new_Time_off" value="">
<%else%>
<tr>
	<td align=right class="td_admin_4">
		<a href="javascript:new_Time_on.popup();"><img src="img/cal.gif" width="16" height="16" border="0"></a>
		<input type=text dir="ltr" name="new_Time_on" value="<%=vfix(new_Time_on)%>" size="20" style="width:150px;font-family:arial" onblur="javascript: auto_news_flash();">
	</td>
	<td align="center" class="td_admin_4" nowrap><font style="color:green;"><b>שעת הופעה</b></font></td>
</tr>
<tr>
	<td align=right class="td_admin_4">
		<a href="javascript:new_Time_off.popup();"><img src="img/cal.gif" width="16" height="16" border="0" alt="Click Here to Pick up the date"></a>
		<input type=text dir="ltr" name="new_Time_off" value="<%=vfix(new_Time_off)%>" size="20" style="width:150px;font-family:arial">
	</td>
	<td align="center" class="td_admin_4" nowrap><font style="color:red;"><b>שעת הורדה</b></font></td>
</tr>
<%end if%>

<tr>
   <td class="td_admin_4" align="right"><textarea dir="rtl" name="desc" rows="3" style="width:400px;" style="font-family:Arial;" ><%=perDescription%></textarea></td>
   <td class="td_admin_4" align="center" nowrap>תקציר</td>
</tr>
<tr>
	<td align="right" class="td_admin_4_right_align">px&nbsp;<input type=text  name="PageWidth" value="<%=Page_Width%>" size=5 maxlength=4></td>
	<td align="center" class="td_admin_4">רוחב דף</td>
</tr>				
<tr>
   <td align=right bgcolor=#e6e6e6>
	   <textarea name="content" id="content" dir="rtl" style="width:<%=Page_Width%>; height:360px"><%=perContent%></textarea>
	</td>
    <td class="td_admin_4"  align="center" nowrap>תוכן</td>
</tr>
<tr><td class="td_admin_5" colspan="2" align="left"><input type=submit value="עדכון" class="button_admin_1"></tr>
<tr><td colspan="2" class="td_line_between"></td></tr>
</form>
</table></td></tr>
<tr><td colspan="2">&nbsp;</td></tr>

<%if 0 and id <> nil and id <> "" then%>
	<%'//start logo image file%>
		<tr><td bgcolor="#DDDDDD" align="center">
		    <%if New_Picture<>nil then %>
		          <img src="../../download/news/<%=New_Picture%>">
		    <%else%>
		       &nbsp;
		    <%end if%>
		 </td>
		<td align=center bgcolor="#DDDDDD" nowrap>תמונה</td>
		</tr>

		<FORM id="form1" name="form1" ACTION="imageAdd.asp?C=1&catID=<%=catID%>&New_id=<%=id%>" METHOD="post" ENCTYPE="multipart/form-data">
		<tr>
		<td align="center" bgcolor="#DDDDDD"><font face="Arial (Hebrew)"><INPUT TYPE="FILE" NAME="UploadFile1" SIZE=45 ID="File1"></td>
		<td align="center" bgcolor="#DDDDDD">תמונה להורדה</td>
		</tr>
		<tr><td colspan="2" align="center" bgcolor="#DDDDDD"><input type=submit value="העלאת תמונה" style="font-family:arial" onClick="return ifFieldEmpty();" ID="Submit1" NAME="Submit1"></td></tr>
		</form>

		<%if New_Picture<>"" and New_Picture<>nil then %>
		   <FORM ACTION="imageAdd.asp?C=0&catID=<%=catID%>&New_id=<%=id%>" METHOD="POST" id=form2 name=form2>
		     <tr><td colspan="2" align="center" bgcolor="#DDDDDD"><font face="Arial (Hebrew)"><input type=submit value="מחיקת תמונה" ONCLICK="return CheckDel()" ID="Submit2" NAME="Submit2"></input></td></tr>
		   </form>
		<%end if%>	
	<%'//end of logo image file%>
<tr><td colspan="2" bgcolor="#ffffff" height="1"></td></tr>
<tr><td colspan="2" bgcolor="#808080" height="1"></td></tr>
<%END IF%>
</table>
</div>
		<script language="JavaScript">
		<!-- // create calendar object(s) just after form tag closed
			 // specify form element as the only parameter (document.forms['formname'].elements['inputname']);
			 // note: you can have as many calendar objects as you need for your application
			var new_Time_on = new calendar2(document.forms['frmMain'].elements['new_Time_on']);
			new_Time_on.year_scroll = false;
			new_Time_on.time_comp = true;
			
			var new_Time_off = new calendar2(document.forms['frmMain'].elements['new_Time_off']);
			new_Time_off.year_scroll = false;
			new_Time_off.time_comp = true;			
		//-->
		</script>
</body>
</html>
<%
con.close
set con = nothing
%>
